#!/usr/bin/env python
"""Real-server load test: N simulated browsers on one running Streamlit server.

  python tools/load_test.py --students 30 --clicks 20

Starts `streamlit run tools/load_app.py` on a free port, opens N websocket sessions the way a
browser does, and has every student click (rerun the page) repeatedly with a short think
time, cycling through the steps. Reports click latency (what a student waits from click
to finished page) and the server process's CPU and memory, sampled with psutil.

Everything goes through Streamlit's real server: websocket handling, script runner
threads, protobuf serialisation and the app's own code."""

from __future__ import annotations

import argparse
import asyncio
import os
import random
import socket
import statistics
import subprocess
import sys
import time

import psutil
from streamlit.proto.BackMsg_pb2 import BackMsg
from streamlit.proto.ForwardMsg_pb2 import ForwardMsg
from tornado.websocket import websocket_connect

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STEPS = ["level", "chase", "level", "start", "level", "report"]


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


ERRORS: list = []          # (query, first line of the exception) for every page that failed


SESSION: dict = {}         # ws -> {"sid": session id, "uploader": widget id of the uploader}
EARLY = ForwardMsg.ScriptFinishedStatus.Value("FINISHED_EARLY_FOR_RERUN")


async def rerun(ws, query: str, widgets=None) -> float:
    """One click: rerun the page and wait for the run that finishes it. Returns seconds."""
    msg = BackMsg()
    msg.rerun_script.query_string = query
    if widgets is not None:
        msg.rerun_script.widget_states.CopyFrom(widgets)
    t = time.perf_counter()
    await ws.write_message(msg.SerializeToString(), binary=True)
    info = SESSION.setdefault(id(ws), {})
    while True:
        raw = await ws.read_message()
        if raw is None:
            raise RuntimeError("server closed the connection")
        fm = ForwardMsg()
        fm.ParseFromString(raw)
        kind = fm.WhichOneof("type")
        if kind == "new_session":
            info["sid"] = fm.new_session.initialize.session_id
        if kind == "delta" and fm.delta.WhichOneof("type") == "new_element" \
                and fm.delta.new_element.WhichOneof("type") == "file_uploader":
            info["uploader"] = fm.delta.new_element.file_uploader.id
        if kind == "delta" and fm.delta.WhichOneof("type") == "new_element" \
                and fm.delta.new_element.WhichOneof("type") == "exception":
            e = fm.delta.new_element.exception
            ERRORS.append((query, f"{e.type}: {e.message}"[:160]))
        if kind == "script_finished" and fm.script_finished != EARLY:
            return time.perf_counter() - t


UPLOADS: list = []         # (transfer seconds, check-and-show seconds) per upload


def unique_workbook(base: bytes, tag: str) -> bytes:
    """The answer-key workbook with one harmless extra zip member, so every upload is a
    file the server has never seen (no parse-cache hits), as in a real class."""
    import io
    import zipfile
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(base)) as zin, zipfile.ZipFile(out, "w",
                                                                  zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            z.writestr(item, zin.read(item.filename))
        z.writestr(f"docProps/loadtest-{tag}.txt", tag)
    return out.getvalue()


async def upload(ws, port: int, query: str, data: bytes, name: str) -> float:
    """Ask for an upload URL over the websocket, PUT the file over HTTP, then rerun with the
    uploader's new state — the browser's own sequence. Returns total seconds."""
    from streamlit.proto.WidgetStates_pb2 import WidgetStates
    from tornado.httpclient import AsyncHTTPClient, HTTPRequest
    info = SESSION[id(ws)]
    t0 = time.perf_counter()
    req = BackMsg()
    req.file_urls_request.request_id = f"r{random.random()}"
    req.file_urls_request.file_names.append(name)
    req.file_urls_request.session_id = info["sid"]
    await ws.write_message(req.SerializeToString(), binary=True)
    while True:
        fm = ForwardMsg()
        fm.ParseFromString(await ws.read_message())
        if fm.WhichOneof("type") == "file_urls_response":
            urls = fm.file_urls_response.file_urls[0]
            break
    boundary = "----loadtest%08x" % random.getrandbits(32)
    body = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; "
            f"filename=\"{name}\"\r\nContent-Type: application/vnd.openxmlformats-"
            f"officedocument.spreadsheetml.sheet\r\n\r\n").encode() + data \
        + f"\r\n--{boundary}--\r\n".encode()
    url = urls.upload_url if urls.upload_url.startswith("http") \
        else f"http://127.0.0.1:{port}{urls.upload_url}"
    await AsyncHTTPClient().fetch(HTTPRequest(
        url, method="PUT", body=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}))
    t1 = time.perf_counter()
    ws_state = WidgetStates()
    w = ws_state.widgets.add()
    w.id = info["uploader"]
    f = w.file_uploader_state_value.uploaded_file_info.add()
    f.name, f.size, f.file_id = name, len(data), urls.file_id
    f.file_urls.CopyFrom(urls)
    await rerun(ws, query, ws_state)
    t2 = time.perf_counter()
    UPLOADS.append((t1 - t0, t2 - t1))
    return t2 - t0


KEY = open(os.path.join(ROOT, "tests", "fixtures", "key_seed7.xlsx"), "rb").read()


async def student(port: int, i: int, clicks: int, think: float, lat: list, start_gate,
                  upload_every: int = 0):
    """Click through the steps; with upload_every=k, every k-th action uploads a new workbook
    on the level step instead (then the student is S1 on scenario 7, the fixture's owner)."""
    ws = await websocket_connect(f"ws://127.0.0.1:{port}/_stcore/stream",
                                 subprotocols=["streamlit"])
    who = "sid=S1&seed=7&u=" + str(i) if upload_every else f"sid=S{i}"
    await start_gate.wait()
    await asyncio.sleep(random.uniform(0, think))         # students don't click in lockstep
    for k in range(clicks):
        if upload_every and k % upload_every == upload_every - 1:
            q = f"{who}&step=level"
            await rerun(ws, q)                             # on the level step, uploader shown
            lat.append(await upload(ws, port, q, unique_workbook(KEY, f"{i}-{k}"),
                                    f"workbook_{i}_{k}.xlsx"))
        else:
            step = STEPS[(i + k) % len(STEPS)]
            lat.append(await rerun(ws, f"{who}&step={step}"))
        await asyncio.sleep(random.uniform(0.5, 1.5) * think)
    ws.close()


async def main(students: int, clicks: int, think: float, port: int = 0,
               upload_every: int = 0):
    if port:                                   # a server you started yourself (profiling)
        lat: list = []
        gate = _opened()
        t0 = time.perf_counter()
        await asyncio.gather(*(student(port, i, clicks, think, lat, gate, upload_every)
                               for i in range(students)))
        lat.sort()
        print(f"{len(lat)} page loads in {time.perf_counter() - t0:.1f} s; median "
              f"{statistics.median(lat) * 1000:.0f} ms, p95 {lat[int(len(lat) * .95)] * 1000:.0f} ms")
        return
    port = free_port()
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    server = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", os.path.join("tools", "load_app.py"),
         "--server.port", str(port), "--server.headless", "true",
         "--server.enableXsrfProtection", "false", "--server.enableCORS", "false"],
        cwd=ROOT, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        for _ in range(120):                                 # wait for the server
            try:
                socket.create_connection(("127.0.0.1", port), timeout=0.5).close()
                break
            except OSError:
                time.sleep(0.25)
        proc = psutil.Process(server.pid)
        procs = [proc] + proc.children(recursive=True)
        warm = []                                            # warm-up: one student, 3 clicks
        await student(port, 999, 3, 0.05, warm, _opened())
        rss0 = sum(p.memory_info().rss for p in procs)
        cpu0 = sum(sum(p.cpu_times()[:2]) for p in procs)
        lat: list = []
        gate = asyncio.Event()
        t0 = time.perf_counter()
        tasks = [asyncio.create_task(student(port, i, clicks, think, lat, gate,
                                             upload_every))
                 for i in range(students)]
        await asyncio.sleep(1.0)                             # all connected, then go
        gate.set()
        peak = rss0
        while not all(t.done() for t in tasks):
            await asyncio.sleep(0.2)
            peak = max(peak, sum(p.memory_info().rss for p in procs))
        for t in tasks:
            t.result()
        wall = time.perf_counter() - t0 - 1.0
        cpu = sum(sum(p.cpu_times()[:2]) for p in procs) - cpu0
        lat.sort()
        n = len(lat)
        print(f"{students} students x {clicks} clicks = {n} page loads in {wall:.1f} s "
              f"(think time ~{think:.1f} s)")
        print(f"click latency: median {statistics.median(lat) * 1000:.0f} ms, "
              f"p95 {lat[int(n * .95)] * 1000:.0f} ms, max {lat[-1] * 1000:.0f} ms")
        print(f"server CPU: {cpu:.1f} s = {cpu / wall * 100:.0f}% of one core "
              f"({cpu / n * 1000:.0f} ms per page load)")
        print(f"server memory: {rss0 / 2**20:.0f} MB with sessions open, "
              f"peak {peak / 2**20:.0f} MB")
        if UPLOADS:
            tr = sorted(u[0] for u in UPLOADS)
            pr = sorted(u[1] for u in UPLOADS)
            print(f"uploads: {len(UPLOADS)}; transfer median {statistics.median(tr) * 1000:.0f} "
                  f"ms (max {tr[-1] * 1000:.0f}); check & show median "
                  f"{statistics.median(pr) * 1000:.0f} ms, p95 "
                  f"{pr[int(len(pr) * .95)] * 1000:.0f} ms, max {pr[-1] * 1000:.0f} ms")
        print(f"errors: {len(ERRORS)} of {n} page loads"
              + (f" — e.g. {ERRORS[0][1]} ({ERRORS[0][0]})" if ERRORS else ""))
    finally:
        server.terminate()
        try:
            server.wait(10)
        except subprocess.TimeoutExpired:
            server.kill()


def _opened():
    e = asyncio.Event()
    e.set()
    return e


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--students", type=int, default=30)
    ap.add_argument("--clicks", type=int, default=20)
    ap.add_argument("--think", type=float, default=2.0,
                    help="average seconds between a student's clicks")
    ap.add_argument("--upload-every", type=int, default=0,
                    help="every k-th action uploads a new workbook (0 = clicks only)")
    ap.add_argument("--port", type=int, default=0,
                    help="use a server already running on this port")
    a = ap.parse_args()
    asyncio.run(main(a.students, a.clicks, a.think, a.port, a.upload_every))
