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


async def rerun(ws, query: str) -> float:
    msg = BackMsg()
    msg.rerun_script.query_string = query
    t = time.perf_counter()
    await ws.write_message(msg.SerializeToString(), binary=True)
    while True:
        raw = await ws.read_message()
        if raw is None:
            raise RuntimeError("server closed the connection")
        fm = ForwardMsg()
        fm.ParseFromString(raw)
        if fm.WhichOneof("type") == "script_finished":
            return time.perf_counter() - t


async def student(port: int, i: int, clicks: int, think: float, lat: list, start_gate):
    ws = await websocket_connect(f"ws://127.0.0.1:{port}/_stcore/stream",
                                 subprotocols=["streamlit"])
    await start_gate.wait()
    await asyncio.sleep(random.uniform(0, think))         # students don't click in lockstep
    for k in range(clicks):
        step = STEPS[(i + k) % len(STEPS)]
        lat.append(await rerun(ws, f"sid=S{i}&step={step}"))
        await asyncio.sleep(random.uniform(0.5, 1.5) * think)
    ws.close()


async def main(students: int, clicks: int, think: float, port: int = 0):
    if port:                                   # a server you started yourself (profiling)
        lat: list = []
        gate = _opened()
        t0 = time.perf_counter()
        await asyncio.gather(*(student(port, i, clicks, think, lat, gate)
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
        tasks = [asyncio.create_task(student(port, i, clicks, think, lat, gate))
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
    ap.add_argument("--port", type=int, default=0,
                    help="use a server already running on this port")
    a = ap.parse_args()
    asyncio.run(main(a.students, a.clicks, a.think, a.port))
