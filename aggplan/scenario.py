"""Scenario: the immutable bundle of resolved parameters + demand that every calculation
receives explicitly. One student ID -> one reproducible scenario."""

from __future__ import annotations

import hashlib
import json
import random
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from manifest import MANIFEST

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
N = len(MONTHS)

DEFAULTS: Dict[str, object] = {k: s["default"] for k, s in MANIFEST["params"].items()}

# Parameters an instructor preset may change (only if the instructor left them at default).
PRESETS: Dict[str, dict] = {
    "standard": {},
    "strong_seasonal": {"_amplitude": 1.9},
    "stability_favored": {"hiring_cost": 2400, "layoff_cost": 2700},
    "limited_flex": {"overtime_pct": 0.05, "subcontract_capacity": 2000},
    "tight_storage": {"max_inventory": 12000},
}
PRESET_LABELS: Dict[str, str] = {
    "standard": "Standard seasonal demand",
    "strong_seasonal": "Strong seasonal demand (deep trough, tall peak)",
    "stability_favored": "Hiring and layoff costs favor workforce stability",
    "limited_flex": "Limited overtime and subcontracting",
    "tight_storage": "Restricted storage capacity",
}

# Parameters that change presentation only, never a correct answer. They are excluded from
# the config hash, so an instructor can retune them mid-course without invalidating workbooks.
NON_MODEL_PARAMS = frozenset({
    "include_hybrid", "require_formulas", "min_justification_words", "assignment_name",
    "report_timezone",
})

_SHAPE = [0.55, 0.60, 0.75, 1.00, 1.30, 1.55, 1.65, 1.40, 1.10, 0.85, 0.70, 0.55]

# Excel defined names for the given data (Scenario sheet). Readable on purpose: students
# type them into formulas, e.g. =F2*LaborCost.
NAME_KEYS: List[Tuple[str, str, str, str]] = [
    # (Excel name, parameter key, meaning, unit)
    ("BeginInv", "beginning_inventory", "Inventory on hand before January", "bottles"),
    ("SafetyStock", "safety_stock", "Safety-stock quantity", "bottles"),
    ("EndTarget", "_target", "Inventory each chase month aims to end with "
                             "(and the level plan's December target)", "bottles"),
    ("MaxInventory", "max_inventory", "Storage capacity (most bottles you can hold)", "bottles"),
    ("StartWorkers", "starting_workforce", "Workers employed before January", "workers"),
    ("BottlesPerHour", "bottles_per_hour", "Output of one worker per hour", "bottles/hour"),
    ("HoursPerDay", "hours_per_day", "Paid hours per worker per day", "hours/day"),
    ("DaysPerMonth", "working_days", "Working days per month", "days/month"),
    ("LaborCost", "regular_labor_cost", "Regular labor cost per worker per month", "$/worker/month"),
    ("HireCost", "hiring_cost", "Cost to hire one worker", "$/worker"),
    ("LayoffCost", "layoff_cost", "Cost to lay off one worker", "$/worker"),
    ("HoldCost", "holding_cost", "Cost to hold one bottle for one month", "$/bottle/month"),
    ("BackorderCost", "backorder_cost", "Cost of one bottle backordered for one month",
     "$/bottle/month"),
    ("OTLimitPct", "overtime_pct", "Overtime limit as a share of regular capacity", "fraction"),
    ("OTCost", "overtime_cost", "Overtime cost per bottle (extra, on top of labor)", "$/bottle"),
    ("SubCost", "subcontract_cost", "Subcontracting cost per bottle", "$/bottle"),
    ("SubCapacity", "subcontract_capacity", "Most bottles you can subcontract per month",
     "bottles/month"),
]


def stable_int(*parts, lo: int = 0, hi: int = 10 ** 9) -> int:
    """Deterministic integer from arbitrary parts (SHA-256, same in every process)."""
    key = "|".join(str(p) for p in parts).encode("utf-8")
    h = hashlib.sha256(key).digest()
    return lo + int.from_bytes(h[:8], "big") % max(1, hi - lo + 1)


def generate_forecast(seed: int, amplitude: float = 1.0) -> List[int]:
    """A unique but easy-to-calculate 12-month forecast (whole multiples of 500, annual
    total a multiple of 6,000). Same algorithm as V2/V3 so seeds reproduce the same demand."""
    r = random.Random(seed)
    avg_h = r.choice([17, 19, 21, 23, 25, 27])                     # half-thousands / month
    shape = [max(0.35, s + r.uniform(-0.12, 0.12)) for s in _SHAPE]
    if amplitude != 1.0:
        mean = sum(shape) / len(shape)
        shape = [max(0.12, mean + (s - mean) * amplitude) for s in shape]
    tot_h = avg_h * 12
    ssum = sum(shape)
    dh = [max(1, round(s / ssum * tot_h)) for s in shape]
    dh[dh.index(max(dh))] += tot_h - sum(dh)                       # fix rounding drift on peak
    return [int(x) * 500 for x in dh]


@dataclass(frozen=True)
class Scenario:
    params: Dict[str, object]
    demand: Tuple[float, ...]
    seed: int
    preset: str = "standard"
    fixed_demand: bool = False
    notices: Tuple[str, ...] = ()
    months: Tuple[str, ...] = tuple(MONTHS)

    def __getitem__(self, k):
        return self.params[k]

    @property
    def total_demand(self) -> float:
        return float(sum(self.demand))

    @property
    def cap_month(self) -> float:
        """Bottles one worker makes per month (= rate x hours x days)."""
        return float(self.params["bottles_per_worker"])

    @property
    def maintains_safety(self) -> bool:
        return self.params.get("safety_stock_policy") == "maintain"

    @property
    def target(self) -> float:
        """Ending-inventory target used by both plans (0 unless safety stock is maintained)."""
        return float(self.params["safety_stock"]) if self.maintains_safety else 0.0

    def name_value(self, key: str) -> float:
        return self.target if key == "_target" else float(self.params[key])

    def named_values(self) -> Dict[str, float]:
        return {name: self.name_value(key) for name, key, _, _ in NAME_KEYS}

    # ---- identity ------------------------------------------------------------
    @property
    def model_params(self) -> Dict[str, object]:
        return {k: v for k, v in self.params.items() if k not in NON_MODEL_PARAMS}

    @property
    def config_hash(self) -> str:
        """Hash of everything that changes a correct answer (model parameters + demand)."""
        blob = json.dumps({"p": self.model_params, "d": list(self.demand)}, sort_keys=True,
                          separators=(",", ":"))
        return hashlib.sha256(blob.encode()).hexdigest()[:16]


def _num(x):
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    if f != f or f in (float("inf"), float("-inf")):
        return None
    return f


def resolve_scenario(cfg: Optional[dict], seed: int) -> Scenario:
    """Build a Scenario from the Director-resolved params (`cfg`) and a seed.

    Validation never raises: problems become `notices` and a safe fallback is used, so a
    bad config cannot crash the app mid-class."""
    notices: List[str] = []
    params = dict(DEFAULTS)
    for k, v in (cfg or {}).items():
        if k in params:
            params[k] = v

    preset = params.get("scenario_preset") or "standard"
    if preset not in PRESETS:
        notices.append(f"Unknown scenario preset '{preset}' — using the standard scenario.")
        preset = "standard"
    amplitude = 1.0
    for k, v in PRESETS[preset].items():
        if k == "_amplitude":
            amplitude = float(v)
        elif params.get(k) == DEFAULTS[k]:               # never overwrite an instructor value
            params[k] = v
    params["scenario_preset"] = preset

    fixed = params.pop("forecast_demand", []) or []
    fixed_ok = False
    if fixed:
        vals = [_num(x) for x in fixed]
        if len(vals) == N and all(v is not None and v >= 0 for v in vals):
            fixed_ok = True
            fixed = [float(v) for v in vals]
        else:
            notices.append("The configured 12-month demand must contain exactly 12 non-negative "
                           "numbers; it was ignored and random demand is used instead.")

    derived = params["bottles_per_hour"] * params["hours_per_day"] * params["working_days"]
    if params["bottles_per_worker"] != derived:
        notices.append(
            f"Configured bottles/worker/month ({params['bottles_per_worker']:,}) does not equal "
            f"bottles/hour × hours/day × working days ({derived:,}); {derived:,} is used.")
        params["bottles_per_worker"] = derived

    demand = list(fixed) if fixed_ok else generate_forecast(seed, amplitude)
    return Scenario(params=params, demand=tuple(float(d) for d in demand), seed=int(seed),
                    preset=preset, fixed_demand=fixed_ok, notices=tuple(notices))
