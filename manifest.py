# manifest.py — parameter schema for Aggregate Anxiety V3 (Excel edition).
# The Director reads this (via ?manifest=1) to render an instructor config form.
# Every parameter has a default, so an app opened with no config behaves as the
# standalone practice version.
APP_KEY = "app"; NAME = "Aggregate Anxiety"; SCHEMA_VERSION = 1
MODEL_VERSION = "5.1.0"   # bump when the calculation model or the workbook layout changes

_PRESETS = ["standard", "strong_seasonal", "stability_favored", "limited_flex", "tight_storage"]

MANIFEST = {"app_key": APP_KEY, "name": NAME, "schema_version": SCHEMA_VERSION,
            "model_version": MODEL_VERSION, "params": {
  "beginning_inventory": {"type": "int", "default": 2400, "min": 0, "group": "Inventory", "label": "Beginning inventory"},
  "beginning_backorders": {"type": "int", "default": 0, "min": 0, "group": "Inventory", "label": "Beginning backorders (orders still unfilled before January)"},
  "safety_stock": {"type": "int", "default": 2400, "min": 0, "group": "Inventory", "label": "Safety stock"},
  "safety_stock_policy": {"type": "str", "default": "none", "choices": ["none", "maintain"], "group": "Inventory", "label": "Inventory target: none (end each chase month at 0) or maintain (keep safety stock on hand)"},
  "worker_policy": {"type": "str", "default": "whole", "choices": ["whole", "partial"], "group": "Capacity", "label": "Workforce: whole (whole workers only; round up and pay for idle capacity) or partial (fractional workers allowed)"},
  "max_inventory": {"type": "int", "default": 30000, "min": 0, "group": "Inventory", "label": "Max inventory (storage capacity)"},
  "bottles_per_worker": {"type": "int", "default": 1000, "min": 1, "group": "Capacity", "label": "Bottles/worker/month (must equal rate x hours x days; derived if not)"},
  "bottles_per_hour": {"type": "int", "default": 5, "min": 1, "group": "Capacity", "label": "Bottles/worker-hour"},
  "hours_per_day": {"type": "int", "default": 8, "min": 1, "max": 24, "group": "Capacity", "label": "Paid hours/day per worker"},
  "working_days": {"type": "int", "default": 25, "min": 1, "max": 31, "group": "Capacity", "label": "Working days/month"},
  "starting_workforce": {"type": "int", "default": 8, "min": 0, "group": "Capacity", "label": "Starting workforce"},
  "regular_labor_cost": {"type": "int", "default": 3200, "min": 0, "group": "Costs", "label": "Regular labor $/worker/mo"},
  "hiring_cost": {"type": "int", "default": 600, "min": 0, "group": "Costs", "label": "Hiring $/worker"},
  "layoff_cost": {"type": "int", "default": 900, "min": 0, "group": "Costs", "label": "Layoff $/worker"},
  "holding_cost": {"type": "float", "default": 0.25, "min": 0, "group": "Costs", "label": "Holding $/bottle/mo"},
  "backorder_cost": {"type": "float", "default": 1.50, "min": 0, "group": "Costs", "label": "Backorder $/bottle/mo"},
  "overtime_pct": {"type": "float", "default": 0.20, "min": 0, "max": 1, "group": "Costs", "label": "Overtime cap (fraction of regular capacity)"},
  "overtime_cost": {"type": "float", "default": 4.50, "min": 0, "group": "Costs", "label": "Overtime $ per bottle made on overtime (the full cost of an overtime bottle)"},
  "subcontract_cost": {"type": "float", "default": 5.25, "min": 0, "group": "Costs", "label": "Subcontract $/bottle"},
  "subcontract_capacity": {"type": "int", "default": 5000, "min": 0, "group": "Capacity", "label": "Subcontract capacity (bottles/month)"},
  "service_requirement_pct": {"type": "float", "default": 95.0, "min": 0, "max": 100, "group": "Policy", "label": "Required on-time fulfillment % for a feasible hybrid plan (0 = none)"},
  "terminal_backlog_max": {"type": "int", "default": 0, "min": 0, "group": "Policy", "label": "Max backlog allowed at end of December (hybrid)"},
  "scenario_preset": {"type": "str", "default": "standard", "choices": _PRESETS, "group": "Scenario", "label": "Scenario preset (changes only parameters left at default)"},
  "student_sets_policies": {"type": "bool", "default": True, "group": "Pedagogy", "label": "Let students choose the safety-stock and workforce policies on step 1 (off = use the values above)"},
  "hybrid_mode": {"type": "str", "default": "guided", "choices": ["guided", "build", "off"], "group": "Pedagogy", "label": "Hybrid plan: guided (starting plan + questions), build (students build it themselves) or off (chase and level only); students can switch in the sidebar"},
  "include_hybrid": {"type": "bool", "default": True, "group": "Pedagogy", "label": "Include the Hybrid Plan piece (students design and submit a hybrid strategy)"},
  "require_formulas": {"type": "bool", "default": True, "group": "Pedagogy", "label": "Calculated cells must contain Excel formulas (a typed number is not accepted)"},
  "min_justification_words": {"type": "int", "default": 50, "min": 0, "max": 1000, "group": "Pedagogy", "label": "Minimum words in the recommendation justification"},
  "assignment_name": {"type": "str", "default": "Aggregate Anxiety: Aggregate Planning in Excel", "group": "Report", "label": "Assignment name shown in the workbook and PDF"},
  "report_timezone": {"type": "str", "default": "", "group": "Report", "label": "Report timezone (IANA name, e.g. America/Phoenix; blank = server time)"},
  "forecast_demand": {"type": "list", "default": [], "group": "Demand", "label": "12-month demand (blank = random per student)"},
}}
