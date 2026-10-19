"""Hints that grow with each attempt — guidance, never the answer.

Every column, planning cell and Summary item has a four-step ladder:

  1  a question to think about
  2  the idea: what the value depends on
  3  the Excel approach: which function, which cells, absolute references
  4  the full rule in words (still no numbers: the student computes those)

The app keeps a counter per piece and per item. Each upload in which the item is still
wrong moves its ladder up one step, so a student who is stuck on Ending Inventory gets
more help on Ending Inventory, while a column they just started gets the gentlest nudge.
Missing columns have their own three-step ladder that ends with a suggested heading.
"""

from __future__ import annotations

from typing import Dict, List

from . import layout as L

MAX_LEVEL = 4
MAX_COLUMN_LEVEL = 3
LEVEL_NAMES = {1: "Think about it", 2: "The idea", 3: "In Excel", 4: "The rule"}

# What each column answers, as a question about one month (used for building columns).
QUESTIONS: Dict[str, str] = {
    "Beginning Inventory": "How many bottles are on the shelf when the month starts?",
    "Beginning Backorders": "How many customer orders are still waiting from earlier months?",
    "Regular Production": "How many bottles will you make this month?",
    "Workers": "How many workers do you employ this month?",
    "Hires": "How many workers did you add this month?",
    "Layoffs": "How many workers did you let go this month?",
    "Ending Inventory": "How many bottles are left on the shelf at the end of the month?",
    "Ending Backorders": "How many orders are still unfilled at the end of the month?",
    "Labor Cost": "What did your workforce cost this month?",
    "Hiring Cost": "What did hiring cost this month?",
    "Layoff Cost": "What did layoffs cost this month?",
    "Holding Cost": "What did it cost to keep stock this month?",
    "Backorder Cost": "What did unfilled orders cost you this month?",
    "Total Cost": "What did the whole month cost?",
    "Overtime Production": "How many bottles did you make on overtime?",
    "Subcontract Production": "How many bottles did you buy from a subcontractor?",
    "Total Production": "How many bottles came in altogether this month?",
    "Regular Capacity": "How many bottles could your workforce make on regular time?",
    "Overtime Cost": "What did overtime cost this month?",
    "Subcontract Cost": "What did subcontracting cost this month?",
}


def _production(kind: str) -> List[str]:
    if kind == "chase":
        return [
            "A chase plan makes just what each month needs. What does 'needs' include besides "
            "this month's demand?",
            "What you must make = this month's demand, plus the stock you want left at the end "
            "(your target), minus the stock you already have, plus any orders you still owe.",
            "Combine Demand, your target planning cell (locked with $), Beginning Inventory and "
            "Beginning Backorders in one formula, and wrap it in MAX(0, …) so a month with "
            "plenty of stock makes 0, never a negative number.",
            "Production = MAX(0, Demand + EndTarget − Beginning Inventory + Beginning Backorders).",
        ]
    if kind == "level":
        return [
            "In a level plan, does production change from one month to the next?",
            "No: one steady rate all year, big enough to cover the whole year's demand plus the "
            "stock you want in December, after using the stock you start with.",
            "Work the rate out ONCE in your planning cell, then point every month at that cell "
            "with an absolute reference (for example =$C$20) so it doesn't move when you fill down.",
            "Production every month = the level rate = (total demand + December target − BeginInv) ÷ 12.",
        ]
    return [
        "This is one of your decisions. Is it within what your workforce can make?",
        "Regular production can't be more than workers × capacity per worker.",
        "Type your choice or build it from a formula; check it against Regular Capacity.",
        "Regular Production ≤ Workers × capacity per worker, every month.",
    ]


def _workers(kind: str, whole: bool) -> List[str]:
    if kind == "chase":
        if not whole:
            return [
                "How many workers' worth of capacity does this month's production take?",
                "Divide the bottles you must make by what one worker makes in a month. Your "
                "workforce policy allows fractions, so keep them.",
                "Divide Production by your capacity-per-worker planning cell, locked with $ so "
                "every month uses it. Don't round.",
                "Workers = Production ÷ capacity per worker (fractions allowed).",
            ]
        return [
            "How many people does it take to make this month's production?",
            "Divide the bottles you must make by what one worker makes in a month. Your "
            "workforce policy says whole workers only, so a fraction means one more person.",
            "Use ROUNDUP(…, 0) on Production ÷ capacity per worker, and lock the capacity "
            "planning cell with $ so every month uses it.",
            "Workers = ROUNDUP(Production ÷ capacity per worker, 0).",
        ]
    if kind == "level":
        how = ("ROUNDUP of level rate ÷ capacity per worker" if whole
               else "level rate ÷ capacity per worker, fractions kept")
        return [
            "Does a level plan change its workforce during the year?",
            "No: one workforce, big enough to make the level rate every month.",
            f"Work it out once in the planning cell ({how}) and point every month at it with $.",
            "Workers every month = " + ("ROUNDUP(level rate ÷ capacity per worker, 0)."
                                        if whole else "level rate ÷ capacity per worker."),
        ]
    return [
        "This is one of your decisions. How many workers does your strategy keep?",
        "More workers mean more labor cost every month; fewer may need overtime, "
        "subcontracting or backorders.",
        "Type your choice; the rest of the sheet should react to it.",
        "Workers must be whole numbers." if whole else "Fractional workers are allowed.",
    ]


_COMMON: Dict[str, List[str]] = {
    "Beginning Inventory": [
        "Where does the stock you start a month with come from?",
        "A month starts with whatever the previous month ended with. January starts with "
        "the stock you have before the year begins.",
        "Point at last month's Ending Inventory cell (one row up). January uses the name "
        "BeginInv.",
        "Beginning Inventory = previous month's Ending Inventory; January = BeginInv.",
    ],
    "Beginning Backorders": [
        "Can orders you couldn't fill last month just disappear?",
        "Unfilled orders carry into the next month. January starts with the orders still owed "
        "before the year begins (see the Scenario sheet).",
        "Point at last month's Ending Backorders cell (one row up); January uses the name "
        "BeginBackorders.",
        "Beginning Backorders = previous month's Ending Backorders; January = BeginBackorders.",
    ],
    "Hires": [
        "When do you have to hire?",
        "Compare this month's workers with last month's. Only an increase is a hire. "
        "January compares with the workers you start the year with.",
        "Use MAX(0, …) on the change in workers so a decrease shows as 0 hires. January "
        "uses the name StartWorkers.",
        "Hires = MAX(0, Workers − last month's Workers); January compares with StartWorkers.",
    ],
    "Layoffs": [
        "When do you have to lay people off?",
        "Compare this month's workers with last month's. Only a decrease is a layoff. "
        "January compares with the workers you start the year with.",
        "Use MAX(0, …) on the change the other way round. January uses StartWorkers.",
        "Layoffs = MAX(0, last month's Workers − Workers); January compares with StartWorkers.",
    ],
    "Ending Inventory": [
        "What is left on the shelf after this month's orders go out?",
        "Start with what you had, add what you made, ship the orders you owed from before "
        "and this month's demand. Stock can never be negative: a shortfall is unfilled "
        "orders, not negative bottles.",
        "One formula with Beginning Inventory, Production, Beginning Backorders and Demand, "
        "wrapped in MAX(0, …).",
        "Ending Inventory = MAX(0, Beginning Inventory + Production − Beginning Backorders − Demand).",
    ],
    "Ending Backorders": [
        "If you can't ship everything, what happens to the orders you couldn't fill?",
        "They become backorders: what you owed (old orders plus this month's demand) minus "
        "what you had available (beginning stock plus production), but never below 0.",
        "The same four cells as Ending Inventory, the other way round, wrapped in MAX(0, …).",
        "Ending Backorders = MAX(0, Demand + Beginning Backorders − Beginning Inventory − Production).",
    ],
    "Labor Cost": [
        "What do you pay for regular labor: the bottles made, or the people employed?",
        "Workers are paid for the whole month whether they are busy or not.",
        "Multiply the Workers cell by the name LaborCost.",
        "Labor Cost = Workers × LaborCost.",
    ],
    "Hiring Cost": [
        "What causes a hiring cost?",
        "Each worker you add costs a one-off amount.",
        "Multiply the Hires cell by the name HireCost.",
        "Hiring Cost = Hires × HireCost.",
    ],
    "Layoff Cost": [
        "What causes a layoff cost?",
        "Each worker you let go costs a one-off amount.",
        "Multiply the Layoffs cell by the name LayoffCost.",
        "Layoff Cost = Layoffs × LayoffCost.",
    ],
    "Holding Cost": [
        "Which stock do you pay to keep: what you start with, or what is left over?",
        "You pay to hold the bottles still on the shelf at the end of the month.",
        "Multiply the Ending Inventory cell by the name HoldCost.",
        "Holding Cost = Ending Inventory × HoldCost.",
    ],
    "Backorder Cost": [
        "What does it cost you when customers are kept waiting?",
        "You pay a penalty on every order still unfilled at the end of the month.",
        "Multiply the Ending Backorders cell by the name BackorderCost.",
        "Backorder Cost = Ending Backorders × BackorderCost.",
    ],
    "Overtime Cost": [
        "What do overtime bottles cost?",
        "Every bottle made on overtime costs OTCost. Regular labor pays only for regular-time "
        "capacity, so overtime bottles are charged separately.",
        "Multiply the Overtime Production cell by the name OTCost.",
        "Overtime Cost = Overtime Production × OTCost.",
    ],
    "Subcontract Cost": [
        "What do subcontracted bottles cost?",
        "You pay the subcontractor for every bottle they make.",
        "Multiply the Subcontract Production cell by the name SubCost.",
        "Subcontract Cost = Subcontract Production × SubCost.",
    ],
    "Regular Capacity": [
        "How much could your workers make on regular time?",
        "Every worker can make the same number of bottles per month.",
        "Multiply the Workers cell by your capacity-per-worker planning cell (with $).",
        "Regular Capacity = Workers × capacity per worker.",
    ],
    "Total Production": [
        "Where did all of this month's bottles come from?",
        "Regular time, overtime and the subcontractor together.",
        "Add the three production cells in the row.",
        "Total Production = Regular + Overtime + Subcontract production.",
    ],
    "Overtime Production": [
        "This is one of your decisions. When is overtime allowed?",
        "Only once regular capacity is fully used, and only up to the overtime limit.",
        "Compare it with Regular Capacity and your overtime-limit planning cell.",
        "Overtime ≤ Workers × overtime limit per worker, and only when Regular Production = "
        "Regular Capacity.",
    ],
    "Subcontract Production": [
        "This is one of your decisions. How much can you subcontract?",
        "The subcontractor has a monthly limit.",
        "Compare it with the name SubCapacity.",
        "Subcontract Production ≤ SubCapacity every month.",
    ],
}


def _total_cost(kind: str) -> List[str]:
    parts = ("Labor + Hiring + Layoff + Overtime + Subcontract + Holding + Backorder"
             if kind == "hybrid" else "Labor + Hiring + Layoff + Holding + Backorder")
    return [
        "Which of your cost columns belong in the month's total?",
        "Every cost the month caused — and nothing else (quantities like workers or bottles "
        "are not costs).",
        "Use SUM over that row's cost cells, or add them one by one.",
        f"Total Cost = {parts} (that month's row).",
    ]


PLANNING: Dict[str, List[str]] = {
    "capacity": [
        "How many bottles can ONE worker make in a month?",
        "Bottles per hour, times hours per day, times working days.",
        "Multiply three names from the Scenario sheet.",
        "Capacity per worker = BottlesPerHour × HoursPerDay × DaysPerMonth.",
    ],
    "target": [
        "How much stock should be left at the end?",
        "It depends on your company's inventory policy (see the Start Here and Scenario sheets).",
        "There is a name on the Scenario sheet for exactly this.",
        "Ending inventory target = EndTarget.",
    ],
    "total_demand": [
        "How many bottles do customers want in the whole year?",
        "Add up the twelve months of demand.",
        "SUM the Demand column.",
        "Total demand = SUM of the twelve Demand cells.",
    ],
    "level_rate": [
        "If you make the same amount every month, how much must that be?",
        "Over the year you must make the total demand plus the December stock you want, minus "
        "the stock you already have, plus any orders you already owe; spread it evenly over "
        "12 months.",
        "Use your total-demand and target planning cells and the names BeginInv and "
        "BeginBackorders; divide by 12.",
        "Level production = (total demand + December target − BeginInv + BeginBackorders) ÷ 12.",
    ],
    "level_workers": [
        "How many workers can make the level rate every month?",
        "Divide the level rate by one worker's monthly capacity — and your policy says "
        "whole workers only.",
        "ROUNDUP(…, 0) of your level-rate cell ÷ your capacity cell.",
        "Level workforce = ROUNDUP(level production ÷ capacity per worker, 0).",
    ],
    "ot_limit": [
        "How much overtime may one worker do in a month?",
        "A fixed share of their regular monthly capacity.",
        "Multiply your capacity cell by the name OTLimitPct.",
        "Overtime limit per worker = capacity per worker × OTLimitPct.",
    ],
}

TOTALS = [
    "Have you added up the whole year under this column?",
    "Row 14 holds each column's yearly total.",
    "Use SUM over January–December (rows 2 to 13) of that column.",
    "Row 14 = SUM of rows 2–13 in the same column.",
]

SUMMARY: Dict[str, List[str]] = {
    "link": [
        "Where in your plan is this yearly number?",
        "It's one of the totals you built in row 14 of the plan sheet.",
        "Link to that cell instead of retyping it: type = then click the cell on the plan sheet.",
        "Link to the plan sheet's row-14 total for this cost, e.g. ='Chase Plan'!P14 if "
        "that total is in column P.",
    ],
    "Total Hires": [
        "How many workers did this plan add over the year?",
        "Add up the plan's monthly hires.",
        "Link to a row-14 total, or SUM the plan's Hires column.",
        "Total workers hired = SUM of the plan's Hires column.",
    ],
    "Total Layoffs": [
        "How many workers did this plan let go over the year?",
        "Add up the plan's monthly layoffs.",
        "Link to a row-14 total, or SUM the plan's Layoffs column.",
        "Total laid off = SUM of the plan's Layoffs column.",
    ],
    "Peak Workforce": [
        "What was the biggest workforce in any month?",
        "Look for the largest value in the plan's Workers column.",
        "MAX over that column.",
        "Largest workforce = MAX of the plan's Workers column.",
    ],
    "Peak Inventory": [
        "How much storage does this plan need at most?",
        "Look for the largest month-end stock.",
        "MAX over the plan's Ending Inventory column.",
        "Largest ending inventory = MAX of the plan's Ending Inventory column.",
    ],
    "Months With Backorders": [
        "In how many months were customers left waiting?",
        "Count the months whose Ending Backorders are above zero.",
        'COUNTIF over the Ending Backorders column with the criteria ">0".',
        'Months with backorders = COUNTIF(Ending Backorders column, ">0").',
    ],
    "lower": [
        "Which plan costs less over the year?",
        "Compare the two TOTAL PLAN COST cells.",
        'Let Excel decide with IF: IF(chase cost < level cost, "Chase", "Level").',
        'Lower-cost plan = IF(chase total < level total, "Chase", "Level").',
    ],
    "savings": [
        "How much would the cheaper plan save?",
        "The difference between the two total plan costs, as a positive number.",
        "ABS of one total minus the other.",
        "Savings = ABS(chase total − level total).",
    ],
}


def _level_workers(whole: bool) -> List[str]:
    if whole:
        return PLANNING["level_workers"]
    return [
        "How many workers' worth of capacity does the level rate take?",
        "Divide the level rate by one worker's monthly capacity; your policy allows fractions.",
        "Your level-rate cell ÷ your capacity cell. Don't round.",
        "Level workforce = level production ÷ capacity per worker (fractions allowed).",
    ]


def ladder(kind: str, item: str, whole: bool = True) -> List[str]:
    """The four hint steps for a plan column, planning cell ('plan:<key>'), the totals
    row ('total'), or a Summary item ('summary:<metric>')."""
    if item == "plan:level_workers":
        return _level_workers(whole)
    if item.startswith("plan:"):
        return PLANNING[item[5:]]
    if item == "total" or item.startswith("total:"):
        return TOTALS
    if item.startswith("summary:"):
        key = item[8:]
        return SUMMARY.get(key, SUMMARY["link"])
    if item == "Regular Production":
        return _production(kind)
    if item == "Workers":
        return _workers(kind, whole)
    if item == "Total Cost":
        return _total_cost(kind)
    return _COMMON.get(item, ["Check this value.", "", "", ""])


def hint_text(kind: str, item: str, level: int, whole: bool = True) -> List[str]:
    """Hint steps 1..level (the student sees everything unlocked so far)."""
    steps = ladder(kind, item, whole)
    level = max(1, min(MAX_LEVEL, level))
    return [s for s in steps[:level] if s]


def column_hint(kind: str, concept: str, level: int) -> str:
    """Help for a REQUIRED column the student hasn't built yet (steps 1..3)."""
    q = QUESTIONS.get(concept, concept)
    if level <= 1:
        return f"Your plan can't answer this yet: “{q}”"
    if level == 2:
        return f"Add a column that answers: “{q}”"
    return (f"Add a column headed “{L.heading_for(kind, concept)}” (any clear wording "
            f"works) that answers: “{q}”")


def build_questions(kind: str) -> List[str]:
    """The questions a plan's columns must answer, in a sensible working order."""
    return [QUESTIONS[c] for c in L.REQUIRED[kind]]
