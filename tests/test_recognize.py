import pytest

from aggplan.recognize import HELPER, REQ_PRODUCTION, classify, map_columns


@pytest.mark.parametrize("heading,concept", [
    ("Beginning Inventory", "Beginning Inventory"), ("Beg Inv", "Beginning Inventory"),
    ("Starting Stock", "Beginning Inventory"), ("Opening inventory", "Beginning Inventory"),
    ("Beginning Backorders", "Beginning Backorders"), ("Backlog Carried In", "Beginning Backorders"),
    ("Production", "Regular Production"), ("Bottles Made", "Regular Production"),
    ("Workers", "Workers"), ("Crew", "Workers"), ("Workers Needed", "Workers"),
    ("Hires", "Hires"), ("New Hires", "Hires"), ("Layoffs", "Layoffs"), ("Fired", "Layoffs"),
    ("Ending Inventory", "Ending Inventory"), ("Closing Stock", "Ending Inventory"),
    ("Inventory", "Ending Inventory"), ("Ending Backorders", "Ending Backorders"),
    ("Unfilled Orders", "Ending Backorders"), ("Labor Cost", "Labor Cost"), ("Wages", "Labor Cost"),
    ("Total Labor Cost", "Labor Cost"), ("Hiring Expense", "Hiring Cost"),
    ("Severance", "Layoff Cost"), ("Storage Cost", "Holding Cost"),
    ("Shortage Penalty", "Backorder Cost"), ("Monthly Total Cost", "Total Cost"),
    ("Total Cost", "Total Cost"), ("Regular Capacity", "Regular Capacity"),
    ("OT Cost", "Overtime Cost"), ("Subcontract Production", "Subcontract Production"),
])
def test_headings(heading, concept):
    assert classify(heading, "chase") == concept


def test_helpers_and_unknowns():
    assert classify("Idle Capacity", "chase") == HELPER
    assert classify("Demand check", "chase") == HELPER
    assert classify("Required Production", "chase") == REQ_PRODUCTION
    assert classify("Stuff", "chase") is None


def test_total_production_depends_on_plan():
    assert classify("Total Production", "hybrid") == "Total Production"
    cm = map_columns({3: "Total Production"}, "chase")
    assert cm.by_concept["Regular Production"] == "C"


def test_required_production_is_used_when_alone():
    cm = map_columns({3: "Required Production"}, "chase")
    assert cm.by_concept["Regular Production"] == "C"
    cm = map_columns({3: "Required Production", 4: "Production"}, "chase")
    assert cm.by_concept["Regular Production"] == "D" and cm.read_as["C"] == HELPER


def test_duplicates_unknowns_and_guesses():
    cm = map_columns({3: "Ending Inventory", 4: "Inventory", 5: "Blah"}, "chase",
                     guess=lambda concept, letter: concept == "Hires" and letter == "E")
    assert cm.duplicates == {"Ending Inventory": ["C", "D"]}
    assert cm.guessed == {"Hires": "E"} and not cm.unknown
    assert "Layoffs" in cm.missing("chase")
