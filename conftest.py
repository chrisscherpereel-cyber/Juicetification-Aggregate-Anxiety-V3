# Makes the repository root importable for the test-suite (pytest rootdir conftest).

# tools/ holds scripts (load_test.py is a load generator, not a test), so don't collect it.
collect_ignore = ["tools"]
