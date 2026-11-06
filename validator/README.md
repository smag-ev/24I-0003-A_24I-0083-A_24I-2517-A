# LIFELINE Phase 1 validator

Requires Python 3.10+ and `openpyxl`.

From the project root:

```bash
python validator/validate.py cases.xlsx --report validation_report.txt
```

The validator independently checks the workbook layout and case identity (V1),
tier and field ranges (V2), identifiers and references (V3), network
connectivity and closure waves (V4), radio patterns, edit bounds and tier-S
claims (V5), planted tours/covers/assignments/rations (V6), all six case-study
signatures and cross-case variety (V7), and all four trap constructions (V8).
It prints `PASS` only when no rule fails; otherwise it prints the failing rule
and case.

To regenerate the two valid copies and two five-fault corrupted copies in `validator/tests/`, run:

```bash
python validator/make_tests.py
```

Then run the full acceptance/rejection test suite:

```bash
python validator/test_suite.py
```

The test suite validates the master workbook, confirms that both valid copies
are byte-identical to it, and runs the validator against both corrupted copies.
Each corrupted copy contains five independent faults. The checked-in
`validation_report.txt` records the observed output.
