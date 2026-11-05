# LIFELINE — Iteration 1 submission

**Group number / generator seed:** 24  
**Scope:** Iteration 1 only — synthetic dataset generation, independent validation, two-page approach note, and two-page algorithm-selection proposal. The A–I relief-solving implementations belong to Iteration 2 and are intentionally not included.

## Requirements

- Python 3.10 or newer
- `openpyxl`
- `reportlab` is only needed if regenerating the two PDFs from the documentation source (the PDFs are included in this package).

Install workbook dependency:

```bash
python -m pip install openpyxl
```

## Reproduce the required workbook

```bash
python generator/generate.py --seed 24 --output cases.xlsx
```

## Validate the workbook and write the report

```bash
python validator/validate.py cases.xlsx --report validation_report.txt
```

## Rebuild and run validator tests

```bash
python validator/make_tests.py
python validator/test_suite.py
```

The master workbook contains exactly 50 synthetic cases across exactly 20 sheets. The validator implements V1–V8 and the tests include two valid copies and two invalid copies, each corrupted in five independent ways.

## Files

- `cases.xlsx` — generated 50-case master dataset.
- `generator/` — deterministic dataset generator and reproduction instructions.
- `validator/` — independent V1–V8 validator, test-fixture builder, and test runner.
- `validation_report.txt` — actual validator and corruption-test output.
- `approach.pdf` — two-page approach note.
- `proposal.pdf` — two-page algorithm-selection report and work-division table.
- `members.txt` — group member details.
- `repository.txt` — repository URL.
