#!/usr/bin/env python3
"""Reproducible acceptance/rejection tests for two valid and two corrupted copies."""
from pathlib import Path
from hashlib import sha256
from shutil import copy2
from validate import validate
ROOT=Path(__file__).resolve().parents[1]
TESTS=Path(__file__).resolve().parent/'tests'
MASTER=ROOT/'cases.xlsx'
for name in ('valid_A.xlsx','valid_B.xlsx'):
    p=TESTS/name
    if not p.exists():copy2(MASTER,p)
def digest(p):return sha256(p.read_bytes()).hexdigest()
master_hash=digest(MASTER)
master_errors,master_result=validate(MASTER)
lines=['LIFELINE ITERATION 1 — VALIDATION REPORT','Group seed: 24',
       f'Master cases.xlsx: {"PASS" if not master_errors else "FAIL"} (independent V1–V8 check).', master_result,
       'Valid-copy tests:']
for name in ('valid_A.xlsx','valid_B.xlsx'):
    ok=digest(TESTS/name)==master_hash
    lines.append(f'{name}: {"PASS" if ok else "FAIL"} — byte-identical copy of the V1–V8-passing master workbook')
lines+=['','Corrupted-copy rejection tests:']
failed=[]
if master_errors: failed.append('master workbook')
for name in ('invalid_A_five_faults.xlsx','invalid_B_five_faults.xlsx'):
    errors,_=validate(TESTS/name)
    rejected=bool(errors)
    lines.append(f'{name}: {"REJECTED" if rejected else "NOT REJECTED"}; {len(errors)} validator finding(s).')
    if errors:lines.append('  '+ ' | '.join(errors[:8]))
    else:failed.append(name)
valid_ok=all('PASS' in x for x in lines if x.startswith('valid_'))
lines+=['',f'Test summary: {2 if valid_ok else 0} valid copies accepted; {2-len(failed)} of 2 five-fault corrupted copies rejected.',
       'TEST SUITE STATUS: '+('PASS' if valid_ok and not failed else 'FAIL')]
(ROOT/'validation_report.txt').write_text('\n'.join(lines)+'\n',encoding='utf-8')
print('\n'.join(lines))
raise SystemExit(0 if valid_ok and not failed else 1)
