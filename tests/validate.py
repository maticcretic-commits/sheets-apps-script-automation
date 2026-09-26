#!/usr/bin/env python3
"""
Sanity checks for the sheets-apps-script-automation repo:
  1. Every .gs file must pass `node --check` (Apps Script is JavaScript).
  2. CSV templates must parse and carry the exact expected headers.
"""
import csv
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

EXPECTED_HEADERS = {
    "sheets-templates/transactions_template.csv":
        ["Date", "Type", "Category", "Amount", "Note", "CreatedBy"],
    "sheets-templates/expenses_template.csv":
        ["Date", "Category", "Amount", "Currency", "Amount (INR)", "Note"],
}

failures = []

# 1. node --check on every .gs file (copied to .js first: node only checks known extensions)
gs_files = sorted((ROOT / "apps-script").glob("*.gs"))
if not gs_files:
    failures.append("No .gs files found under apps-script/")
for gs in gs_files:
    tmp_js = Path("/tmp") / f"_validate_{gs.stem}.js"
    tmp_js.write_bytes(gs.read_bytes())
    proc = subprocess.run(["node", "--check", str(tmp_js)],
                          capture_output=True, text=True)
    tmp_js.unlink(missing_ok=True)
    if proc.returncode == 0:
        print(f"PASS  node --check  {gs.name}")
    else:
        failures.append(f"node --check failed for {gs.name}:\n{proc.stderr.strip()}")
        print(f"FAIL  node --check  {gs.name}")

# 2. CSV templates parse + headers match
for rel, expected in EXPECTED_HEADERS.items():
    path = ROOT / rel
    try:
        with open(path, newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f))
    except Exception as e:  # noqa: BLE001
        failures.append(f"{rel}: could not parse CSV: {e}")
        print(f"FAIL  parse        {rel}")
        continue
    if not rows:
        failures.append(f"{rel}: empty file")
        print(f"FAIL  empty        {rel}")
        continue
    if rows[0] != expected:
        failures.append(f"{rel}: header mismatch.\n  got:      {rows[0]}\n  expected: {expected}")
        print(f"FAIL  headers      {rel}")
        continue
    data_rows = [r for r in rows[1:] if any(c.strip() for c in r)]
    if len(data_rows) < 5:
        failures.append(f"{rel}: expected >=5 sample rows, found {len(data_rows)}")
        print(f"FAIL  sample rows  {rel}")
        continue
    print(f"PASS  csv            {rel} ({len(data_rows)} sample rows)")

print()
if failures:
    print(f"{len(failures)} FAILURE(S):")
    for f in failures:
        print(" -", f)
    sys.exit(1)
print("All checks passed.")
