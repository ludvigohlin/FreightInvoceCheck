"""One-off correction of 3 historical rows in invoice_header.csv that were
corrupted by a day/month-swap bug in parse_date() (fixed in commit fe1816e,
2026-06-16). These rows were written before the fix and CSVs are append-only,
so they were never retroactively corrected. Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_header.csv"

# (run_id, invoice_number) -> {field: corrected_value}
CORRECTIONS = {
    ("20260602_110755_99a67274", "903108957122"): {"due_date": "2026-06-10"},
    ("20260606_073009_e3e16304", "903111702523"): {"due_date": "2026-07-05"},
    ("20260613_073008_9dba0f06", "903113089226"): {
        "invoice_date": "2026-06-12",
        "due_date": "2026-07-12",
    },
}

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main():
    backup_path = CSV_PATH.with_suffix(".csv.bak")
    shutil.copy2(CSV_PATH, backup_path)
    print(f"Backup written to {backup_path}")

    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        fieldnames = reader.fieldnames
        rows = list(reader)

    changed = []
    remaining = dict(CORRECTIONS)
    for row in rows:
        key = (row.get("run_id", ""), row.get("invoice_number", ""))
        if key in remaining:
            fixes = remaining.pop(key)
            for field, new_value in fixes.items():
                old_value = row.get(field, "")
                if old_value != new_value:
                    row[field] = new_value
                    changed.append((key, field, old_value, new_value))

    if remaining:
        print("WARNING: could not find these rows, no changes made for them:")
        for k in remaining:
            print(f"  {k}")

    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nRows changed: {len(changed)}")
    for key, field, old, new in changed:
        print(f"  {key}  {field}: {old!r} -> {new!r}")
    print(f"\nTotal data rows in file: {len(rows)} (unchanged count expected)")


if __name__ == "__main__":
    main()
