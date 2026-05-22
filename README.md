# Freight Invoice Control — Python MVP

Automated freight invoice reconciliation and cost analysis for Bring and PostNord invoices.

## Purpose

This script processes freight invoices and carrier specifications from the `00_Inbox` folder,
classifies them, extracts structured data, validates totals, detects anomalies, and outputs
dashboard-ready CSV files and Markdown management summaries.

## Folder Structure

```
Freight_Invoice_Control/
  main.py                  ← Entry point
  requirements.txt
  .env.example             ← Copy to .env and configure
  README.md
  src/                     ← Python modules
  00_Inbox/                ← Drop files here (PDF, XLSX)
  01_Raw/                  ← Archive of processed files (if MOVE_FILES=true)
  02_Processing/           ← Reserved for future use
  03_Processed/            ← Reserved for future use
  04_Failed/               ← Failed files (if MOVE_FILES=true)
  05_Output/
    Dashboard_Data/        ← CSV outputs (Power BI / Excel ready)
    Summaries/             ← Markdown run summaries
    Reports/               ← Reserved
    Checks/                ← Reserved
  06_Logs/
    processing_log.csv     ← Run log
    Claude_API/            ← Audit logs for Claude API calls
  07_Config/               ← JSON configuration files
```

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

Copy `.env.example` to `.env`:

```bash
copy .env.example .env
```

Edit `.env`:

```env
ANTHROPIC_API_KEY=your_key_here
CLAUDE_MODEL=claude-sonnet-4-6
USE_CLAUDE_API=false
MOVE_FILES_AFTER_PROCESSING=false
```

## Running

```bash
python main.py
```

Optional flags:

```bash
python main.py --dry-run          # Scan and classify only, no output files
python main.py --use-claude       # Enable Claude API for this run
python main.py --move-files       # Copy files to 01_Raw after processing
python main.py --input-folder /path/to/other/inbox
```

## Expected Inputs

Place files in `00_Inbox/`:

| File type | Example | Detected as |
|-----------|---------|-------------|
| Bring PDF invoice | `Faktura 4040264491.pdf` | Bring Invoice |
| Bring Excel specification | `Specificeradfaktura_..._Fakturanummer_4040264491.xlsx` | Bring Specification |
| PostNord PDF | `postnord_faktura_*.pdf` | PostNord InvoiceAndSpecification |

## Expected Outputs

After a run, find outputs in `05_Output/Dashboard_Data/`:

| File | Contents |
|------|----------|
| `file_inventory.csv` | All scanned files and their classification |
| `invoice_header.csv` | Invoice-level header data (totals, dates, carrier) |
| `invoice_lines.csv` | All invoice lines (base freight + surcharges) |
| `surcharge_lines.csv` | Surcharge lines with category breakdown |
| `invoice_checks.csv` | Reconciliation and validation check results |

Summaries in `05_Output/Summaries/`:
- `summary_<run_id>_deterministic.md` — Always generated
- `summary_<run_id>_ai.md` — Generated if `USE_CLAUDE_API=true`

Logs in `06_Logs/`:
- `processing_log.csv` — All run events
- `Claude_API/claude_requests_<run_id>.jsonl` — Claude API audit log

## CSV Format

All CSV files use:
- **Delimiter:** semicolon (`;`) — compatible with Swedish Excel locale
- **Encoding:** UTF-8 with BOM (`utf-8-sig`) — opens correctly in Windows Excel
- **Append-friendly:** each run appends rows; headers written only once

## Claude API Usage and Safety

Claude API is **disabled by default** (`USE_CLAUDE_API=false`).

When enabled, Claude may only:
- Classify ambiguous service or surcharge names (when deterministic rules fail)
- Explain anomalies already detected by Python code
- Write a management summary from pre-calculated numbers

Claude may **never**:
- Calculate invoice totals, VAT, or reconciliation results
- Decide reconciliation status
- Change amounts or generate dashboard data
- Replace deterministic validation logic

All financial calculations, reconciliation checks, and validation statuses
are computed deterministically in Python. The script runs correctly with
Claude disabled.

## Why Claude Is Not Used for Financial Calculations

Freight invoices involve legally binding amounts. Using an AI model to calculate
totals or decide reconciliation status would:
1. Be non-auditable (non-deterministic output)
2. Risk financial errors from model hallucinations
3. Violate internal controls requirements for financial approval workflows

Claude is positioned as a soft-intelligence layer only — helping humans understand
the data, not determining what the data means financially.

## Current Limitations (MVP)

- **Bring** is fully implemented (PDF header + Excel specification lines)
- **PostNord** extracts invoice header only; detailed shipment-line parsing is a future phase
- CSV outputs are append-friendly but not deduplicated across runs (same invoice may appear twice)
- No Business Central integration
- No Power BI dashboard (use CSV outputs directly)
- No cloud deployment
- Runs locally on Windows; paths are relative to the project root

## Configuration Files (07_Config/)

| File | Purpose |
|------|---------|
| `carrier_rules.json` | Carrier detection keywords |
| `service_mapping.json` | Service category mappings |
| `surcharge_mapping.json` | Surcharge category mappings |
| `validation_rules.json` | Reconciliation tolerances |
| `anomaly_thresholds.json` | Cost and surcharge thresholds |
| `report_settings.json` | CSV delimiter, encoding, summary options |

## Next Development Steps

1. PostNord detailed shipment-line parsing
2. Deduplication of output rows across runs (run_id + invoice_number key)
3. Multi-invoice run handling (multiple Bring invoices in one inbox scan)
4. Power BI / Excel dashboard template connected to Dashboard_Data CSV files
5. Business Central API integration for approved invoice posting
6. SharePoint / OneDrive trigger automation via Power Automate
7. Email alerting on reconciliation errors
