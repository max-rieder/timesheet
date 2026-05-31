# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```bash
# Build shaded uber JAR (output: target/timesheet-1.0-SNAPSHOT.jar)
mvn package

# Clean rebuild
mvn clean package

# Stage 1 — parse the timesheet xlsx and emit work_log_summary.xlsx
java -cp target/timesheet-1.0-SNAPSHOT.jar ExcelParser <file_path> <hourly_rate>

# Stage 2 — generate the invoice PDF from work_log_summary.xlsx
java -cp target/timesheet-1.0-SNAPSHOT.jar InvoiceGenerator \
    work_log_summary.xlsx <invoice_number> "<invoice_date>" "<service_period>" <hourly_rate> [output_pdf]
# example:
# java -cp target/timesheet-1.0-SNAPSHOT.jar InvoiceGenerator \
#     work_log_summary.xlsx 2026-03 "31. März 2026" "März 2026" 96
```

The `maven-shade-plugin` produces a self-contained JAR — no need to add POI / iText to the classpath at runtime. Default main class set in the manifest is `ExcelParser`, so `java -jar` runs Stage 1; use `-cp` + class name to invoke `InvoiceGenerator`. No test framework is configured.

## Architecture

Two single-file Java CLI apps in `src/main/java/`, designed as a two-stage pipeline. There is no shared package — both classes sit in the default package.

```
input.xls  ──► ExcelParser  ──►  work_log_summary.xlsx  ──► InvoiceGenerator  ──►  invoice_<n>.pdf
```

### Stage 1 — `ExcelParser`

Parses a Jira-exported timesheet (`.xls` or `.xlsx`), aggregates by issue key, writes `work_log_summary.xlsx` to **CWD** (hard-coded path).

- Skips row 0; reads positional columns `0=issueKey, 1=summary, 2=hours, 3=date` (not header-driven).
- Inner classes: `WorkLogEntry` (immutable per-row holder), `WorkLogSummary` (per-issue aggregator that recalculates `netCost` on each `addHours()` call).
- Output layout: header row, one row per issue (`Pos | Issue | Total Hours | Net Cost`), `Total` row, `Total Sales Tax (20%)` row, `Gross Cost` row. German number format `#,##0.00`. All money is `BigDecimal` with `HALF_UP`.

### Stage 2 — `InvoiceGenerator`

Reads `work_log_summary.xlsx` and writes a German-language A4 invoice PDF (`Rechnung`) using iText 7.

- Splits rows back out by inspecting column 0: numeric → line item, `Total` / `Total Sales Tax…` / `Gross…` → summary numbers (the totals' values live in **column 2** in the summary, not column 3 — a quirk of the Stage 1 layout that Stage 2 has to mirror).
- Layout sections (each its own `add*` method): header (logo + sender address + banking on grey background), recipient block, invoice metadata, intro paragraph, services table, totals table, payment request.
- Logo (`AgileItSolutions.png`) is composited onto the grey background by `flattenOnBackground` because iText doesn't blend transparency cleanly into table cells. Read from CWD.
- Fonts are loaded by absolute path from `/mnt/c/Windows/Fonts/calibri.ttf` and `calibrib.ttf` — this is a **WSL-only assumption**. On native Linux/macOS these paths must be adjusted.
- Page geometry: A4, margins 2.5cm / 2.5cm / 2.0cm / 2.5cm, usable width 453pt — the column widths in the services and totals tables are tuned to that and shouldn't be changed without re-balancing.

## Notable Constraints

- `ExcelParser` always writes `work_log_summary.xlsx` to CWD, regardless of input file location.
- `ExcelParser` reads input columns positionally, not by header name.
- Formula cells in the input are not evaluated; `getCellValueAsString` returns the formula text.
- `InvoiceGenerator` parses the summary by matching column-0 strings (`Total`, `Total Sales Tax`, `Gross`) — changes to those labels in `ExcelParser` will break it silently.
- Calibri font paths are hard-coded for WSL.
- No streaming — entire workbook is loaded into memory (fine for invoices, would not be for arbitrary timesheets).
