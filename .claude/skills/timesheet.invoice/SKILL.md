---
name: timesheet.invoice
description: "Generate a monthly invoice PDF from a Jira-exported timesheet. Runs the two-stage ExcelParser → InvoiceGenerator pipeline and writes Rechnung_Rieder_YYYY_MM.pdf. Usage: /timesheet.invoice [month] [rate] [invoice-date]"
argument-hint: "[month e.g. 'May 2026'] [hourly_rate, default 96] [invoice_date e.g. '01. Juni 2026']"
allowed-tools: Bash, Read, AskUserQuestion
---

Generate a monthly invoice PDF from `input.xls` in the project root, mirroring the layout of past invoices like `Rechnung_Rieder_2026_04.pdf`.

This skill must be invoked from `/home/luser/projects/temp/timesheet/` (the project root) — `ExcelParser` writes `work_log_summary.xlsx` to CWD and `InvoiceGenerator` reads it from CWD plus the logo (`AgileItSolutions.png`).

## Step 0 — Parse `$ARGUMENTS`

Three optional positional args, all space-separated. Quote multi-word values:

1. **Target month** — accepts `"May 2026"`, `"Mai 2026"`, `2026-05`, or `05/2026`. If omitted, ask the user via `AskUserQuestion` (header: "Month", options: current month, previous month, "Other").
2. **Hourly rate** — decimal, default `96`.
3. **Invoice date** — German-formatted day, e.g. `"31. Mai 2026"` or `"01. Juni 2026"`. Default: last day of the target month, German-formatted.

Derive these variables for use below. Always render German month names — map `01..12` → `Januar Februar März April Mai Juni Juli August September Oktober November Dezember`:

- `$YEAR` = 4-digit year of target month
- `$MONTH_NUM` = zero-padded 2-digit month (`05`)
- `$MONTH_DE` = German month name (`Mai`)
- `$INVOICE_NUMBER` = `$YEAR-$MONTH_NUM` (e.g. `2026-05`)
- `$SERVICE_PERIOD` = `"$MONTH_DE $YEAR"` (e.g. `"Mai 2026"`)
- `$INVOICE_DATE` = arg 3, or last day of target month formatted `"DD. $MONTH_DE $YEAR"` (e.g. `"31. Mai 2026"`)
- `$RATE` = arg 2 or `96`
- `$OUTPUT_PDF` = `Rechnung_Rieder_${YEAR}_${MONTH_NUM}.pdf`

Print the resolved values back to the user as a one-line sanity check before running the pipeline.

## Step 1 — Ensure JAR is built

```bash
test -f target/timesheet-1.0-SNAPSHOT.jar || mvn -q package
```

If the JAR is missing, build it. If `mvn` fails, stop and surface the error.

## Step 2 — Resolve the input timesheet file

Default `$INPUT_FILE=input.xls`. Check it exists:

```bash
test -f "$INPUT_FILE"
```

If it does **not** exist, list candidate `.xls`/`.xlsx` files in CWD to give the user options:

```bash
ls -1 *.xls *.xlsx 2>/dev/null
```

Then ask via `AskUserQuestion` (header: "Input file"):

- If candidates were found, offer them as the first options plus an "Other" path.
- If none, just ask the user to provide a path (relative to project root or absolute).

Validate the chosen path with `test -f "$INPUT_FILE"`. If still missing, ask again or stop.

## Step 3 — Stage 1: parse timesheet

```bash
java -cp target/timesheet-1.0-SNAPSHOT.jar ExcelParser "$INPUT_FILE" "$RATE"
```

Expect `work_log_summary.xlsx` to appear in CWD. If it doesn't, stop.

## Step 4 — Stage 2: generate the PDF

```bash
java -cp target/timesheet-1.0-SNAPSHOT.jar InvoiceGenerator \
    work_log_summary.xlsx \
    "$INVOICE_NUMBER" \
    "$INVOICE_DATE" \
    "$SERVICE_PERIOD" \
    "$RATE" \
    "$OUTPUT_PDF"
```

## Step 5 — Verify

```bash
ls -la "$OUTPUT_PDF"
pdftotext "$OUTPUT_PDF" - 2>/dev/null | grep -E "Rechnungsnummer|Rechnungsdatum|Leistungszeitraum" || true
```

Confirm the three header fields match `$INVOICE_NUMBER`, `$INVOICE_DATE`, `$SERVICE_PERIOD`. If any mismatches, regenerate Stage 2 with corrected args.

## Final report

Print one line: `Wrote $OUTPUT_PDF — $INVOICE_NUMBER, $INVOICE_DATE, $SERVICE_PERIOD, €$RATE/h`.

## Notes & gotchas

- `ExcelParser` always writes `work_log_summary.xlsx` to **CWD** (hard-coded). Run from the project root.
- `InvoiceGenerator` also reads from CWD and needs `AgileItSolutions.png` there.
- Fonts are hard-coded to `/mnt/c/Windows/Fonts/calibri{,b}.ttf` — WSL only.
- The Stage-1 output overwrites any existing `work_log_summary.xlsx`.
- If the user wants to **only change the invoice date** of an already-generated PDF, skip Stage 1 and re-run Stage 2 with the new date — `work_log_summary.xlsx` from the previous run is still on disk.
- To support a non-WSL host, the font paths in `InvoiceGenerator.java` must be edited; this skill does not patch them.
