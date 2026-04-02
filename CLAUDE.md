# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```bash
# Build JAR (output: target/timesheet-1.0-SNAPSHOT.jar)
mvn package

# Clean rebuild
mvn clean package

# Run (requires Apache POI on classpath — not bundled in JAR)
java -cp target/timesheet-1.0-SNAPSHOT.jar ExcelParser <file_path> <hourly_rate>
```

Note: The JAR contains only compiled classes, not POI dependencies. To run standalone, you'd need to add a Maven Shade or Assembly plugin to produce an uber JAR.

## Architecture

Single-file Java CLI app (`src/main/java/ExcelParser.java`). No test framework is configured.

**Flow:**
1. Takes two CLI args: path to an Excel file (.xls or .xlsx) and an hourly rate
2. Reads the first sheet, skipping row 0 (header), expecting columns: Issue Key, Issue Summary, Hours (numeric), Date (Excel date)
3. Aggregates total hours and cost per issue key
4. Writes a formatted summary to `work_log_summary.xlsx` in the **current working directory** (hard-coded path)

**Key classes (all in `ExcelParser.java`):**
- `ExcelParser` — main class with `main()`, parsing, and export logic
- `WorkLogEntry` — immutable data holder per row
- `WorkLogSummary` — mutable per-issue aggregator; recalculates `netCost` on each `addHours()` call

**Output format:** XLSX with Position, Issue, Total Hours, Net Cost rows, plus a 20% sales tax row and gross cost row. Numbers formatted with German locale (`#,##0.00`).

**Precision:** All monetary values use `BigDecimal` with `HALF_UP` rounding.

## Notable Constraints

- Output is always `work_log_summary.xlsx` relative to CWD, regardless of input file location
- Column order in the input is positional (0=key, 1=summary, 2=hours, 3=date) — not header-driven
- Formula cells are not evaluated; they return the formula string
- No streaming — entire workbook is loaded into memory
