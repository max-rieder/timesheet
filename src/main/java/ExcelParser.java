import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.*;

public class ExcelParser {

    public static void main(String[] args) {
        // Check if the file path and hourly rate are provided as arguments
        if (args.length < 2) {
            System.err.println("Please provide the Excel file path and hourly rate as program arguments.");
            System.exit(1); // Exit with an error code
        }

        String filePath = args[0]; // Read the file path from the program argument
        BigDecimal hourlyRate;

        try {
            hourlyRate = new BigDecimal(args[1]).setScale(2, RoundingMode.HALF_UP); // Parse the hourly rate
        } catch (NumberFormatException e) {
            System.err.println("Invalid hourly rate. Please provide a valid number.");
            System.exit(1); // Exit with an error code
            return; // To keep the code flow valid
        }

        List<WorkLogEntry> workLogEntries = new ArrayList<>();
        Map<String, WorkLogSummary> issueSummaryMap = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath)) {
            Workbook workbook;

            // Determine the file format (HSSF for .xls, XSSF for .xlsx)
            if (filePath.toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (filePath.toLowerCase().endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("The specified file is not an Excel file");
            }

            // Get the first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            // Skip the header row (start at row index 1)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue; // Skip empty rows

                // Parse each cell
                String issueKey = getCellValueAsString(row.getCell(0));
                String issueSummary = getCellValueAsString(row.getCell(1));
                BigDecimal hours = getCellValueAsBigDecimal(row.getCell(2));
                LocalDate workDate = getCellValueAsLocalDate(row.getCell(3));

                // Create a WorkLogEntry object and add it to the list
                WorkLogEntry entry = new WorkLogEntry(issueKey, issueSummary, hours, workDate);
                workLogEntries.add(entry);

                // Update the issue summary map with total work hours
                issueSummaryMap.computeIfAbsent(issueKey, key -> new WorkLogSummary(issueKey, issueSummary, hourlyRate))
                        .addHours(hours);
            }

            // Close the workbook
            workbook.close();

            // Export the results to a new Excel file
            exportResultsToExcel(issueSummaryMap, hourlyRate);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Helper method to get cell value as a String
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue()); // Handle numeric cells as strings
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }

    // Helper method to get cell value as BigDecimal
    private static BigDecimal getCellValueAsBigDecimal(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC) {
            return BigDecimal.ZERO; // Default value
        }
        return BigDecimal.valueOf(cell.getNumericCellValue());
    }

    // Helper method to get cell value as LocalDate
    private static LocalDate getCellValueAsLocalDate(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC) {
            return null; // Return null if not a valid date cell
        }

        Date date = cell.getDateCellValue();
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    // Export the results to a new Excel file
    private static void exportResultsToExcel(Map<String, WorkLogSummary> issueSummaryMap, BigDecimal hourlyRate) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Work Log Summary");

            // Create a cell style for bold text
            CellStyle boldStyle = workbook.createCellStyle();
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldStyle.setFont(boldFont);

            // Create a cell style for number formatting (German format)
            CellStyle numberStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            numberStyle.setDataFormat(dataFormat.getFormat("#,##0.00")); // German number format

            // Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Pos");
            headerRow.createCell(1).setCellValue("Issue");
            headerRow.createCell(2).setCellValue("Total Work Hours");
            headerRow.createCell(3).setCellValue("Net Cost");

            // Apply bold style to header
            for (Cell cell : headerRow) {
                cell.setCellStyle(boldStyle);
            }

            int rowNum = 1;
            BigDecimal totalNetCost = BigDecimal.ZERO;
            BigDecimal totalHoursWorked = BigDecimal.ZERO;

            // Write the summary data
            for (WorkLogSummary summary : issueSummaryMap.values()) {
                Row row = sheet.createRow(rowNum++);

                // Set the position column (Pos)
                row.createCell(0).setCellValue(rowNum - 1); // Position starts from 1

                // Combine issue key and issue summary into one "Issue" column
                String combinedIssue = summary.issueKey + " - " + summary.issueSummary;
                row.createCell(1).setCellValue(combinedIssue);
                row.createCell(2).setCellValue(summary.totalHours.doubleValue());
                row.createCell(3).setCellValue(summary.netCost.doubleValue());

                // Apply number formatting to hours and net cost
                row.getCell(2).setCellStyle(numberStyle);
                row.getCell(3).setCellStyle(numberStyle);

                // Accumulate total hours worked and net cost
                totalHoursWorked = totalHoursWorked.add(summary.totalHours);
                totalNetCost = totalNetCost.add(summary.netCost);
            }

            // Write the totals (make it bold) and shift the values one column to the right
            Row totalsRow = sheet.createRow(rowNum);
            totalsRow.createCell(0).setCellValue("Total");
            totalsRow.createCell(2).setCellValue(totalHoursWorked.doubleValue());
            totalsRow.createCell(3).setCellValue(totalNetCost.doubleValue());

            // Apply bold style to totals row
            for (Cell cell : totalsRow) {
                cell.setCellStyle(boldStyle);
            }

            // Apply number formatting to total hours and net cost
            totalsRow.getCell(2).setCellStyle(numberStyle);
            totalsRow.getCell(3).setCellStyle(numberStyle);

            // Calculate the sales tax and gross cost
            BigDecimal totalSalesTax = totalNetCost.multiply(new BigDecimal("0.2"));
            BigDecimal grossCost = totalNetCost.add(totalSalesTax);

            // Add a row for the financial summary (shifted one column to the right)
            rowNum++;
            Row financialRow = sheet.createRow(rowNum++);
            financialRow.createCell(0).setCellValue("Total Sales Tax (20%)");
            financialRow.createCell(2).setCellValue(totalSalesTax.doubleValue());

            // Apply bold style to financial rows
            for (Cell cell : financialRow) {
                cell.setCellStyle(boldStyle);
            }

            Row grossCostRow = sheet.createRow(rowNum);
            grossCostRow.createCell(0).setCellValue("Gross Cost (Net + Sales Tax)");
            grossCostRow.createCell(2).setCellValue(grossCost.doubleValue());

            // Apply bold style to the gross cost row
            for (Cell cell : grossCostRow) {
                cell.setCellStyle(boldStyle);
            }

            // Apply number formatting to financial rows
            financialRow.getCell(2).setCellStyle(numberStyle);
            grossCostRow.getCell(2).setCellStyle(numberStyle);

            // Adjust column widths to fit content
            for (int i = 0; i < 4; i++) {
                sheet.autoSizeColumn(i); // Resize columns to fit the content
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("work_log_summary.xlsx")) {
                workbook.write(fileOut);
            }

            System.out.println("Results exported to work_log_summary.xlsx");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }







    // WorkLogEntry class to represent a single work log entry
    static class WorkLogEntry {
        private final String issueKey;
        private final String issueSummary;
        private final BigDecimal hours;
        private final LocalDate workDate;

        public WorkLogEntry(String issueKey, String issueSummary, BigDecimal hours, LocalDate workDate) {
            this.issueKey = issueKey;
            this.issueSummary = issueSummary;
            this.hours = hours;
            this.workDate = workDate;
        }

        @Override
        public String toString() {
            return String.format("Issue Key: %s, Issue Summary: %s, Hours: %s, Work Date: %s",
                    issueKey, issueSummary, hours, workDate);
        }
    }

    // WorkLogSummary class to store the summary of total work hours per issue
    static class WorkLogSummary {
        private final String issueKey;
        private final String issueSummary;
        private final BigDecimal hourlyRate;
        private BigDecimal totalHours;
        private BigDecimal netCost;

        public WorkLogSummary(String issueKey, String issueSummary, BigDecimal hourlyRate) {
            this.issueKey = issueKey;
            this.issueSummary = issueSummary;
            this.hourlyRate = hourlyRate;
            this.totalHours = BigDecimal.ZERO;
            this.netCost = BigDecimal.ZERO;
        }

        // Add hours to the total work hours for this issue
        public void addHours(BigDecimal hours) {
            this.totalHours = this.totalHours.add(hours);
            // Recalculate the net cost for this issue
            this.netCost = this.totalHours.multiply(this.hourlyRate);
        }

        @Override
        public String toString() {
            return String.format("Issue Key: %s, Issue Summary: %s, Total Work Hours: %s, Net Cost: %s",
                    issueKey, issueSummary, totalHours, netCost);
        }
    }
}
