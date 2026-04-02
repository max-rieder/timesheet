import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Generates an invoice PDF from a work_log_summary.xlsx file.
 *
 * Usage: InvoiceGenerator <xlsx_path> <invoice_number> <invoice_date> <service_period> <hourly_rate> [output_pdf]
 * Example: InvoiceGenerator work_log_summary.xlsx 2026-03 "31. März 2026" "März 2026" 96
 *
 * Expects AgileItSolutions.png in the current working directory.
 */
public class InvoiceGenerator {

    private static final DeviceRgb GRAY_BG   = new DeviceRgb(0xD9, 0xD9, 0xD9);
    private static final DeviceRgb DARK_TEXT = new DeviceRgb(0x26, 0x26, 0x26);

    private static final String FONT_REGULAR = "/mnt/c/Windows/Fonts/calibri.ttf";
    private static final String FONT_BOLD    = "/mnt/c/Windows/Fonts/calibrib.ttf";

    private static final float FONT_SIZE = 10f;
    private static final Locale DE = Locale.GERMANY;

    // -------------------------------------------------------------------------
    // Entry point
    // -------------------------------------------------------------------------

    public static void main(String[] args) throws Exception {
        if (args.length < 5) {
            System.err.println("Usage: InvoiceGenerator <xlsx_path> <invoice_number> <invoice_date> <service_period> <hourly_rate> [output_pdf]");
            System.exit(1);
        }
        String xlsxPath      = args[0];
        String invoiceNumber = args[1];
        String invoiceDate   = args[2];
        String servicePeriod = args[3];
        String hourlyRate    = args[4];
        String outputPdf     = args.length > 5 ? args[5] : "invoice_" + invoiceNumber + ".pdf";

        List<InvoiceItem> rows = readXlsx(xlsxPath);
        generatePdf(rows, invoiceNumber, invoiceDate, servicePeriod, hourlyRate, outputPdf);
        System.out.println("Invoice written to: " + outputPdf);
    }

    // -------------------------------------------------------------------------
    // Read xlsx
    // -------------------------------------------------------------------------

    static List<InvoiceItem> readXlsx(String path) throws Exception {
        List<InvoiceItem> items = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(path);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet = wb.getSheetAt(0);
            boolean header = true;
            for (Row row : sheet) {
                if (header) { header = false; continue; }
                String pos   = cellStr(row, 0);
                String issue = cellStr(row, 1);
                BigDecimal hours = cellDecimal(row, 2);
                BigDecimal cost  = cellDecimal(row, 3);
                items.add(new InvoiceItem(pos, issue, hours, cost));
            }
        }
        return items;
    }

    private static String cellStr(Row row, int col) {
        var cell = row.getCell(col);
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                yield v == Math.floor(v) ? String.valueOf((long) v) : String.valueOf(v);
            }
            case STRING  -> cell.getStringCellValue();
            default      -> "";
        };
    }

    private static BigDecimal cellDecimal(Row row, int col) {
        var cell = row.getCell(col);
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue());
            case STRING  -> {
                try { yield new BigDecimal(cell.getStringCellValue().trim()); }
                catch (Exception e) { yield null; }
            }
            default -> null;
        };
    }

    // -------------------------------------------------------------------------
    // Generate PDF
    // -------------------------------------------------------------------------

    static void generatePdf(List<InvoiceItem> allRows,
                            String invoiceNumber, String invoiceDate,
                            String servicePeriod, String hourlyRate,
                            String outputPath) throws Exception {

        // Split line items from summary rows
        List<InvoiceItem> lineItems = new ArrayList<>();
        BigDecimal totalHours = BigDecimal.ZERO;
        BigDecimal netTotal   = BigDecimal.ZERO;
        BigDecimal salesTax   = BigDecimal.ZERO;
        BigDecimal grossTotal = BigDecimal.ZERO;

        for (InvoiceItem item : allRows) {
            if (item.pos.matches("\\d+")) {
                lineItems.add(item);
            } else if (item.pos.equals("Total")) {
                if (item.hours != null) totalHours = item.hours;
                if (item.cost  != null) netTotal   = item.cost;
            } else if (item.pos.startsWith("Total Sales Tax")) {
                if (item.hours != null) salesTax = item.hours;
            } else if (item.pos.startsWith("Gross")) {
                if (item.hours != null) grossTotal = item.hours;
            }
        }

        // Fonts
        PdfFont regular = PdfFontFactory.createFont(FONT_REGULAR, PdfEncodings.IDENTITY_H);
        PdfFont bold    = PdfFontFactory.createFont(FONT_BOLD,    PdfEncodings.IDENTITY_H);

        try (PdfDocument pdf = new PdfDocument(new PdfWriter(outputPath));
             Document doc = new Document(pdf, PageSize.A4)) {

            // A4: 595pt wide, 841pt tall
            // Margins: top/left/right 2.5cm = 70.9pt, bottom 2cm = 56.7pt
            doc.setMargins(70.9f, 70.9f, 56.7f, 70.9f);

            addHeader(doc, regular, bold);
            addRecipient(doc, regular, bold);
            addInvoiceMeta(doc, regular, invoiceNumber, invoiceDate, servicePeriod);
            addDescriptionParagraph(doc, regular, bold, hourlyRate);
            addServicesTable(doc, regular, bold, lineItems, totalHours, netTotal, salesTax, grossTotal);
            addPaymentRequest(doc, bold);
        }
    }

    // -------------------------------------------------------------------------
    // Header: two-column table with logo+address | contact+banking
    // -------------------------------------------------------------------------

    private static void addHeader(Document doc, PdfFont regular, PdfFont bold) throws Exception {
        Table table = new Table(new float[]{1f, 1f})
                .setWidth(UnitValue.createPercentValue(100));

        // Left cell: gray background, logo, address in green
        Cell left = new Cell()
                .setBackgroundColor(GRAY_BG)
                .setBorder(Border.NO_BORDER)
                .setPadding(8f);

        Image logo = new Image(ImageDataFactory.create(flattenOnBackground("AgileItSolutions.png", 0xD9, 0xD9, 0xD9)));
        logo.setWidth(UnitValue.createPointValue(207f));   // ~7.3cm
        logo.setHeight(UnitValue.createPointValue(38f));   // ~1.35cm
        logo.setMarginBottom(16f);
        left.add(logo);

        left.add(new Paragraph()
                .add(new Text("Dipl.-Ing. Max Rieder\n").setFont(bold))
                .add(new Text("IT-Dienstleistung\n"))
                .add(new Text("Utendorfgasse 27/2/1\n"))
                .add(new Text("1140 Wien\n"))
                .add(new Text("Österreich"))
                .setFont(regular).setFontSize(FONT_SIZE).setFontColor(DARK_TEXT)
                .setMultipliedLeading(1.0f)
                .setMarginBottom(0));
        table.addCell(left);

        // Right cell: contact info and banking details
        Cell right = new Cell()
                .setBackgroundColor(GRAY_BG)
                .setBorder(Border.NO_BORDER)
                .setPaddingTop(8f)
                .setPaddingLeft(12f);

        for (String line : new String[]{
                "Telefon: +43-650-8648193",
                "Email: max.rieder@zoho.eu",
                " ",
                "Bankinstitut: easybank AG, BLZ: 14200",
                "Kontonummer: 20011726217",
                "IBAN: AT461420020011726217, BIC: EASYATW1",
                "UID: ATU68110055, Steuernummer: 361/3514"
        }) {
            right.add(new Paragraph(line).setFont(regular).setFontSize(FONT_SIZE).setMarginBottom(0));
        }
        table.addCell(right);

        doc.add(table);
    }

    // -------------------------------------------------------------------------
    // Recipient block
    // -------------------------------------------------------------------------

    private static void addRecipient(Document doc, PdfFont regular, PdfFont bold) {
        doc.add(new Paragraph()
                .add(new Text("Expleo Group Austria GmbH\n"))
                .add(new Text("Theresianumgasse 11/1\n"))
                .add(new Text("1040 Wien"))
                .setFont(bold).setFontSize(FONT_SIZE).setMultipliedLeading(1.0f).setMarginTop(20f).setMarginBottom(2f));

        doc.add(spacer(4f));

        doc.add(new Paragraph("Umsatzsteuer-Identifikationsnummer: ATU51868206")
                .setFont(bold).setFontSize(FONT_SIZE).setMarginBottom(0));

        doc.add(spacer(8f));
    }

    // -------------------------------------------------------------------------
    // Invoice metadata
    // -------------------------------------------------------------------------

    private static void addInvoiceMeta(Document doc, PdfFont regular,
                                       String invoiceNumber, String invoiceDate, String servicePeriod) {
        doc.add(new Paragraph("Rechnungsnummer: " + invoiceNumber)
                .setFont(regular).setFontSize(FONT_SIZE).setMarginBottom(2f));
        doc.add(new Paragraph("Rechnungsdatum: " + invoiceDate)
                .setFont(regular).setFontSize(FONT_SIZE).setMarginBottom(2f));
        doc.add(new Paragraph("Leistungszeitraum: " + servicePeriod)
                .setFont(regular).setFontSize(FONT_SIZE).setMarginBottom(2f));
        doc.add(spacer(8f));
    }

    // -------------------------------------------------------------------------
    // Description paragraph
    // -------------------------------------------------------------------------

    private static void addDescriptionParagraph(Document doc, PdfFont regular, PdfFont bold, String hourlyRate) {
        doc.add(new Paragraph()
                .add(new Text("Für den von Ihnen erteilten Auftrag gemäß Einzelvertrag vom 21. Juli 2024 " +
                              "verrechne ich vereinbarungsgemäß zu einem ").setFont(regular))
                .add(new Text("Stundenhonorar von \u20ac " + hourlyRate + ".- exkl. USt.").setFont(bold))
                .add(new Text(" wie folgt:").setFont(regular))
                .setFontSize(FONT_SIZE).setMultipliedLeading(1.0f)
                .setMarginBottom(8f));
    }

    // -------------------------------------------------------------------------
    // Services table
    // -------------------------------------------------------------------------

    private static void addServicesTable(Document doc, PdfFont regular, PdfFont bold,
                                         List<InvoiceItem> lineItems,
                                         BigDecimal totalHours, BigDecimal netTotal,
                                         BigDecimal salesTax, BigDecimal grossTotal) {

        // Column widths in points (page 595pt - margins 2x70.9 = 453pt usable)
        Table table = new Table(UnitValue.createPointArray(new float[]{25f, 210f, 90f, 128f}))
                .setWidth(UnitValue.createPointValue(453f))
                .setFontSize(FONT_SIZE);

        // Header row
        for (String label : new String[]{"Pos", "Leistung", "Leistungsstunden", "Nettopreis in\u00a0\u20ac"}) {
            table.addHeaderCell(new Cell()
                    .add(new Paragraph(label).setFont(bold).setMultipliedLeading(1.0f).setTextAlignment(TextAlignment.CENTER))
                    .setBackgroundColor(GRAY_BG)
                    .setBorder(new SolidBorder(0.5f))
                    .setPaddingTop(4f).setPaddingBottom(4f).setPaddingLeft(2f).setPaddingRight(2f));
        }

        // Data rows
        for (InvoiceItem item : lineItems) {
            table.addCell(dataCell(item.pos, regular, TextAlignment.LEFT));
            table.addCell(dataCell(item.issue, regular, TextAlignment.LEFT));
            table.addCell(dataCell(item.hours != null ? fmtHours(item.hours) : "", regular, TextAlignment.RIGHT));
            table.addCell(dataCell(item.cost  != null ? fmtCurrency(item.cost) : "", regular, TextAlignment.RIGHT));
        }

        // Summary: Zwischensumme + MwSt (spans 3 cols for label, 1 col for amounts)
        table.addCell(new Cell(1, 3)
                .add(new Paragraph()
                        .add(new Text("Zwischensumme, exkl. MwSt.\n").setFont(bold))
                        .add(new Text("MwSt. 20%").setFont(bold)))
                .setBorder(new SolidBorder(0.5f))
                .setPadding(4f));
        table.addCell(new Cell(1, 1)
                .add(new Paragraph()
                        .add(new Text(fmtCurrency(netTotal) + "\n").setFont(bold))
                        .add(new Text(fmtCurrency(salesTax)).setFont(bold)))
                .setBorder(new SolidBorder(0.5f))
                .setTextAlignment(TextAlignment.RIGHT)
                .setPadding(4f));

        doc.add(table);

        // Totals below table, label right-aligned to the Leistungsstunden/Nettopreis divider (325pt)
        // Value right-aligned within the 128pt Nettopreis column
        Table totalsTable = new Table(UnitValue.createPointArray(new float[]{325f, 128f}))
                .setWidth(UnitValue.createPointValue(453f));

        totalsTable.addCell(new Cell()
                .add(new Paragraph("Leistungsstunden gesamt").setFont(regular).setMultipliedLeading(1.0f).setTextAlignment(TextAlignment.RIGHT))
                .setBorder(Border.NO_BORDER).setPaddingTop(14f).setPaddingBottom(2f).setPaddingRight(4f));
        totalsTable.addCell(new Cell()
                .add(new Paragraph(fmtHours(totalHours)).setFont(regular).setMultipliedLeading(1.0f).setTextAlignment(TextAlignment.RIGHT))
                .setBorder(Border.NO_BORDER).setPaddingTop(14f).setPaddingBottom(2f).setPaddingRight(4f));

        totalsTable.addCell(new Cell()
                .add(new Paragraph("Rechnungsbetrag, inkl. MwSt.").setFont(bold).setMultipliedLeading(1.0f).setTextAlignment(TextAlignment.RIGHT))
                .setBorder(Border.NO_BORDER).setPaddingTop(10f).setPaddingBottom(4f).setPaddingRight(4f));
        totalsTable.addCell(new Cell()
                .add(new Paragraph("\u20ac " + fmtCurrency(grossTotal)).setFont(bold).setMultipliedLeading(1.0f).setTextAlignment(TextAlignment.RIGHT))
                .setBorder(Border.NO_BORDER).setPaddingTop(10f).setPaddingBottom(4f).setPaddingRight(4f));

        doc.add(totalsTable);
    }

    // -------------------------------------------------------------------------
    // Payment request
    // -------------------------------------------------------------------------

    private static void addPaymentRequest(Document doc, PdfFont bold) {
        doc.add(new Paragraph(
                "Ich bitte Sie, den Betrag innerhalb von 30 Tagen auf oben genanntes Bankkonto, " +
                "lautend auf meinen Namen, zu überweisen.")
                .setFont(bold).setFontSize(FONT_SIZE).setMultipliedLeading(1.0f).setMarginTop(20f));
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static Cell dataCell(String text, PdfFont font, TextAlignment align) {
        return new Cell()
                .add(new Paragraph(text).setFont(font).setMultipliedLeading(1.0f))
                .setBorder(new SolidBorder(0.5f))
                .setTextAlignment(align)
                .setPaddingTop(2f).setPaddingBottom(2f).setPaddingLeft(4f).setPaddingRight(4f);
    }

    private static Paragraph spacer(float fontSize) {
        return new Paragraph(" ").setFontSize(fontSize).setMarginBottom(0).setMarginTop(0);
    }

    /** Composites a PNG with transparency onto a solid background colour, returning PNG bytes. */
    private static byte[] flattenOnBackground(String path, int r, int g, int b) throws Exception {
        BufferedImage src = ImageIO.read(new File(path));
        BufferedImage dst = new BufferedImage(src.getWidth(), src.getHeight(), BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = dst.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        g2.setColor(new Color(r, g, b));
        g2.fillRect(0, 0, dst.getWidth(), dst.getHeight());
        g2.drawImage(src, 0, 0, null);
        g2.dispose();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(dst, "JPEG", baos);
        return baos.toByteArray();
    }

    private static String fmtHours(BigDecimal v) {
        return String.format(DE, "%,.2f", v);
    }

    private static String fmtCurrency(BigDecimal v) {
        return String.format(DE, "%,.2f", v);
    }

    // -------------------------------------------------------------------------
    // Data holder
    // -------------------------------------------------------------------------

    record InvoiceItem(String pos, String issue, BigDecimal hours, BigDecimal cost) {}
}
