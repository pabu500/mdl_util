package org.pabuff.utils;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.AreaBreakType;
import com.itextpdf.layout.properties.Property;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class PdfUtil {
//    public static void convertExcelToPdf(File excelFile, File outputDir) {
//        try {
//            Files.createDirectories(outputDir.toPath());
//
//            ProcessBuilder processBuilder = new ProcessBuilder(
//                    "libreoffice",
//                    "--headless",
//                    "--convert-to",
//                    "pdf",
//                    "--outdir",
//                    outputDir.getAbsolutePath(),
//                    excelFile.getAbsolutePath()
//            );
//            processBuilder.redirectErrorStream(true);
//            Process process = processBuilder.start();
//            boolean finished = process.waitFor(60, TimeUnit.SECONDS);
//            if (!finished) {
//                process.destroyForcibly();
//                throw new RuntimeException("LibreOffice PDF conversion timeout");
//            }
//
//            int exitCode = process.exitValue();
//            if (exitCode != 0) {
//                throw new RuntimeException("LibreOffice PDF conversion failed. Exit code: " + exitCode);
//            }
//
//        } catch (IOException e) {
//            throw new RuntimeException("Failed to run LibreOffice PDF conversion", e);
//        } catch (InterruptedException e) {
//            Thread.currentThread().interrupt();
//            throw new RuntimeException("PDF conversion interrupted", e);
//        }
//    }
//
//    public static File getExpectedPdfFile(File excelFile, File outputDir) {
//        String excelName = excelFile.getName();
//        String pdfName = excelName.replaceFirst("(?i)\\.xlsx$", ".pdf")
//                .replaceFirst("(?i)\\.xls$", ".pdf");
//
//        File pdfFile = new File(outputDir, pdfName);
//        if (!pdfFile.exists()) {
//            throw new RuntimeException("PDF file was not created: " + pdfFile.getAbsolutePath());
//        }
//        return pdfFile;
//    }
//
//    public static File getOutputDir(File excelFile, Map<String, Object> request) {
//        String outputDir = (String)  request.get("output_dir");
//        if (outputDir != null && !outputDir.isBlank()) {
//            return new File(outputDir);
//        }
//        return excelFile.getParentFile();
//    }

    public static Map<String, Object> convertExcelToPdf(
            File excelFile,
            File outputDir,
            Map<String, Object> request) {

        if (excelFile == null || !excelFile.exists() || !excelFile.isFile()) {
            throw new RuntimeException("Excel file not found: " + excelFile);
        }

        if (outputDir == null) {
            outputDir = excelFile.getParentFile();
        }

        if (!outputDir.exists() && !outputDir.mkdirs()) {
            throw new RuntimeException("Failed to create output directory: " + outputDir.getAbsolutePath());
        }

        if (request == null) {
            request = new HashMap<>();
        }

        String pdfFileName = excelFile.getName()
                .replaceFirst("(?i)\\.xlsx$", ".pdf")
                .replaceFirst("(?i)\\.xls$", ".pdf");
        File pdfFile = new File(outputDir, pdfFileName);

        String orientation = (String) request.getOrDefault("orientation", "portrait");
        boolean landscape = "landscape".equalsIgnoreCase(orientation);
//        boolean showSheetName = getBoolean(request, "show_sheet_name", true);
        boolean includeHiddenSheets = (Boolean) request.getOrDefault("include_hidden_sheets", false);

        PageSize pageSize = landscape ? PageSize.A4.rotate() : PageSize.A4;

        try (
                FileInputStream fis = new FileInputStream(excelFile);
                Workbook workbook = WorkbookFactory.create(fis);
                PdfWriter writer = new PdfWriter(pdfFile);
                PdfDocument pdfDocument = new PdfDocument(writer);
                Document document = new Document(pdfDocument, pageSize)
        ) {
            double leftMargin = (Double) request.getOrDefault("left_margin", 20.0);
            double rightMargin = (Double) request.getOrDefault("right_margin", 20.0);
            double topMargin = (Double) request.getOrDefault("top_margin", 20.0);
            double bottomMargin = (Double) request.getOrDefault("bottom_margin", 20.0);

            document.setMargins(
                    (float) topMargin,
                    (float) rightMargin,
                    (float) bottomMargin,
                    (float) leftMargin
            );

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter formatter = new DataFormatter();

            boolean firstSheetAdded = false;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (!includeHiddenSheets && (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i))) {
                    continue;
                }

                Sheet sheet = workbook.getSheetAt(i);
                if (getMaxColumnCount(sheet) <= 0) {
                    continue;
                }

                int maxColumns = getMaxColumnCount(sheet);
                if (maxColumns <= 0) {
                    continue;
                }

                double tableScaleRatio = getTableScaleRatio(
                        sheet, maxColumns, pageSize,
                        leftMargin, rightMargin
                );

                if (firstSheetAdded) {
                    document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }

                addSheetToPdf(document, workbook, sheet, formatter, evaluator, tableScaleRatio);
                firstSheetAdded = true;
            }

            Map<String, Object> result = new HashMap<>();
            result.put("file", pdfFile);
            result.put("pdf_files", java.util.List.of(pdfFile));
            result.put("excel_file", excelFile);
            result.put("extension", ".pdf");
            result.put("full_report_name", pdfFileName.replaceFirst("(?i)\\.pdf$", ""));
            return result;

        } catch (Exception e) {
            throw new RuntimeException("Failed to convert Excel to PDF using POI + iText", e);
        }
    }

    private static void addSheetToPdf(
            Document document,
            Workbook workbook,
            Sheet sheet,
            DataFormatter formatter,
            FormulaEvaluator evaluator,
            double tableScaleRatio) {

        int maxColumns = getMaxColumnCount(sheet);
        if (maxColumns <= 0) {
            return;
        }

        float[] columnWidths = getColumnWidths(sheet, maxColumns);

        com.itextpdf.layout.element.Table table = new Table(UnitValue.createPercentArray(columnWidths));
        table.setWidth(UnitValue.createPercentValue(100));
        table.setFixedLayout();

        Set<String> skippedCells = new HashSet<>();
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {

            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getZeroHeight()) {
                continue;
            }

            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                if (sheet.isColumnHidden(colIndex)) {
                    continue;
                }

                String cellKey = rowIndex + ":" + colIndex;
                if (skippedCells.contains(cellKey)) {
                    continue;
                }

                CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, colIndex);
                if (mergedRegion != null) {
                    boolean isTopLeft =
                            mergedRegion.getFirstRow() == rowIndex &&
                                    mergedRegion.getFirstColumn() == colIndex;
                    if (!isTopLeft) {
                        continue;
                    }
                    markMergedCellsAsSkipped(skippedCells, mergedRegion, rowIndex, colIndex);
                }

                org.apache.poi.ss.usermodel.Cell poiCell = row == null
                        ? null
                        : row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                int rowSpan = 1;
                int colSpan = 1;

                if (mergedRegion != null) {
                    rowSpan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                    colSpan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                }

                com.itextpdf.layout.element.Cell pdfCell =
                        new com.itextpdf.layout.element.Cell(rowSpan, colSpan);

                String cellText = getCellText(poiCell, formatter, evaluator);
                Paragraph paragraph = new Paragraph(cellText).setMargin(0);

                if (poiCell != null) {
                    CellStyle style = poiCell.getCellStyle();
                    if (style != null && !style.getWrapText()) {
                        paragraph.setProperty(Property.NO_SOFT_WRAP_INLINE, true);
                    }
                    applyCellStyle(pdfCell, workbook, poiCell, tableScaleRatio);
                } else {
//                    pdfCell.setBorder(new SolidBorder(new DeviceRgb(230, 230, 230), 0.3f));
                    pdfCell.setBorder(Border.NO_BORDER);
                }

                pdfCell.add(paragraph);
                table.addCell(pdfCell);
            }
        }
        document.add(table);
    }

    private static String getCellText(
            org.apache.poi.ss.usermodel.Cell cell,
            DataFormatter formatter,
            FormulaEvaluator evaluator) {

        if (cell == null) {
            return "";
        }

        try {
            if (cell.getCellType() == CellType.FORMULA) {
                return formatter.formatCellValue(cell, evaluator);
            }

            return formatter.formatCellValue(cell);
        } catch (Exception e) {
            return cell.toString();
        }
    }

    private static void applyCellStyle(
            com.itextpdf.layout.element.Cell pdfCell,
            Workbook workbook,
            org.apache.poi.ss.usermodel.Cell poiCell,
            double tableScaleRatio) {

        CellStyle style = poiCell.getCellStyle();

        if (style == null) {
            return;
        }

        applyFont(pdfCell, workbook, style, tableScaleRatio);
        applyFillColor(pdfCell, style);
        applyAlignment(pdfCell, style);
        applyBorders(pdfCell, style);
//        pdfCell.setProperty(Property.NO_SOFT_WRAP_INLINE, style.getWrapText());

        float padding = (float) Math.max(1.0, 3.0 * tableScaleRatio);
        pdfCell.setPadding(padding);
    }

    private static void applyFont(
            com.itextpdf.layout.element.Cell pdfCell,
            Workbook workbook,
            CellStyle style,
            double tableScaleRatio) {

        Font font = workbook.getFontAt(style.getFontIndex());
        if (font == null) {
            return;
        }

        // 1. Handle Bold and Italic using standard font families
        try {
//            String fontName = font.getFontName();
            boolean isBold = font.getBold();
            boolean isItalic = font.getItalic();

            PdfFont pdfFont;
            if (isBold && isItalic) {
                pdfFont = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLDOBLIQUE);
            } else if (isBold) {
                pdfFont = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
            } else if (isItalic) {
                pdfFont = PdfFontFactory.createFont(StandardFonts.HELVETICA_OBLIQUE);
            } else {
                pdfFont = PdfFontFactory.createFont(StandardFonts.HELVETICA);
            }

            pdfCell.setFont(pdfFont);
        } catch (IOException e) {
            // Handle exception for font creation
        }

        // 2. Set Font Size
        if (font.getFontHeightInPoints() > 0) {
            float scaledFontSize = (float) (font.getFontHeightInPoints() * tableScaleRatio);
            // Optional minimum font size, to avoid unreadable text
            scaledFontSize = Math.max(5.5f, scaledFontSize);
            pdfCell.setFontSize(scaledFontSize);
        }

        // 3. Set Font Color
        Color fontColor = getPoiFontColorAsITextColor(font);
        if (fontColor != null) {
            pdfCell.setFontColor((com.itextpdf.kernel.colors.Color) fontColor);
        }
    }

    private static void applyFillColor(
            com.itextpdf.layout.element.Cell pdfCell,
            CellStyle style) {

        if (style.getFillPattern() == FillPatternType.NO_FILL) {
            return;
        }

        com.itextpdf.kernel.colors.Color color = getPoiFillColorAsITextColor(style);

        if (color != null) {
            pdfCell.setBackgroundColor(color);
        }
    }

    private static void applyAlignment(
            com.itextpdf.layout.element.Cell pdfCell,
            CellStyle style) {

        HorizontalAlignment horizontalAlignment = style.getAlignment();

        if (horizontalAlignment == HorizontalAlignment.CENTER) {
            pdfCell.setTextAlignment(TextAlignment.CENTER);
        } else if (horizontalAlignment == HorizontalAlignment.RIGHT) {
            pdfCell.setTextAlignment(TextAlignment.RIGHT);
        } else {
            pdfCell.setTextAlignment(TextAlignment.LEFT);
        }

        VerticalAlignment verticalAlignment = style.getVerticalAlignment();

        if (verticalAlignment == VerticalAlignment.CENTER) {
            pdfCell.setVerticalAlignment(com.itextpdf.layout.properties.VerticalAlignment.MIDDLE);
        } else if (verticalAlignment == VerticalAlignment.BOTTOM) {
            pdfCell.setVerticalAlignment(com.itextpdf.layout.properties.VerticalAlignment.BOTTOM);
        } else {
            pdfCell.setVerticalAlignment(com.itextpdf.layout.properties.VerticalAlignment.TOP);
        }
    }

    private static void applyBorders(
            com.itextpdf.layout.element.Cell pdfCell,
            CellStyle style) {

        boolean hasBorder =
                style.getBorderTop() != BorderStyle.NONE ||
                        style.getBorderBottom() != BorderStyle.NONE ||
                        style.getBorderLeft() != BorderStyle.NONE ||
                        style.getBorderRight() != BorderStyle.NONE;

        if (!hasBorder) {
//            pdfCell.setBorder(new SolidBorder(new DeviceRgb(230, 230, 230), 0.3f));
            pdfCell.setBorder(Border.NO_BORDER);
            return;
        }

        if (style.getBorderTop() != BorderStyle.NONE) {
            pdfCell.setBorderTop(new SolidBorder(0.5f));
        }

        if (style.getBorderBottom() != BorderStyle.NONE) {
            pdfCell.setBorderBottom(new SolidBorder(0.5f));
        }

        if (style.getBorderLeft() != BorderStyle.NONE) {
            pdfCell.setBorderLeft(new SolidBorder(0.5f));
        }

        if (style.getBorderRight() != BorderStyle.NONE) {
            pdfCell.setBorderRight(new SolidBorder(0.5f));
        }
    }

    private static com.itextpdf.kernel.colors.Color getPoiFillColorAsITextColor(CellStyle style) {
        org.apache.poi.ss.usermodel.Color poiColor = style.getFillForegroundColorColor();

        if (poiColor instanceof XSSFColor xssfColor) {
            byte[] rgb = xssfColor.getRGB();

            if (rgb != null && rgb.length >= 3) {
                return new DeviceRgb(
                        Byte.toUnsignedInt(rgb[0]),
                        Byte.toUnsignedInt(rgb[1]),
                        Byte.toUnsignedInt(rgb[2])
                );
            }
        }

        if (poiColor instanceof HSSFColor hssfColor) {
            short[] triplet = hssfColor.getTriplet();

            if (triplet != null && triplet.length >= 3) {
                return new DeviceRgb(triplet[0], triplet[1], triplet[2]);
            }
        }

        short colorIndex = style.getFillForegroundColor();

        if (colorIndex == IndexedColors.YELLOW.getIndex()) {
            return new DeviceRgb(255, 255, 0);
        }

        if (colorIndex == IndexedColors.GREY_25_PERCENT.getIndex()) {
            return new DeviceRgb(217, 217, 217);
        }

        if (colorIndex == IndexedColors.LIGHT_BLUE.getIndex()) {
            return new DeviceRgb(173, 216, 230);
        }

        return null;
    }

    private static com.itextpdf.kernel.colors.Color getPoiFontColorAsITextColor(Font font) {
        if (font == null) {
            return null;
        }

        short colorIndex = font.getColor();

        if (colorIndex == IndexedColors.RED.getIndex()) {
            return new DeviceRgb(255, 0, 0);
        }

        if (colorIndex == IndexedColors.BLUE.getIndex()) {
            return new DeviceRgb(0, 0, 255);
        }

        if (colorIndex == IndexedColors.GREEN.getIndex()) {
            return new DeviceRgb(0, 128, 0);
        }

        if (colorIndex == IndexedColors.WHITE.getIndex()) {
            return new DeviceRgb(255, 255, 255);
        }

        return null;
    }

    private static int getMaxColumnCount(Sheet sheet) {
        int maxColumns = 0;

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            if (row == null) {
                continue;
            }

            if (row.getLastCellNum() > maxColumns) {
                maxColumns = row.getLastCellNum();
            }
        }

        return maxColumns;
    }

    private static float[] getColumnWidths(Sheet sheet, int maxColumns) {
        float[] widths = new float[maxColumns];

        for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
            if (sheet.isColumnHidden(colIndex)) {
                widths[colIndex] = 0.1f;
                continue;
            }

            int excelWidth = sheet.getColumnWidth(colIndex);

            // Excel column width is based on 1/256 of character width.
            // This maps it roughly into iText relative table widths.
            widths[colIndex] = Math.max(1f, excelWidth / 256f);
        }

        return widths;
    }

    private static CellRangeAddress getMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);

            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }

        return null;
    }

    private static void markMergedCellsAsSkipped(
            Set<String> skippedCells,
            CellRangeAddress mergedRegion,
            int currentRow,
            int currentCol) {

        for (int row = mergedRegion.getFirstRow(); row <= mergedRegion.getLastRow(); row++) {
            for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
                if (row == currentRow && col == currentCol) {
                    continue;
                }

                skippedCells.add(row + ":" + col);
            }
        }
    }

    private static double getTableScaleRatio(
            Sheet sheet,
            int maxColumns,
            PageSize pageSize,
            double leftMargin,
            double rightMargin) {

        double totalExcelWidth = 0;

        for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
            if (sheet.isColumnHidden(colIndex)) {
                continue;
            }
            totalExcelWidth += sheet.getColumnWidth(colIndex) / 256.0;
        }

        double availablePdfWidth = pageSize.getWidth() - leftMargin - rightMargin;

        /*
         * Rough conversion:
         * Excel column width unit is character-based.
         * 7.0 is an approximation for points per Excel character width.
         */
        double estimatedExcelWidthInPoints = totalExcelWidth * 7.0;
        if (estimatedExcelWidthInPoints <= 0) {
            return 1.0;
        }
        double ratio = availablePdfWidth / estimatedExcelWidthInPoints;
        // Do not enlarge font if Excel is already smaller than PDF page.
        return Math.min(1.0, ratio);
    }
}
