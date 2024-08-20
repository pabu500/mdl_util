package org.pabuff.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;

//@Service
public class ExcelUtil {

    public static Workbook createWorkbook(String sheetName, LinkedHashMap<String, Integer> headers, CellStyle headerStyle, XSSFFont headerFont) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetNm = "Sheet1";
        if(!sheetName.isEmpty()) {
            sheetNm = sheetName;
        }
        Sheet sheet = workbook.createSheet(sheetNm);

        if(headerStyle == null) {
            headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        if(headerFont == null) {
            headerFont = workbook.createFont();
            headerFont.setFontName("Arial");
            headerFont.setFontHeightInPoints((short) 13);
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
        }

        int i = 0;
        Row headerRow = sheet.createRow(0);

        for(Map.Entry<String, Integer> entry : headers.entrySet()) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(entry.getKey());
            cell.setCellStyle(headerStyle);
            sheet.setColumnWidth(i, entry.getValue());
            i++;
        }
        return workbook;
    }

    public static Workbook createWorkbookEmpty() {
        XSSFWorkbook workbook = new XSSFWorkbook();

        return workbook;
    }

    public static void addSheet(Workbook workbook,
                                String sheetName,
                                LinkedHashMap<String, Integer> headers,
                                List<LinkedHashMap<String, Object>> dataRows,
                                CellStyle headerStyle, XSSFFont headerFont,
                                Map<String, Object> excelMap) {
        if(headerStyle == null) {
            headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setWrapText(true);
        }

        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;

        if(headerFont == null) {
            headerFont = xssfWorkbook.createFont();
            headerFont.setFontName("Arial");
            headerFont.setFontHeightInPoints((short) 13);
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
        }

        Sheet sheet = workbook.createSheet(sheetName);
        int rowCount = 0;
        Row headerRow = sheet.createRow(rowCount++);
        int columnCount = 0;
        for (Map.Entry<String, Integer> entry : headers.entrySet()) {
            Cell cell = headerRow.createCell(columnCount);
            cell.setCellValue(entry.getKey());
            cell.setCellStyle(headerStyle);
            sheet.setColumnWidth(columnCount, entry.getValue());
            columnCount++;
        }

        ExcelStyleConfig excelStyleConfig = null;
        if(excelMap != null && !excelMap.isEmpty()) {
            excelStyleConfig = new ExcelStyleConfig(excelMap);
        }

        for (Map<String, Object> row : dataRows) {
            Row dataRow = sheet.createRow(rowCount++);
            columnCount = 0;
            for (Map.Entry<String, Object> entry : row.entrySet()) {

                if(excelStyleConfig != null){
                    // Retrieve suffix values
                    String cellColorSuffix = excelStyleConfig.getCellColorSuffix();
                    String cellWrapTextSuffix = excelStyleConfig.getCellWrapTextSuffix();
                    String cellFillPatternSuffix = excelStyleConfig.getCellFillPatternSuffix();

                    String fontColorSuffix = excelStyleConfig.getFontColorSuffix();
                    String fontNameSuffix = excelStyleConfig.getFontNameSuffix();
                    String fontBoldSuffix = excelStyleConfig.getFontBoldSuffix();
                    String fontHeightInPointsSuffix = excelStyleConfig.getFontHeightInPointsSuffix();
                    String fontItalicSuffix = excelStyleConfig.getFontItalicSuffix();

                    // Check for non-empty suffixes
                    if ((cellColorSuffix != null && !cellColorSuffix.isEmpty() && entry.getKey().contains(cellColorSuffix)) ||
                            (cellFillPatternSuffix != null && !cellFillPatternSuffix.isEmpty() && entry.getKey().contains(cellFillPatternSuffix)) ||
                            (cellWrapTextSuffix != null && !cellWrapTextSuffix.isEmpty() && entry.getKey().contains(cellWrapTextSuffix)) ||
                            (fontColorSuffix != null && !fontColorSuffix.isEmpty() && entry.getKey().contains(fontColorSuffix)) ||
                            (fontNameSuffix != null && !fontNameSuffix.isEmpty() && entry.getKey().contains(fontNameSuffix)) ||
                            (fontBoldSuffix != null && !fontBoldSuffix.isEmpty() && entry.getKey().contains(fontBoldSuffix)) ||
                            (fontHeightInPointsSuffix != null && !fontHeightInPointsSuffix.isEmpty() && entry.getKey().contains(fontHeightInPointsSuffix)) ||
                            (fontItalicSuffix != null && !fontItalicSuffix.isEmpty() && entry.getKey().contains(fontItalicSuffix))) {
                        continue;
                    }
                }

                Cell cell = dataRow.createCell(columnCount++);

                if(excelStyleConfig != null){
                    Short color = null;
                    FillPatternType fillPattern = null;
                    Boolean wrapText = null;
                    String fontName = null;
                    Short fontHeight = null;
                    Boolean isBold = null;
                    Short fontColor = null;
                    Boolean isItalic = null;

                    if(excelStyleConfig.getCellColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellColorSuffix())) {
                        color = (Short) row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
                    }
                    if(excelStyleConfig.getCellFillPatternSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellFillPatternSuffix())) {
                        fillPattern = (FillPatternType) row.get(entry.getKey() + excelStyleConfig.getCellFillPatternSuffix());
                    }
                    if (excelStyleConfig.getCellWrapTextSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellWrapTextSuffix())) {
                        wrapText = (Boolean) row.get(entry.getKey() + excelStyleConfig.getCellWrapTextSuffix());
                    }

                    if(excelStyleConfig.getFontNameSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontNameSuffix())) {
                        fontName = (String) row.get(entry.getKey() + excelStyleConfig.getFontNameSuffix());
                    }
                    if(excelStyleConfig.getFontHeightInPointsSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontHeightInPointsSuffix())) {
                        fontHeight = (Short) row.get(entry.getKey() + excelStyleConfig.getFontHeightInPointsSuffix());
                    }
                    if(excelStyleConfig.getFontBoldSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontBoldSuffix())) {
                        isBold = (Boolean) row.get(entry.getKey() + excelStyleConfig.getFontBoldSuffix());
                    }
                    if(excelStyleConfig.getFontItalicSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontItalicSuffix())) {
                        isItalic = (Boolean) row.get(entry.getKey() + excelStyleConfig.getFontItalicSuffix());
                    }
                    if(excelStyleConfig.getFontColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontColorSuffix())) {
                        fontColor = (Short) row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
                    }

                        Font font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                        CellStyle style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
                        setCell(workbook, sheetName, rowCount-1, columnCount-1, entry.getValue(), null, style);
                        continue;

                }

                if (entry.getValue() instanceof String) {
                    cell.setCellValue((String) entry.getValue());
                } else if (entry.getValue() instanceof Integer) {
                    cell.setCellValue((Integer) entry.getValue());
                } else if (entry.getValue() instanceof Long) {
                    cell.setCellValue((Long) entry.getValue());
                } else if (entry.getValue() instanceof Double) {
                    cell.setCellValue((Double) entry.getValue());
                } else if (entry.getValue() instanceof Float) {
                    cell.setCellValue((Float) entry.getValue());
                } else if (entry.getValue() instanceof Boolean) {
                    cell.setCellValue((Boolean) entry.getValue());
                } else if (entry.getValue() instanceof Date) {
                    cell.setCellValue((Date) entry.getValue());
                } else if (entry.getValue() instanceof LocalDateTime) {
                    cell.setCellValue((LocalDateTime) entry.getValue());
                }
            }
        }
    }

    @Deprecated
    public static void addSheet(Workbook workbook,
                                String sheetName,
                                LinkedHashMap<String, Integer> headers,
                                List<LinkedHashMap<String, Object>> dataRows,
                                CellStyle headerStyle, XSSFFont headerFont
                                ) {
        if(headerStyle == null) {
            headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setWrapText(true);
        }

        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;

        if(headerFont == null) {
            headerFont = xssfWorkbook.createFont();
            headerFont.setFontName("Arial");
            headerFont.setFontHeightInPoints((short) 13);
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
        }

        Sheet sheet = workbook.createSheet(sheetName);
        int rowCount = 0;
        Row headerRow = sheet.createRow(rowCount++);
        int columnCount = 0;
        for (Map.Entry<String, Integer> entry : headers.entrySet()) {
            Cell cell = headerRow.createCell(columnCount);
            cell.setCellValue(entry.getKey());
            cell.setCellStyle(headerStyle);
            sheet.setColumnWidth(columnCount, entry.getValue());
            columnCount++;
        }

        for (Map<String, Object> row : dataRows) {
            Row dataRow = sheet.createRow(rowCount++);
            columnCount = 0;
            for (Map.Entry<String, Object> entry : row.entrySet()) {
                Cell cell = dataRow.createCell(columnCount++);

                if (entry.getValue() instanceof String) {
                    cell.setCellValue((String) entry.getValue());
                } else if (entry.getValue() instanceof Integer) {
                    cell.setCellValue((Integer) entry.getValue());
                } else if (entry.getValue() instanceof Long) {
                    cell.setCellValue((Long) entry.getValue());
                } else if (entry.getValue() instanceof Double) {
                    cell.setCellValue((Double) entry.getValue());
                } else if (entry.getValue() instanceof Float) {
                    cell.setCellValue((Float) entry.getValue());
                } else if (entry.getValue() instanceof Boolean) {
                    cell.setCellValue((Boolean) entry.getValue());
                } else if (entry.getValue() instanceof Date) {
                    cell.setCellValue((Date) entry.getValue());
                } else if (entry.getValue() instanceof LocalDateTime) {
                    cell.setCellValue((LocalDateTime) entry.getValue());
                }
            }
        }
    }

    public static void addRows(Workbook workbook, String sheetName, List<LinkedHashMap<String, Object>> dataRows, Map<String, Object> excelMap) {

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        ExcelStyleConfig excelStyleConfig = null;

        if(excelMap != null && !excelMap.isEmpty()) {
            excelStyleConfig = new ExcelStyleConfig(excelMap);
        }

        int rowCount = 1;
        for (Map<String, Object> row : dataRows) {
            Sheet sheet = workbook.getSheet(sheetName);
            if(sheet == null) {
                sheet = workbook.createSheet(sheetName);
            }
            rowCount = sheet.getLastRowNum() + 1;
            Row dataRow = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Map.Entry<String, Object> entry : row.entrySet()) {

                if(excelStyleConfig != null){
                    // Retrieve suffix values
                    String cellColorSuffix = excelStyleConfig.getCellColorSuffix();
                    String cellWrapTextSuffix = excelStyleConfig.getCellWrapTextSuffix();
                    String cellFillPatternSuffix = excelStyleConfig.getCellFillPatternSuffix();

                    String fontColorSuffix = excelStyleConfig.getFontColorSuffix();
                    String fontNameSuffix = excelStyleConfig.getFontNameSuffix();
                    String fontBoldSuffix = excelStyleConfig.getFontBoldSuffix();
                    String fontHeightInPointsSuffix = excelStyleConfig.getFontHeightInPointsSuffix();
                    String fontItalicSuffix = excelStyleConfig.getFontItalicSuffix();

                    // Check for non-empty suffixes
                    if ((cellColorSuffix != null && !cellColorSuffix.isEmpty() && entry.getKey().contains(cellColorSuffix)) ||
                            (cellFillPatternSuffix != null && !cellFillPatternSuffix.isEmpty() && entry.getKey().contains(cellFillPatternSuffix)) ||
                            (cellWrapTextSuffix != null && !cellWrapTextSuffix.isEmpty() && entry.getKey().contains(cellWrapTextSuffix)) ||
                            (fontColorSuffix != null && !fontColorSuffix.isEmpty() && entry.getKey().contains(fontColorSuffix)) ||
                            (fontNameSuffix != null && !fontNameSuffix.isEmpty() && entry.getKey().contains(fontNameSuffix)) ||
                            (fontBoldSuffix != null && !fontBoldSuffix.isEmpty() && entry.getKey().contains(fontBoldSuffix)) ||
                            (fontHeightInPointsSuffix != null && !fontHeightInPointsSuffix.isEmpty() && entry.getKey().contains(fontHeightInPointsSuffix)) ||
                            (fontItalicSuffix != null && !fontItalicSuffix.isEmpty() && entry.getKey().contains(fontItalicSuffix))) {
                        continue;
                    }
                }

                Cell cell = dataRow.createCell(columnCount++);

                if(excelStyleConfig != null){
                    Short color = null;
                    FillPatternType fillPattern = null;
                    Boolean wrapText = null;
                    String fontName = null;
                    Short fontHeight = null;
                    Boolean isBold = null;
                    Short fontColor = null;
                    Boolean isItalic = null;

                    if(excelStyleConfig.getCellColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellColorSuffix())) {
                        color = (Short) row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
                    }
                    if(excelStyleConfig.getCellFillPatternSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellFillPatternSuffix())) {
                        fillPattern = (FillPatternType) row.get(entry.getKey() + excelStyleConfig.getCellFillPatternSuffix());
                    }
                    if (excelStyleConfig.getCellWrapTextSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellWrapTextSuffix())) {
                        wrapText = (Boolean) row.get(entry.getKey() + excelStyleConfig.getCellWrapTextSuffix());
                    }

                    if(excelStyleConfig.getFontNameSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontNameSuffix())) {
                        fontName = (String) row.get(entry.getKey() + excelStyleConfig.getFontNameSuffix());
                    }
                    if(excelStyleConfig.getFontHeightInPointsSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontHeightInPointsSuffix())) {
                        fontHeight = (Short) row.get(entry.getKey() + excelStyleConfig.getFontHeightInPointsSuffix());
                    }
                    if(excelStyleConfig.getFontBoldSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontBoldSuffix())) {
                        isBold = (Boolean) row.get(entry.getKey() + excelStyleConfig.getFontBoldSuffix());
                    }
                    if(excelStyleConfig.getFontItalicSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontItalicSuffix())) {
                        isItalic = (Boolean) row.get(entry.getKey() + excelStyleConfig.getFontItalicSuffix());
                    }
                    if(excelStyleConfig.getFontColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontColorSuffix())) {
                        fontColor = (Short) row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
                    }

                    Font font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                    style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
                    setCell(workbook, sheetName, rowCount-1, columnCount-1, entry.getValue(), null, style);
                    continue;

                }

                if (entry.getValue() instanceof String) {
                    cell.setCellValue((String) entry.getValue());
                } else if (entry.getValue() instanceof Integer) {
                    cell.setCellValue((Integer) entry.getValue());
                } else if (entry.getValue() instanceof Long) {
                    cell.setCellValue((Long) entry.getValue());
                } else if (entry.getValue() instanceof Double) {
                    cell.setCellValue((Double) entry.getValue());
                } else if (entry.getValue() instanceof Float) {
                    cell.setCellValue((Float) entry.getValue());
                } else if (entry.getValue() instanceof Boolean) {
                    cell.setCellValue((Boolean) entry.getValue());
                } else if (entry.getValue() instanceof Date) {
                    cell.setCellValue((Date) entry.getValue());
                } else if (entry.getValue() instanceof LocalDateTime) {
                    cell.setCellValue((LocalDateTime) entry.getValue());
                }
            }
        }
    }

    @Deprecated
    public static void addRows(Workbook workbook, String sheetName, List<LinkedHashMap<String, Object>> dataRows) {

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        int rowCount = 1;
        for (Map<String, Object> row : dataRows) {
            Sheet sheet = workbook.getSheet(sheetName);
            if(sheet == null) {
                sheet = workbook.createSheet(sheetName);
            }
            rowCount = sheet.getLastRowNum() + 1;
            Row dataRow = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Map.Entry<String, Object> entry : row.entrySet()) {
                Cell cell = dataRow.createCell(columnCount++);

                if (entry.getValue() instanceof String) {
                    cell.setCellValue((String) entry.getValue());
                } else if (entry.getValue() instanceof Integer) {
                    cell.setCellValue((Integer) entry.getValue());
                } else if (entry.getValue() instanceof Long) {
                    cell.setCellValue((Long) entry.getValue());
                } else if (entry.getValue() instanceof Double) {
                    cell.setCellValue((Double) entry.getValue());
                } else if (entry.getValue() instanceof Float) {
                    cell.setCellValue((Float) entry.getValue());
                } else if (entry.getValue() instanceof Boolean) {
                    cell.setCellValue((Boolean) entry.getValue());
                } else if (entry.getValue() instanceof Date) {
                    cell.setCellValue((Date) entry.getValue());
                } else if (entry.getValue() instanceof LocalDateTime) {
                    cell.setCellValue((LocalDateTime) entry.getValue());
                }
            }
        }
    }

    public static void setCell(Workbook workbook, String sheetName, int row, int col, Object value, Double width, CellStyle style) {
        Sheet sheet = workbook.getSheet(sheetName);
        if(sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        Row dataRow = sheet.getRow(row);
        if(dataRow == null) {
            dataRow = sheet.createRow(row);
        }
        Cell cell = dataRow.getCell(col);
        if(cell == null) {
            cell = dataRow.createCell(col);
        }
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        }
        if(width != null) {
            sheet.setColumnWidth(col, width.intValue());
        }
        if(style != null) {
            cell.setCellStyle(style);
        }
    }

    public static void addPatch(Workbook workbook, String sheetName, List<Map<String, Object>> patch) {
        for (Map<String, Object> entry : patch) {
            String celName = (String) entry.get("cell");
            //turn "A1" to "1" and 1
            int col = celName.charAt(0) - 'A';
            int row = Integer.parseInt(celName.substring(1)) - 1;
            Object value = entry.get("value");
            Double width = null;
            if(entry.get("width") != null){
                width = MathUtil.ObjToDouble(entry.get("width"));
            }

            CellStyle style = null;
            if("key".equals(entry.get("style"))){
                style = workbook.createCellStyle();
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
                XSSFFont keyFont = xssfWorkbook.createFont();
                keyFont.setFontName("Arial");
                keyFont.setFontHeightInPoints((short) 12);
                keyFont.setBold(true);

                style.setFont(keyFont);
            }
            setCell(workbook, sheetName, row, col, value, width, style);
        }
    }

    public static void saveWorkbook(Workbook workbook, String fileName) {
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(fileName);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        try {
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static CellStyle addCellStyle(Workbook workbook, String sheetName, Short color, FillPatternType fillPatternType, Boolean setWrapText, Font font) {

        CellStyle style = workbook.createCellStyle();

        if(color != null) {
            style.setFillForegroundColor(color);
            style.setFillPattern(Objects.requireNonNullElse(fillPatternType, FillPatternType.SOLID_FOREGROUND));
        }
        if(setWrapText != null && setWrapText) {
            style.setWrapText(true);
        }
        if(font != null) {
            style.setFont(font);
        }

        return style;
    }

    public static XSSFFont addFontStyle(Workbook workbook, String fontName, Short fontColor, Short fontHeightInPoints, Boolean isBold, Boolean isItalic) {

        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
        XSSFFont font = xssfWorkbook.createFont();

        if(fontName != null) {
            font.setFontName(fontName);
        }
        if(fontHeightInPoints != null) {
            font.setFontHeightInPoints(fontHeightInPoints);
        }
        if(isBold != null && isBold) {
            font.setBold(true);
        }
        if(fontColor != null) {
            font.setColor(fontColor);
        }
        if(isItalic != null && isItalic) {
            font.setItalic(true);
        }

        return font;
    }
}
