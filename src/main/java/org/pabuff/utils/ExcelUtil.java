package org.pabuff.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

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

    public static void addSheet2(Workbook workbook,
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
        List<Map<String, Object>> excelChartList = null;
        if(excelMap != null && !excelMap.isEmpty()) {
            if(excelMap.get("excel_chart") != null){
                excelChartList = (List<Map<String, Object>>) excelMap.get("excel_chart");
                excelMap.remove("excel_chart");
            }
            excelStyleConfig = new ExcelStyleConfig(excelMap);
        }

        for (Map<String, Object> row : dataRows) {
            Row dataRow = sheet.createRow(rowCount++);
            columnCount = 0;
            for (Map.Entry<String, Object> entry : row.entrySet()) {

                if(excelStyleConfig != null){
                    if(excelStyleConfig.containsAnySuffix(entry.getKey())){
                        continue;
                    }
                }

                Cell cell = dataRow.createCell(columnCount++);

                if(excelStyleConfig != null){
                    Short color = null;
                    XSSFColor xssfColor = null;
                    FillPatternType fillPattern = null;
                    Boolean wrapText = null;
                    String fontName = null;
                    Short fontHeight = null;
                    Boolean isBold = null;
                    Short fontColor = null;
                    XSSFColor xssfFontColor = null;
                    Boolean isItalic = null;

                    if(excelStyleConfig.getCellColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellColorSuffix())) {
                        if( row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix()) instanceof XSSFColor){
                            xssfColor = (XSSFColor) row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
                        }
                        else{
                            color = (Short) row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
                        }
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
                        if( row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix()) instanceof XSSFColor){
                            xssfFontColor = (XSSFColor) row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
                        }
                        else{
                            fontColor = (Short) row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
                        }
                    }

                    Font font;
                    if (xssfFontColor != null) {
                        font = addFontStyle2(workbook, fontName, xssfFontColor, fontHeight, isBold, isItalic);
                    } else {
                        font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                    }

                    CellStyle style;
                    if (xssfColor != null) {
                        style = addCellStyle2(workbook, sheetName, xssfColor, fillPattern, wrapText, font);
                    } else {
                        style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
                    }

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

        if(excelChartList != null){
            int i = 1;
            for(Map<String, Object> excelChartMap : excelChartList){

                if(excelChartMap.get("chart_data") == null){
                    continue;
                }

                String chartName = "Chart" + i;
                String chartTitle = excelChartMap.get("chart_title") != null ? (String) excelChartMap.get("chart_title") : chartName;
                String x_axis_title = excelChartMap.get("x_axis_title") != null ? (String) excelChartMap.get("x_axis_title") : "X-Axis";
                String y_axis_title = excelChartMap.get("y_axis_title") != null ? (String) excelChartMap.get("y_axis_title") : "Y-Axis";
                int left = excelChartMap.get("left") != null ? (int) excelChartMap.get("left") : 0;
                int top = excelChartMap.get("top") != null ? (int) excelChartMap.get("top") : 0;
                int width = excelChartMap.get("width") != null ? (int) excelChartMap.get("width") : 5;
                int height = excelChartMap.get("height") != null ? (int) excelChartMap.get("height") : 10;
                Map<String, Object> chartData = (Map<String, Object>) excelChartMap.get("chart_data");
                ChartTypes chartType = (ChartTypes) excelChartMap.get("chart_type");

                if(chartData == null){
                    continue;
                }

                addChart(workbook, sheetName, chartName, chartTitle, x_axis_title, y_axis_title, chartType, left, top, width, height, chartData);
                i++;
            }
        }
    }

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

    public static void addRows2(Workbook workbook, String sheetName, List<LinkedHashMap<String, Object>> dataRows, Map<String, Object> excelMap) {

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        ExcelStyleConfig excelStyleConfig = null;
        List<Map<String, Object>> excelChartList = null;

        if(excelMap != null && !excelMap.isEmpty()) {
            if(excelMap.get("excel_chart") != null){
                excelChartList = (List<Map<String, Object>>) excelMap.get("excel_chart");
                excelMap.remove("excel_chart");
            }
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
                    if(excelStyleConfig.containsAnySuffix(entry.getKey())){
                        continue;
                    }
                }

                Cell cell = dataRow.createCell(columnCount++);

                if(excelStyleConfig != null){
                    Short color = null;
                    XSSFColor xssfColor = null;
                    FillPatternType fillPattern = null;
                    Boolean wrapText = null;
                    String fontName = null;
                    Short fontHeight = null;
                    Boolean isBold = null;
                    Short fontColor = null;
                    XSSFColor xssfFontColor = null;
                    Boolean isItalic = null;

                    if(excelStyleConfig.getCellColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellColorSuffix())) {
                        if( row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix()) instanceof XSSFColor){
                            xssfColor = (XSSFColor) row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
                        }
                        else{
                            color = (Short) row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
                        }
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
                        if( row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix()) instanceof XSSFColor){
                            xssfFontColor = (XSSFColor) row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
                        }
                        else{
                            fontColor = (Short) row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
                        }
                    }

                    Font font;
                    if (xssfFontColor != null) {
                        font = addFontStyle2(workbook, fontName, xssfFontColor, fontHeight, isBold, isItalic);
                    } else {
                        font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                    }

                    if (xssfColor != null) {
                        style = addCellStyle2(workbook, sheetName, xssfColor, fillPattern, wrapText, font);
                    } else {
                        style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
                    }

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

        if(excelChartList != null){
            int i = 1;
            for(Map<String, Object> excelChartMap : excelChartList){
                String chartName = "Chart" + i;
                String chartTitle = excelChartMap.get("chart_title") != null ? (String) excelChartMap.get("chart_title") : chartName;
                String x_axis_title = excelChartMap.get("x_axis_title") != null ? (String) excelChartMap.get("x_axis_title") : "X-Axis";
                String y_axis_title = excelChartMap.get("y_axis_title") != null ? (String) excelChartMap.get("y_axis_title") : "Y-Axis";
                int left = excelChartMap.get("left") != null ? (int) excelChartMap.get("left") : 0;
                int top = excelChartMap.get("top") != null ? (int) excelChartMap.get("top") : 0;
                int width = excelChartMap.get("width") != null ? (int) excelChartMap.get("width") : 5;
                int height = excelChartMap.get("height") != null ? (int) excelChartMap.get("height") : 10;
                Map<String, Object> chartData = (Map<String, Object>) excelChartMap.get("chart_data");
                ChartTypes chartType = (ChartTypes) excelChartMap.get("chart_type");

                if(chartData == null){
                    continue;
                }
                addChart(workbook, sheetName, chartName, chartTitle, x_axis_title, y_axis_title, chartType, left, top, width, height, chartData);
                i++;
            }
        }
    }

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

            if(entry.get("style") instanceof Map){
                Map<String, Object> styleMap = (Map<String, Object>) entry.get("style");
                Short color = null;
                FillPatternType fillPattern = null;
                Boolean wrapText = null;
                String fontName = null;
                Short fontHeight = null;
                Boolean isBold = null;
                Short fontColor = null;
                Boolean isItalic = null;

                if("key".equals(styleMap.get("style"))){
                    color = IndexedColors.YELLOW.getIndex();
                    fillPattern = FillPatternType.SOLID_FOREGROUND;
                    fontName = "Arial";
                    fontHeight = 12;
                    isBold = true;
                    isItalic = false;
                }

                if(styleMap.containsKey("cell_color")) {
                    color = (Short) styleMap.get("cell_color");
                }
                if(styleMap.containsKey("cell_fill_pattern")) {
                    fillPattern = (FillPatternType) styleMap.get("cell_fill_pattern");
                }
                if(styleMap.containsKey("cell_wrap_text")) {
                    wrapText = (Boolean) styleMap.get("cell_wrap_text");
                }

                if(styleMap.containsKey("font_name")) {
                    fontName = (String) styleMap.get("font_name");
                }
                if(styleMap.containsKey("font_height")) {
                    fontHeight = (Short) styleMap.get("font_height");
                }
                if(styleMap.containsKey("font_bold")) {
                    isBold = (Boolean) styleMap.get("font_bold");
                }
                if(styleMap.containsKey("font_italic")) {
                    isItalic = (Boolean) styleMap.get("font_italic");
                }
                if(styleMap.containsKey("font_color")) {
                    fontColor = (Short) styleMap.get("font_color");
                }

                Font font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
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

    public static CellStyle addCellStyle2(Workbook workbook, String sheetName, XSSFColor color, FillPatternType fillPatternType, Boolean setWrapText, Font font) {

        CellStyle style = workbook.createCellStyle();

        if(workbook instanceof XSSFWorkbook && color != null) {
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

    public static XSSFFont addFontStyle2(Workbook workbook, String fontName, XSSFColor fontColor, Short fontHeightInPoints, Boolean isBold, Boolean isItalic) {

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

    public static void addChart(Workbook workbook, String sheetName, String chartName, String chartTitle, String xAxisTitle, String yAxisTitle, ChartTypes chartTypes, int left, int top, int width, int height, Map<String, Object> dataMap) {
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
        XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
        if(sheet == null) {
            sheet = xssfWorkbook.createSheet(sheetName);
        }

        if(dataMap.get("x_data_ranges") == null || dataMap.get("series_data") == null){
            return;
        }

        int lastRow = sheet.getLastRowNum();
        int chartRow = lastRow + top;
        XSSFDrawing  drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, left, chartRow, left + width, chartRow + height);
        XSSFChart chart = drawing.createChart(anchor);

        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        List<String> xDataRanges = (List<String>) dataMap.get("x_data_ranges");
        XDDFDataSource<String> xData = XDDFDataSourcesFactory.fromArray(xDataRanges.toArray(new String[0]));

        // Y Data Source - dynamically adding series for each object
        List<Map<String, Object>> seriesList = (List<Map<String, Object>>) dataMap.get("series_data");

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(xAxisTitle != null ? xAxisTitle : "X Axis");

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(yAxisTitle != null ? yAxisTitle : "Y Axis");

        ChartTypes chartType = chartTypes != null ? chartTypes : ChartTypes.LINE;
        XDDFChartData data = chart.createData(chartType, bottomAxis, leftAxis);

        for(Map<String, Object> seriesData : seriesList){

            if(seriesData.get("series_name") == null || seriesData.get("y_data_ranges") == null){
                continue;
            }

            String seriesName = (String) seriesData.get("series_name");
            List<Double> yDataRanges = (List<Double>) seriesData.get("y_data_ranges");

            if(yDataRanges.isEmpty()){
                continue;
            }

            XDDFNumericalDataSource<Double> yData = XDDFDataSourcesFactory.fromArray(yDataRanges.toArray(new Double[0]));
            XDDFChartData.Series series = data.addSeries(xData, yData);
            series.setTitle(seriesName, null);

            if(series instanceof XDDFLineChartData.Series lineSeries){
                if(dataMap.get("line_marker_style") instanceof MarkerStyle){
                    lineSeries.setMarkerStyle((MarkerStyle) dataMap.get("line_marker_style"));
                }

                if(dataMap.get("line_marker_size") instanceof Short){
                    lineSeries.setMarkerSize((Short) dataMap.get("line_marker_size"));
                }

                if(dataMap.get("line_smooth") instanceof Boolean){
                    lineSeries.setSmooth((Boolean) dataMap.get("line_smooth"));
                }
            }

            if(series instanceof XDDFBarChartData.Series){
                XDDFBarChartData barChartData = (XDDFBarChartData) data;

                // Set bar direction with default value
                if(dataMap.get("bar_direction") instanceof BarDirection){
                    barChartData.setBarDirection((BarDirection) dataMap.get("bar_direction"));
                }

                // Set bar grouping with default value
                if(dataMap.get("bar_grouping") instanceof BarGrouping){
                    barChartData.setBarGrouping((BarGrouping) dataMap.get("bar_grouping"));
                }

                // Set bar overlap only if present
                if (dataMap.get("bar_overlap") instanceof Byte) {
                    barChartData.setOverlap((Byte) dataMap.get("bar_overlap"));
                }

                // Set gap width only if present
                if (dataMap.get("gap_width") instanceof Integer) {
                    barChartData.setGapWidth((Integer) dataMap.get("gap_width"));
                }

                // Set axis crosses with default value
                if(dataMap.get("axis_crosses") instanceof AxisCrosses){
                    leftAxis.setCrosses((AxisCrosses) dataMap.get("axis_crosses"));
                }
            }
        }

        chart.plot(data);
    }

}
