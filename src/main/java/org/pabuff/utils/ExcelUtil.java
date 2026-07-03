package org.pabuff.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

import static java.lang.Integer.parseInt;

//@Service
public class ExcelUtil {

//    private static class ExtractedCellStyle {
//        Short color;
//        XSSFColor xssfColor;
//        FillPatternType fillPattern;
//        Boolean wrapText;
//
//        String fontName;
//        Short fontHeight;
//        Boolean isBold;
//        Boolean isItalic;
//        Short fontColor;
//        XSSFColor xssfFontColor;
//
//        HorizontalAlignment horizontalAlignment;
//        VerticalAlignment verticalAlignment;
//    }

    public static Workbook createWorkbook(String sheetName, LinkedHashMap<String, Integer> headers, CellStyle headerStyle, XSSFFont headerFont) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetNm = "Sheet1";
        if(sheetName != null && !sheetName.isEmpty()) {
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

        // cap sheet name to 31 characters
        if(sheetName.length() > 31){
            sheetName = sheetName.substring(0,31);
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

                    Font font = null;
                    if(xssfFontColor != null || fontName != null || fontHeight != null || isBold != null || isItalic != null || fontColor != null) {
                        if (xssfFontColor != null) {
                            font = addFontStyle2(workbook, fontName, xssfFontColor, fontHeight, isBold, isItalic);
                        } else {
                            font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                        }
                    }

                    CellStyle style = null;
                    if(xssfColor != null || color != null || fillPattern != null || wrapText != null || font != null) {
                        if (xssfColor != null) {
                            style = addCellStyle2(workbook, sheetName, xssfColor, fillPattern, wrapText, font);
                        } else {
                            style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
                        }
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
                String chartPosition = excelChartMap.get("chart_position") != null ? (String) excelChartMap.get("chart_position") : "bottom";
                boolean overlap = excelChartMap.get("overlap") != null && (boolean) excelChartMap.get("overlap");
                boolean showLegend = excelChartMap.get("show_legend") == null || (boolean) excelChartMap.get("show_legend");

                if(chartData == null){
                    continue;
                }

                addChart(workbook, sheetName, chartName, chartTitle, x_axis_title, y_axis_title, chartType, chartPosition, overlap, left, top, width, height, showLegend, chartData);
                i++;
            }
        }
    }

    public static void addSheet3(Workbook workbook,
                                 String sheetName,
                                 LinkedHashMap<String, Integer> bodyHeader,
                                 List<LinkedHashMap<String, Object>> bodyDataRows,
                                 Map<String, Object> bodyHeaderStyleMap, Map<String, Object> bodyHeaderFontMap,
                                 Map<String, Object> excelMap) {
        // cap sheet name to 31 characters
        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        Sheet sheet = workbook.createSheet(sheetName);
        int rowCount = 0;

        Map<String, Object> styleMap = null;
        Map<String, Object> headerMap = null;
        Map<String, Object> summaryMap = null;
        String defaultBodyHeaderStyle = null;
        String defaultBodyDataStyle = null;

        if (excelMap != null) {
            styleMap = (Map<String, Object>) excelMap.get("excel_styles");
            headerMap = (Map<String, Object>) excelMap.get("excel_header");
            summaryMap = (Map<String, Object>) excelMap.get("excel_summary");
            defaultBodyHeaderStyle = (String) excelMap.get("default_body_header_style");
            defaultBodyDataStyle = (String) excelMap.get("default_body_data_style");
        }

        // Add excel header
        if (headerMap != null) {
            rowCount = renderExcelPart(sheet, workbook, headerMap, styleMap, rowCount);
        }

        // Create reusable body header style
        CellStyle bodyHeaderStyle = createBodyHeaderStyle(
                workbook, sheet, styleMap, defaultBodyHeaderStyle,
                bodyHeaderStyleMap, bodyHeaderFontMap
        );

        // Body header row
        Row bodyHeaderRow = getOrCreateRow(sheet, rowCount++);
        int columnCount = 0;

        for (Map.Entry<String, Integer> entry : bodyHeader.entrySet()) {
            Cell cell = bodyHeaderRow.createCell(columnCount);
            cell.setCellValue(entry.getKey());
            cell.setCellStyle(bodyHeaderStyle);

            if (entry.getValue() != null) {
                sheet.setColumnWidth(columnCount, entry.getValue());
            }

            columnCount++;
        }

        ExcelStyleConfig excelStyleConfig = null;
        List<Map<String, Object>> excelChartList = null;
        if(excelMap != null && !excelMap.isEmpty()) {
            if (excelMap.get("excel_chart") != null) {
                excelChartList = (List<Map<String, Object>>) excelMap.get("excel_chart");
                excelMap.remove("excel_chart");
            }
            excelStyleConfig = new ExcelStyleConfig(excelMap);
        }

        // Reusable cache to avoid creating too many duplicated styles
        Map<String, CellStyle> cellStyleCache = new HashMap<>();
        Map<String, Object> defaultBodyDataStyleMap = createBodyDataStyleMap(styleMap, defaultBodyDataStyle);

        // Body data rows
        for (Map<String, Object> row : bodyDataRows) {
            Row dataRow = sheet.createRow(rowCount++);
            columnCount = 0;
            for (Map.Entry<String, Object> entry : row.entrySet()) {

                if(excelStyleConfig != null && excelStyleConfig.containsAnySuffix(entry.getKey())){
                    continue;
                }

                int currentRow = rowCount - 1;
                int currentCol = columnCount++;
                Object cellValue = entry.getValue();
                CellStyle style = resolveBodyCellStyle(
                        workbook, sheet, row, entry.getKey(),
                        cellValue, excelStyleConfig, defaultBodyDataStyleMap, cellStyleCache
                );
//                CellStyle style = null;
//
//                if(excelStyleConfig != null){
//                    Short color = null;
//                    XSSFColor xssfColor = null;
//                    FillPatternType fillPattern = null;
//                    Boolean wrapText = null;
//
//                    String fontName = null;
//                    Short fontHeight = null;
//                    Boolean isBold = null;
//                    Short fontColor = null;
//                    XSSFColor xssfFontColor = null;
//                    Boolean isItalic = null;
//
//                    HorizontalAlignment horizontalAlignment = null;
//                    VerticalAlignment verticalAlignment = null;
//
//                    if(excelStyleConfig.getCellColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellColorSuffix())) {
//                        Object value = row.get(entry.getKey() + excelStyleConfig.getCellColorSuffix());
//                        if(value instanceof XSSFColor){
//                            xssfColor = (XSSFColor) value;
//                        } else {
//                            color = (Short) value;
//                        }
//                    }
//                    if(excelStyleConfig.getCellFillPatternSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellFillPatternSuffix())) {
//                        fillPattern = (FillPatternType) row.get(entry.getKey() + excelStyleConfig.getCellFillPatternSuffix());
//                    }
//                    if (excelStyleConfig.getCellWrapTextSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellWrapTextSuffix())) {
//                        wrapText = (Boolean) row.get(entry.getKey() + excelStyleConfig.getCellWrapTextSuffix());
//                    }
//                    if(excelStyleConfig.getCellHorizontalAlignmentSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellHorizontalAlignmentSuffix())) {
//                        Object value = row.get(entry.getKey() + excelStyleConfig.getCellHorizontalAlignmentSuffix());
//                        if(value instanceof HorizontalAlignment){
//                            horizontalAlignment = (HorizontalAlignment) value;
//                        } else if(value != null) {
//                            horizontalAlignment = HorizontalAlignment.valueOf(value.toString().toUpperCase());
//                        }
//                    }
//                    if(excelStyleConfig.getCellVerticalAlignmentSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getCellVerticalAlignmentSuffix())) {
//                        Object value = row.get(entry.getKey() + excelStyleConfig.getCellVerticalAlignmentSuffix());
//                        if(value instanceof VerticalAlignment){
//                            verticalAlignment = (VerticalAlignment) value;
//                        } else if(value != null) {
//                            verticalAlignment = VerticalAlignment.valueOf(value.toString().toUpperCase());
//                        }
//                    }
//                    if(excelStyleConfig.getFontNameSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontNameSuffix())) {
//                        fontName = (String) row.get(entry.getKey() + excelStyleConfig.getFontNameSuffix());
//                    }
//                    if(excelStyleConfig.getFontHeightInPointsSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontHeightInPointsSuffix())) {
//                        fontHeight = (Short) row.get(entry.getKey() + excelStyleConfig.getFontHeightInPointsSuffix());
//                    }
//                    if(excelStyleConfig.getFontBoldSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontBoldSuffix())) {
//                        isBold = (Boolean) row.get(entry.getKey() + excelStyleConfig.getFontBoldSuffix());
//                    }
//                    if(excelStyleConfig.getFontItalicSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontItalicSuffix())) {
//                        isItalic = (Boolean) row.get(entry.getKey() + excelStyleConfig.getFontItalicSuffix());
//                    }
//                    if(excelStyleConfig.getFontColorSuffix() != null && row.containsKey(entry.getKey() + excelStyleConfig.getFontColorSuffix())) {
//                        Object value = row.get(entry.getKey() + excelStyleConfig.getFontColorSuffix());
//                        if(value instanceof XSSFColor){
//                            xssfFontColor = (XSSFColor) value;
//                        } else {
//                            fontColor = (Short) value;
//                        }
//                    }
//
//                    // Default rule:
//                    // If the cell value is number and no explicit horizontal alignment is provided,
//                    // align it to right.
//                    if(horizontalAlignment == null && isNumericValue(cellValue)){
//                        horizontalAlignment = HorizontalAlignment.RIGHT;
//                    }
//
//                    Font font = null;
//                    if(xssfFontColor != null || fontName != null || fontHeight != null || isBold != null || isItalic != null || fontColor != null) {
//                        if (xssfFontColor != null) {
//                            font = addFontStyle2(workbook, fontName, xssfFontColor, fontHeight, isBold, isItalic);
//                        } else {
//                            font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
//                        }
//                    }
//
//                    if(xssfColor != null || color != null || fillPattern != null || wrapText != null || font != null || horizontalAlignment != null || verticalAlignment != null) {
//                        style = addCellStyle3(workbook, sheetName, color, xssfColor, fillPattern, wrapText, font, null, horizontalAlignment, verticalAlignment);
//                    }
////                    setCell(workbook, sheetName, rowCount-1, columnCount-1, entry.getValue(), null, style);
////                    continue;
//                } else if(isNumericValue(cellValue)) {
//                    style = workbook.createCellStyle();
//                    style.setAlignment(HorizontalAlignment.RIGHT);
//                }

                setCell(workbook, sheetName, currentRow, currentCol, cellValue, null, style);

//                if (entry.getValue() instanceof String) {
//                    cell.setCellValue((String) entry.getValue());
//                } else if (entry.getValue() instanceof Integer) {
//                    cell.setCellValue((Integer) entry.getValue());
//                } else if (entry.getValue() instanceof Long) {
//                    cell.setCellValue((Long) entry.getValue());
//                } else if (entry.getValue() instanceof Double) {
//                    cell.setCellValue((Double) entry.getValue());
//                } else if (entry.getValue() instanceof Float) {
//                    cell.setCellValue((Float) entry.getValue());
//                } else if (entry.getValue() instanceof Boolean) {
//                    cell.setCellValue((Boolean) entry.getValue());
//                } else if (entry.getValue() instanceof Date) {
//                    cell.setCellValue((Date) entry.getValue());
//                } else if (entry.getValue() instanceof LocalDateTime) {
//                    cell.setCellValue((LocalDateTime) entry.getValue());
//                }
            }
        }

        // Summary part
        if(summaryMap != null){
            rowCount = renderExcelPart(sheet, workbook, summaryMap, styleMap, rowCount);
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
                String chartPosition = excelChartMap.get("chart_position") != null ? (String) excelChartMap.get("chart_position") : "bottom";
                boolean overlap = excelChartMap.get("overlap") != null && (boolean) excelChartMap.get("overlap");
                boolean showLegend = excelChartMap.get("show_legend") == null || (boolean) excelChartMap.get("show_legend");

                if(chartData == null){
                    continue;
                }

                addChart(workbook, sheetName, chartName, chartTitle, x_axis_title, y_axis_title, chartType, chartPosition, overlap, left, top, width, height, showLegend, chartData);
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

                    Font font = null;
                    if(xssfFontColor != null || fontName != null || fontHeight != null || isBold != null || isItalic != null || fontColor != null) {
                        if (xssfFontColor != null) {
                            font = addFontStyle2(workbook, fontName, xssfFontColor, fontHeight, isBold, isItalic);
                        } else {
                            font = addFontStyle(workbook, fontName, fontColor, fontHeight, isBold, isItalic);
                        }
                    }

                    CellStyle style = null;
                    if(xssfColor != null || color != null || fillPattern != null || wrapText != null || font != null) {
                        if (xssfColor != null) {
                            style = addCellStyle2(workbook, sheetName, xssfColor, fillPattern, wrapText, font);
                        } else {
                            style = addCellStyle(workbook, sheetName, color, fillPattern, wrapText, font);
                        }
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
                String chartPosition = excelChartMap.get("chart_position") != null ? (String) excelChartMap.get("chart_position") : "bottom";
                boolean overlap = excelChartMap.get("overlap") != null && (boolean) excelChartMap.get("overlap");
                boolean showLegend = excelChartMap.get("show_legend") == null || (boolean) excelChartMap.get("show_legend");

                if(chartData == null){
                    continue;
                }
                addChart(workbook, sheetName, chartName, chartTitle, x_axis_title, y_axis_title, chartType, chartPosition, overlap, left, top, width, height, showLegend, chartData);
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
        setCellValue(cell, value);
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
            int row = parseInt(celName.substring(1)) - 1;
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

    public static CellStyle addCellStyle3(
            Workbook workbook, String sheetName, Short color,
            XSSFColor xssfColor, FillPatternType fillPatternType,
            Boolean setWrapText, Font font, Map<String, BorderStyle> border,
            HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment
            ) {

        CellStyle style = workbook.createCellStyle();

        if (xssfColor != null && workbook instanceof XSSFWorkbook) {
            style.setFillForegroundColor(xssfColor);
            style.setFillPattern(Objects.requireNonNullElse(fillPatternType, FillPatternType.SOLID_FOREGROUND));
        } else if (color != null) {
            style.setFillForegroundColor(color);
            style.setFillPattern(Objects.requireNonNullElse(fillPatternType, FillPatternType.SOLID_FOREGROUND));
        }
        if(setWrapText != null && setWrapText) {
            style.setWrapText(true);
        }
        if(font != null) {
            style.setFont(font);
        }
        if(border != null) {
            if(border.get("all") != null){
                style.setBorderTop(border.get("all"));
                style.setBorderBottom(border.get("all"));
                style.setBorderLeft(border.get("all"));
                style.setBorderRight(border.get("all"));
            }else{
                if(border.get("top") != null){
                    style.setBorderTop(border.get("top"));
                }
                if(border.get("bottom") != null){
                    style.setBorderBottom(border.get("bottom"));
                }
                if(border.get("left") != null){
                    style.setBorderLeft(border.get("left"));
                }
                if(border.get("right") != null){
                    style.setBorderRight(border.get("right"));
                }
            }
        }
        if(horizontalAlignment != null){
            style.setAlignment(horizontalAlignment);
        }
        if(verticalAlignment != null){
            style.setVerticalAlignment(verticalAlignment);
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

    public static void addChart(Workbook workbook, String sheetName, String chartName, String chartTitle, String xAxisTitle, String yAxisTitle, ChartTypes chartTypes, String chartPosition, boolean overlap, int left, int top, int width, int height, boolean showLegend, Map<String, Object> dataMap) {
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
        XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
        if(sheet == null) {
            sheet = xssfWorkbook.createSheet(sheetName);
        }

        if(dataMap.get("x_data_ranges") == null || dataMap.get("series_data") == null){
            return;
        }

        int lastRow = sheet.getLastRowNum();
        int lastColumn = sheet.getLastRowNum() > 0 ? sheet.getRow(0).getLastCellNum() : 0;
        int chartRow = 0;
        int chartColumn = 0;

        if(overlap){
            chartColumn = left;
            chartRow = top;
        }else{
            if("bottom".equals(chartPosition)){
                chartRow = lastRow + top + 1;
                chartColumn = left;
            }else if("right".equals(chartPosition)){
                chartColumn = left + lastColumn;
                chartRow = top;
            }
        }

        XSSFDrawing  drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, chartColumn, chartRow, chartColumn + width, chartRow + height);
        XSSFChart chart = drawing.createChart(anchor);

        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);

        if(showLegend){
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
        }

        List<String> xDataRanges = (List<String>) dataMap.get("x_data_ranges");
        XDDFDataSource<String> xData = XDDFDataSourcesFactory.fromArray(xDataRanges.toArray(new String[0]));

        // Y Data Source - dynamically adding series for each object
        List<Map<String, Object>> seriesList = (List<Map<String, Object>>) dataMap.get("series_data");

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(xAxisTitle != null ? xAxisTitle : "X Axis");

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(yAxisTitle != null ? yAxisTitle : "Y Axis");
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
//        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        // Set axis crosses with default value
        if(dataMap.get("axis_crosses") instanceof AxisCrosses axisCrosses){
            leftAxis.setCrosses(axisCrosses);
        }

        if(dataMap.get("axis_cross_between") instanceof AxisCrossBetween axisCrossBetween){
            leftAxis.setCrossBetween(axisCrossBetween);
        }

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

            if(dataMap.get("fill_properties") instanceof XDDFFillProperties fillProperties){
                series.setFillProperties(fillProperties);
            }

            if(dataMap.get("line_properties") instanceof XDDFLineProperties lineProperties){
                series.setLineProperties(lineProperties);
            }

           if(dataMap.get("shape_properties") instanceof XDDFShapeProperties shapeProperties){
               series.setShapeProperties(shapeProperties);
           }

            if(series instanceof XDDFLineChartData.Series lineSeries){
                XDDFLineChartData lineChartData = (XDDFLineChartData) data;

                if(seriesData.get("line_marker_style") instanceof MarkerStyle lineMarkerStyle){
                    lineSeries.setMarkerStyle(lineMarkerStyle);
                } else if(dataMap.get("line_marker_style") instanceof MarkerStyle lineMarkerStyle){
                    lineSeries.setMarkerStyle(lineMarkerStyle);
                }

                if(seriesData.get("line_marker_size") instanceof Short lineMarkerSize){
                    lineSeries.setMarkerSize(lineMarkerSize);
                }else if(dataMap.get("line_marker_size") instanceof Short lineMarkerSize){
                    lineSeries.setMarkerSize(lineMarkerSize);
                }

                if(seriesData.get("line_smooth") instanceof Boolean lineSmooth){
                    lineSeries.setSmooth(lineSmooth);
                }else if(dataMap.get("line_smooth") instanceof Boolean lineSmooth){
                    lineSeries.setSmooth(lineSmooth);
                }

                if(seriesData.get("vary_colors") instanceof Boolean varyColors){
                    lineChartData.setVaryColors(varyColors);
                }else if(dataMap.get("vary_colors") instanceof Boolean varyColors){
                    lineChartData.setVaryColors(varyColors);
                }

                if(seriesData.get("line_marker_style") instanceof MarkerStyle){
                    lineSeries.setMarkerStyle(MarkerStyle.CIRCLE);
                }
            }

            if(series instanceof XDDFBarChartData.Series barSeries){
                XDDFBarChartData barChartData = (XDDFBarChartData) data;

                // Set bar direction with default value
                if(seriesData.get("bar_direction") instanceof BarDirection barDirection){
                    barChartData.setBarDirection(barDirection);
                }else if(dataMap.get("bar_direction") instanceof BarDirection barDirection){
                    barChartData.setBarDirection(barDirection);
                }

                // Set bar grouping with default value
                if(seriesData.get("bar_grouping") instanceof BarGrouping barGrouping){
                    barChartData.setBarGrouping(barGrouping);
                }else if(dataMap.get("bar_grouping") instanceof BarGrouping barGrouping){
                    barChartData.setBarGrouping(barGrouping);
                }

                // Set bar overlap only if present
                if (seriesData.get("bar_overlap") instanceof Byte barOverlap) {
                    barChartData.setOverlap(barOverlap);
                }else if (dataMap.get("bar_overlap") instanceof Byte barOverlap) {
                    barChartData.setOverlap(barOverlap);
                }

                // Set gap width only if present
                if (seriesData.get("gap_width") instanceof Integer gapWidth) {
                    barChartData.setGapWidth(gapWidth);
                }else if (dataMap.get("gap_width") instanceof Integer gapWidth) {
                    barChartData.setGapWidth(gapWidth);
                }

                if(seriesData.get("vary_colors") instanceof Boolean varyColors){
                    barChartData.setVaryColors(varyColors);
                }else if(dataMap.get("vary_colors") instanceof Boolean varyColors){
                    barChartData.setVaryColors(varyColors);
                }

                if (seriesData.get("bar_color") instanceof XDDFColor barColor) {
                    XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(barColor);
                    barSeries.setFillProperties(fillProperties);
                } else if (dataMap.get("bar_color") instanceof XDDFColor barColor) {
                    XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(barColor);
                    barSeries.setFillProperties(fillProperties);
                }

                if(seriesData.get("error_bars") instanceof XDDFErrorBars errorBars){
                    barSeries.setErrorBars(errorBars);
                }else if(dataMap.get("error_bars") instanceof XDDFErrorBars errorBars){
                    barSeries.setErrorBars(errorBars);
                }

                if(seriesData.get("invert_if_negative") instanceof Boolean invertIfNegative){
                    barSeries.setInvertIfNegative(invertIfNegative);
                }else if(dataMap.get("invert_if_negative") instanceof Boolean invertIfNegative){
                    barSeries.setInvertIfNegative(invertIfNegative);
                }
            }

            if(series instanceof XDDFPieChartData.Series pieSeries){
                XDDFPieChartData pieChartData = (XDDFPieChartData) data;

                if(seriesData.get("vary_colors") instanceof Boolean varyColors){
                    pieChartData.setVaryColors(varyColors);
                }else if(dataMap.get("vary_colors") instanceof Boolean varyColors){
                    pieChartData.setVaryColors(varyColors);
                }

                if(seriesData.get("first_slice_angle") instanceof Integer firstSliceAngle){
                    pieChartData.setFirstSliceAngle(firstSliceAngle);
                }else if(dataMap.get("first_slice_angle") instanceof Integer firstSliceAngle){
                    pieChartData.setFirstSliceAngle(firstSliceAngle);
                }

                if(seriesData.get("explosion") instanceof Long explosion){
                    pieSeries.setExplosion(explosion);
                }else if(dataMap.get("explosion") instanceof Long explosion){
                    pieSeries.setExplosion(explosion);
                }

            }
        }

        chart.plot(data);
    }

    @SuppressWarnings("unchecked")
    public static int renderExcelPart(
            Sheet sheet,
            Workbook workbook,
            Map<String, Object> map,
            Map<String, Object> excelStyles,
            int rowCount
    ) {

        if (sheet == null || workbook == null || map == null) {
            return rowCount;
        }

        if (excelStyles == null) {
            excelStyles = new HashMap<>();
        }

        // Cache to store style used
        Map<String, CellStyle> styleCache = new HashMap<>();

        // Render main title
        Object titleObj = map.get("title");
        if (titleObj instanceof Map<?, ?> titleMapRaw) {
            Map<String, Object> titleMap = (Map<String, Object>) titleMapRaw;

            String titleText = (String) titleMap.get("text");
            String strStyle = (String) titleMap.get("style");
            if (titleText != null && !titleText.isEmpty()) {
                Row titleRow = getOrCreateRow(sheet, rowCount++);
                Cell cell = titleRow.createCell(0);
                cell.setCellValue(titleText);

                CellStyle titleStyle = getOrCreateStyle(styleCache, excelStyles, strStyle, sheet);
                if (titleStyle != null) {
                    cell.setCellStyle(titleStyle);
                }

                sheet.setColumnWidth(0, 10000);
            }
        }

        Object sectionsObj = map.get("sections");
        if (!(sectionsObj instanceof LinkedList<?> sections)) {
            return rowCount;
        }

        for (Object sectionObj : sections) {
            if (!(sectionObj instanceof Map<?, ?> sectionRaw)) {
                continue;
            }

            Map<String, Object> section = (Map<String, Object>) sectionRaw;
            rowCount = renderExcelSection(sheet, section, rowCount, excelStyles, styleCache);
        }
        return rowCount;
    }

    public static int renderKeyValueSection(Sheet sheet, Map<String, Object> section, int rowCount,
                                            Map<String, Object> styles, Map<String, CellStyle> styleCache) {
        Object dataObj = section.get("data");
        if (!(dataObj instanceof LinkedHashMap<?, ?> dataRaw)) {
            return rowCount;
        }

        String keyStyleName = (String) section.get("key_style");
        String valueStyleName = (String) section.get("value_style");

        CellStyle keyStyle = getOrCreateStyle(styleCache, styles, keyStyleName, sheet);
        CellStyle valueStyle = getOrCreateStyle(styleCache, styles, valueStyleName, sheet);

        LinkedHashMap<String, Object> data = (LinkedHashMap<String, Object>) dataRaw;
        for (Map.Entry<String, Object> entry : data.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            if(value == null){
                rowCount++;
                continue;
            }

            Row row = getOrCreateRow(sheet, rowCount++);
            Cell keyCell = row.createCell(0);
            keyCell.setCellValue(key);
            if (keyStyle != null) {
                keyCell.setCellStyle(keyStyle);
            }

            Cell valueCell = row.createCell(1);
            setCellValue(valueCell, value);
            if (valueStyle != null) {
                valueCell.setCellStyle(valueStyle);
            }
        }
        return rowCount;
    }

    // Summary section rendering method
    @SuppressWarnings("unchecked")
    public static int renderSummarySection(Sheet sheet,
                                           Map<String, Object> section,
                                           int rowCount,
                                           Map<String, Object> styles,
                                           Map<String, CellStyle> styleCache) {
        Object dataObj = section.get("data");
        if (!(dataObj instanceof LinkedHashMap<?,  ?> dataRaw)) {
            return rowCount;
        }

        LinkedHashMap<String, Object> data = (LinkedHashMap<String, Object>) dataRaw;

        Map<String, Object> stylesByColumn = new HashMap<>();
        Object stylesByColumnObj = section.get("styles_by_column");
        if (stylesByColumnObj instanceof Map<?, ?> rawStyleMap) {
            stylesByColumn = (Map<String, Object>) rawStyleMap;
        }

        // All Summary data implemented within 1 row
        Row row = getOrCreateRow(sheet, rowCount++);

        for (Map.Entry<String, Object> entry : data.entrySet()) {
            int columnIndex = Integer.parseInt(entry.getKey());
            if (columnIndex < 0) {
                continue;
            }

            Object value = entry.getValue();

            Cell cell = row.createCell(columnIndex);
            setCellValue(cell, value);

            String styleName = (String) stylesByColumn.get(entry.getKey());
            CellStyle cellStyle = getOrCreateStyle(styleCache, styles, styleName, sheet);

            if (cellStyle != null) {
                // clone the style so alignment change will not affect cached style
                CellStyle clonedStyle = sheet.getWorkbook().createCellStyle();
                clonedStyle.cloneStyleFrom(cellStyle);

                if (isNumericValue(value)) {
                    clonedStyle.setAlignment(HorizontalAlignment.RIGHT);
                }
                cell.setCellStyle(clonedStyle);
            } else if (isNumericValue(value)) {
                CellStyle numberStyle = sheet.getWorkbook().createCellStyle();
                numberStyle.setAlignment(HorizontalAlignment.RIGHT);
                cell.setCellStyle(numberStyle);
            }
        }

        return rowCount;
    }

    public static int renderExcelSection(Sheet sheet, Map<String, Object> section, int rowCount, Map<String, Object> styles,
                                         Map<String, CellStyle> styleCache) {
        String type = section.get("type") != null ? section.get("type").toString() : "";
        return switch (type) {
            case "key_value" -> renderKeyValueSection(sheet, section, rowCount, styles, styleCache);
            case "blank" -> rowCount + 1;
            case "summary" -> renderSummarySection(sheet, section, rowCount, styles, styleCache);
            default -> rowCount;
        };
    }

    public static CellStyle createCellStyleFromMap(Map<String, Object> style, Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();

        Short color = null;
        XSSFColor xssfColor = null;
        Object cellColorObj = style.get("cell_color");
        if (cellColorObj instanceof XSSFColor) {
            xssfColor = (XSSFColor) cellColorObj;
        } else if (cellColorObj instanceof Short) {
            color = (Short) cellColorObj;
        } else if (cellColorObj instanceof Integer) {
            color = ((Integer) cellColorObj).shortValue();
        }
        FillPatternType fillPattern = (FillPatternType) style.get("cell_fill_pattern");
        Boolean wrapText = (Boolean) style.get("cell_wrap_text");

        HorizontalAlignment horizontalAlignment = toHorizontalAlignment(style.get("cell_horizontal_alignment"));
        VerticalAlignment verticalAlignment = toVerticalAlignment(style.get("cell_vertical_alignment"));

        String fontName = (String) style.getOrDefault("font_name", null);
        Short fontHeight = (Short) style.get("font_height");
        Boolean isBold = (Boolean) style.get("font_bold");
        Short fontColor = (Short) style.get("font_color");
        Boolean isItalic = (Boolean) style.get("font_italic");

        Map<String, BorderStyle> borderStyleMap = (Map<String, BorderStyle>) style.get("border_style");

        Font font = null;
        if (fontName != null || isBold != null || isItalic != null || fontColor != null || xssfColor != null) {
            if (xssfColor != null) {
                font = addFontStyle2(workbook, fontName, xssfColor, fontHeight, isBold, isItalic);
            } else {
                font = addFontStyle(workbook, fontName, fontColor,
                        fontHeight, isBold, isItalic);
            }
        }

        return addCellStyle3(workbook, sheet.getSheetName(), color,
                null, fillPattern, wrapText, font, borderStyleMap,
                horizontalAlignment, verticalAlignment);
    }

    @SuppressWarnings("unchecked")
    private static CellStyle getOrCreateStyle(
            Map<String, CellStyle> styleCache,
            Map<String, Object> styles,
            String styleName,
            Sheet sheet
    ) {
        if (styleName == null || styleName.isBlank()) {
            return null;
        }

        if (styleCache.containsKey(styleName)) {
            return styleCache.get(styleName);
        }

        Object styleObj = styles.get(styleName);
        if (!(styleObj instanceof Map<?, ?> styleRaw)) {
            return null;
        }

        Map<String, Object> styleMap = (Map<String, Object>) styleRaw;
        CellStyle cellStyle = createCellStyleFromMap(styleMap, sheet);

        styleCache.put(styleName, cellStyle);
        return cellStyle;
    }

    private static Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row != null ? row : sheet.createRow(rowIndex);
    }

    private static HorizontalAlignment toHorizontalAlignment(Object value) {
        if (value == null) {
            return null;
        }
        if (value instanceof HorizontalAlignment) {
            return (HorizontalAlignment) value;
        }
        return HorizontalAlignment.valueOf(value.toString().toUpperCase());
    }

    private static VerticalAlignment toVerticalAlignment(Object value) {
        if (value == null) {
            return null;
        }
        if (value instanceof VerticalAlignment) {
            return (VerticalAlignment) value;
        }
        return VerticalAlignment.valueOf(value.toString().toUpperCase());
    }

    private static void setCellValue(Cell cell, Object value) {

        DecimalFormat decimalFormat = new DecimalFormat("#,##0.##");
        DecimalFormat decimalFormat2 = new DecimalFormat("#,##0.00");

        switch (value) {
            case null -> cell.setBlank();
            case BigDecimal bd -> cell.setCellValue(decimalFormat.format(bd));
            case Integer i -> cell.setCellValue(decimalFormat.format(i));
            case Long l -> cell.setCellValue(decimalFormat.format(l));
            case Double d -> cell.setCellValue(decimalFormat2.format(d));
            case Float f -> cell.setCellValue(decimalFormat2.format(f));
            case String str -> {
                String trimmed = str.trim();
                try {
                    BigDecimal bd = new BigDecimal(trimmed.replace(",", ""));
                    cell.setCellValue(decimalFormat.format(bd));
                } catch (NumberFormatException e) {
                    cell.setCellValue(str);
                }
            }
            case Boolean b -> cell.setCellValue(b);
            case Date date -> cell.setCellValue(date);
            case LocalDateTime localDateTime -> cell.setCellValue(localDateTime);
            default -> cell.setCellValue(value.toString());
        }
    }

    private static boolean isNumericValue(Object value) {
        if (value instanceof Number) {return true;}
        if (!(value instanceof String str)) {return false;}
        if (str.isBlank()) {return false;}

        try {
            new java.math.BigDecimal(str.trim().replace(",", ""));
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static CellStyle createBodyHeaderStyle(
            Workbook workbook, Sheet sheet,
            Map<String, Object> excelStyles,
            String defaultBodyHeaderStyle,
            Map<String, Object> bodyHeaderStyleMap,
            Map<String, Object> bodyHeaderFontMap
    ) {

        Map<String, Object> styleMap = new HashMap<>();

        // Default header style
        styleMap.put("cell_color", IndexedColors.YELLOW.getIndex());
        styleMap.put("cell_fill_pattern", FillPatternType.SOLID_FOREGROUND);
        styleMap.put("cell_wrap_text", true);
        styleMap.put("cell_horizontal_alignment", HorizontalAlignment.LEFT);
        styleMap.put("cell_vertical_alignment", VerticalAlignment.CENTER);

        styleMap.put("font_name", "Arial");
        styleMap.put("font_height", (short) 13);
        styleMap.put("font_bold", true);

        /*
         * If default_body_header_style exists,
         * use it as style name and resolve from excel_styles.
         */
        if (defaultBodyHeaderStyle != null && !defaultBodyHeaderStyle.isBlank()
                && excelStyles != null && excelStyles.get(defaultBodyHeaderStyle) instanceof Map<?, ?> rawStyleMap) {
            styleMap.putAll((Map<String, Object>) rawStyleMap);
        }

        // Override default cell style if provided
        if (bodyHeaderStyleMap != null) {
            styleMap.putAll(bodyHeaderStyleMap);
        }

        // Override default font style if provided
        if (bodyHeaderFontMap != null) {
            styleMap.putAll(bodyHeaderFontMap);
        }

        return createCellStyleFromMap(styleMap, sheet);
    }

    private static CellStyle resolveBodyCellStyle(
            Workbook workbook, Sheet sheet, Map<String, Object> row,
            String fieldName, Object cellValue, ExcelStyleConfig excelStyleConfig,
            Map<String, Object> defaultBodyDataStyleMap,
            Map<String, CellStyle> cellStyleCache
    ) {
        Map<String, Object> styleMap = new HashMap<>();

        // 1. Lowest priority: default style
        if (defaultBodyDataStyleMap != null) {
            styleMap.putAll(defaultBodyDataStyleMap);
        }

        // 2. Middle priority: numeric value alignment
        if (isNumericValue(cellValue)) {
            styleMap.put("cell_horizontal_alignment", HorizontalAlignment.RIGHT);
        }

        if (excelStyleConfig != null) {
            updateStyle(styleMap, "cell_color", row, fieldName, excelStyleConfig.getCellColorSuffix());
            updateStyle(styleMap, "cell_fill_pattern", row, fieldName, excelStyleConfig.getCellFillPatternSuffix());
            updateStyle(styleMap, "cell_wrap_text", row, fieldName, excelStyleConfig.getCellWrapTextSuffix());
            updateStyle(styleMap, "cell_horizontal_alignment", row, fieldName, excelStyleConfig.getCellHorizontalAlignmentSuffix());
            updateStyle(styleMap, "cell_vertical_alignment", row, fieldName, excelStyleConfig.getCellVerticalAlignmentSuffix());
            updateStyle(styleMap, "font_name", row, fieldName, excelStyleConfig.getFontNameSuffix());
            updateStyle(styleMap, "font_height", row, fieldName, excelStyleConfig.getFontHeightInPointsSuffix());
            updateStyle(styleMap, "font_bold", row, fieldName, excelStyleConfig.getFontBoldSuffix());
            updateStyle(styleMap, "font_italic", row, fieldName, excelStyleConfig.getFontItalicSuffix());
            updateStyle(styleMap, "font_color", row, fieldName, excelStyleConfig.getFontColorSuffix());
        }

        if (styleMap.isEmpty()) {
            return null;
        }

        String cacheKey = buildStyleCacheKey(styleMap);
        if (cellStyleCache.containsKey(cacheKey)) {
            return cellStyleCache.get(cacheKey);
        }

        CellStyle style = createCellStyleFromMap(styleMap, sheet);
        cellStyleCache.put(cacheKey, style);

        return style;
    }

    private static void updateStyle(
            Map<String, Object> styleMap, String styleKey, Map<String, Object> row,
            String fieldName, String suffix
    ) {
        if (suffix == null) {
            return;
        }

        String fullKey = fieldName + suffix;
        if (row.containsKey(fullKey)) {
            styleMap.put(styleKey, row.get(fullKey));
        }
    }

    private static String buildStyleCacheKey(Map<String, Object> styleMap) {
        return styleMap.entrySet()
                .stream()
                .sorted(Map.Entry.comparingByKey())
                .map(entry -> entry.getKey() + "=" + String.valueOf(entry.getValue()))
                .collect(Collectors.joining("|"));
    }

    private static CellStyle createBodyDataStyle(
            Workbook workbook, Sheet sheet,
            Map<String, Object> excelStyles,
            String defaultBodyDataStyle
    ) {
        Map<String, Object> styleMap = new HashMap<>();

        // Default body data style
        styleMap.put("cell_wrap_text", true);
        styleMap.put("cell_vertical_alignment", VerticalAlignment.CENTER);
        styleMap.put("font_name", "Arial");
        styleMap.put("font_height", (short) 11);
        styleMap.put("font_bold", false);

        /*
         * If default_body_data_style exists,
         * use it as style name and resolve from excel_styles.
         */
        if (defaultBodyDataStyle != null && !defaultBodyDataStyle.isBlank()
                && excelStyles != null
                && excelStyles.get(defaultBodyDataStyle) instanceof Map<?, ?> rawStyleMap) {
            styleMap.putAll((Map<String, Object>) rawStyleMap);
        }

        return createCellStyleFromMap(styleMap, sheet);
    }

    private static Map<String, Object> createBodyDataStyleMap(
            Map<String, Object> excelStyles,
            String defaultBodyDataStyle
    ) {
        Map<String, Object> styleMap = new HashMap<>();

        styleMap.put("cell_wrap_text", true);
        styleMap.put("cell_vertical_alignment", VerticalAlignment.CENTER);
        styleMap.put("font_name", "Arial");
        styleMap.put("font_height", (short) 11);
        styleMap.put("font_bold", false);

        if (defaultBodyDataStyle != null && !defaultBodyDataStyle.isBlank()
                && excelStyles != null
                && excelStyles.get(defaultBodyDataStyle) instanceof Map<?, ?> rawStyleMap) {
            styleMap.putAll((Map<String, Object>) rawStyleMap);
        }

        return styleMap;
    }
}
