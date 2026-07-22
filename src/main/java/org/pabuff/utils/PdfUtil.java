package org.pabuff.utils;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import java.awt.Color;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.util.*;
import java.util.List;

public class PdfUtil {

    public static void addPage(
            Document document,
            LinkedHashMap<String, Integer> bodyHeader,
            List<LinkedHashMap<String, Object>> bodyDataRows,
            Map<String, Object> pdfMap,
            Rectangle pageSize,
            float leftMargin,
            float rightMargin
    ) throws DocumentException {

        if (bodyHeader == null || bodyHeader.isEmpty()) {
            return;
        }

        if (bodyDataRows == null) {
            bodyDataRows = new ArrayList<>();
        }

        Map<String, Object> styleMap = null;
        Map<String, Object> headerMap = null;
        Map<String, Object> summaryMap = null;
        String defaultBodyHeaderStyle = null;
        String defaultBodyDataStyle = null;

        if (pdfMap != null) {
            styleMap = (Map<String, Object>) pdfMap.get("pdf_style");
            headerMap = (Map<String, Object>) pdfMap.get("pdf_header");
            summaryMap = (Map<String, Object>) pdfMap.get("pdf_summary");
            defaultBodyDataStyle = (String) (pdfMap.getOrDefault("default_body_data_style", null));
            defaultBodyHeaderStyle = (String) (pdfMap.getOrDefault("default_body_header_style", null));
        }

        if (styleMap == null) {
            styleMap = new HashMap<>();
        }

        int columnCount = bodyHeader.size();
        float[] columnWidths = toPdfColumnWidths(bodyHeader);
        float tableScaleRatio = getTableScaleRatio(bodyHeader, pageSize, leftMargin, rightMargin);

        if (headerMap != null) {
            renderPdfPart(document, headerMap, styleMap, columnCount, columnWidths, tableScaleRatio);
            document.add(new Paragraph(" "));
        }

        PdfPTable table = new PdfPTable(columnWidths);
        table.setWidthPercentage(100);
        table.setSplitLate(false);
        table.setSplitRows(true);

        PdfStyle bodyHeaderStyle = createBodyHeaderStyle(styleMap, defaultBodyHeaderStyle, tableScaleRatio);
        for (String headerName : bodyHeader.keySet()) {
            PdfPCell cell = createCell(headerName, bodyHeaderStyle, tableScaleRatio);
            table.addCell(cell);
        }

        // Repeat body header row when table continues to next PDF page
        if(!bodyDataRows.isEmpty()){
            table.setHeaderRows(1);
        }

        CellStyleConfig styleConfig = new CellStyleConfig(pdfMap);
        PdfStyle defaultBodyStyle = getPdfStyle(styleMap, defaultBodyDataStyle, tableScaleRatio);

        for (Map<String, Object> row : bodyDataRows) {
            for (String key : bodyHeader.keySet()) {
                Object value = row.get(key);
                PdfStyle pdfStyle = buildCellStyleFromRow(row, key, value, styleConfig, defaultBodyStyle, tableScaleRatio);
                PdfPCell cell = createCell(value, pdfStyle, tableScaleRatio);
                table.addCell(cell);
            }
        }

        document.add(table);

        if (summaryMap != null) {
            document.add(new Paragraph(" "));
            renderPdfPart(document, summaryMap, styleMap, columnCount, columnWidths, tableScaleRatio);
        }
    }

    private static void renderPdfPart(
            Document document,
            Map<String, Object> partMap,
            Map<String, Object> styles,
            int columnCount,
            float[] columnWidths,
            float tableScaleRatio
    ) throws DocumentException {

        Object titleObj = partMap.get("title");
        if (titleObj instanceof Map<?, ?> titleRaw) {
            Map<String, Object> titleMap = (Map<String, Object>) titleRaw;
            String text = Objects.toString(titleMap.get("text"), "");
            String styleName = Objects.toString(titleMap.get("style"), "");

            if (!text.isBlank()) {
                PdfStyle style = getPdfStyle(styles, styleName, tableScaleRatio);
                Paragraph p = new Paragraph(text, style.font);
                p.setAlignment(style.horizontalAlignment);
                p.setSpacingAfter(6f);
                document.add(p);
            }
        }

        Object sectionsObj = partMap.get("sections");
        if (!(sectionsObj instanceof List<?> sections)) {
            return;
        }

        for (Object sectionObj : sections) {
            if (!(sectionObj instanceof Map<?, ?> sectionRaw)) {
                continue;
            }

            Map<String, Object> section = (Map<String, Object>) sectionRaw;
            String type = Objects.toString(section.get("type"), "");

            switch (type) {
                case "key_value" -> renderKeyValueSection(document, section, styles, tableScaleRatio);
                case "summary" -> renderSummarySection(document, section, styles, columnCount, columnWidths, tableScaleRatio);
                case "blank" -> document.add(new Paragraph(" "));
                default -> {
                }
            }
        }
    }

    private static void renderKeyValueSection(
            Document document,
            Map<String, Object> section,
            Map<String, Object> styles,
            float tableScaleRatio
    ) throws DocumentException {

        Object dataObj = section.get("data");
        if (!(dataObj instanceof LinkedHashMap<?, ?> dataRaw)) {
            return;
        }

        String keyStyleName = Objects.toString(section.get("key_style"), "");
        String valueStyleName = Objects.toString(section.get("value_style"), "");
        PdfStyle keyStyle = getPdfStyle(styles, keyStyleName, tableScaleRatio);
        PdfStyle valueStyle = getPdfStyle(styles, valueStyleName, tableScaleRatio);

        int widthPercentage = section.get("width_percentage") instanceof Number n ? n.intValue() : 100;
        float keyWidth = section.get("key_width") instanceof Number n ? n.floatValue() : 3f;
        float valueWidth = section.get("value_width") instanceof Number n ? n.floatValue() : 9f;

        PdfPTable table = new PdfPTable(new  float[]{keyWidth, valueWidth});
        table.setWidthPercentage(widthPercentage);
        table.setHorizontalAlignment(Element.ALIGN_LEFT);
        table.setSplitLate(false);
        table.setSplitRows(true);

        for (Map.Entry<?, ?> entry : dataRaw.entrySet()) {
            table.addCell(createCell(entry.getKey(), keyStyle, tableScaleRatio));
            table.addCell(createCell(entry.getValue(), valueStyle, tableScaleRatio));
        }

        document.add(table);
    }

    private static void renderSummarySection(
            Document document,
            Map<String, Object> section,
            Map<String, Object> styles,
            int columnCount,
            float[] columnWidths,
            float tableScaleRatio
    ) throws DocumentException {

        Map<String, Object> data = (Map<String, Object>) section.get("data");
        if(data == null || data.isEmpty()){
            return;
        }

        Map<String, Object> stylesByCol = (Map<String, Object>) section.get("styles_by_column");
        PdfStyle defaultStyle = getPdfStyle(styles, Objects.toString(section.get("style"), ""), tableScaleRatio);

//        Map<Integer, Object> rowValueMap = new HashMap<>();
//        int maxRowIndex = -1;

        PdfPTable table = new PdfPTable(columnWidths);
        table.setWidthPercentage(100);
        table.setSplitLate(false);
        table.setSplitRows(true);

        /*
         * all summary data is rendered into one row.
         *
         * data key = target column index
         * example:
         * data.put("0", "Total");
         * data.put("10", "100.00");
         */
        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
            String colKey = String.valueOf(colIndex);
            Object value = data.get(colKey);
            String styleName = null;

            if(stylesByCol != null){
                Object styleNameObj = stylesByCol.get(colKey);
                if (styleNameObj != null) {
                    styleName = styleNameObj.toString();
                }
            }

            PdfStyle cellStyle = styleName != null
                    ? getPdfStyle(styles, styleName, tableScaleRatio)
                    : extractPdfStyle(defaultStyle);

            boolean hasHorizontalAlignment = false;
            if(styles != null && styleName != null && !styleName.isBlank()){
                Object styleObj = styles.get(styleName);
                if ((styleObj instanceof Map<?, ?> rawStyleMap)) {
                    hasHorizontalAlignment = rawStyleMap.get("cell_horizontal_alignment") != null;
                }
            }

            // Apply numeric alignment only when column-specific style is not provided
            if (!hasHorizontalAlignment && isNumericValue(value)) {
                cellStyle.horizontalAlignment = Element.ALIGN_RIGHT;
            }

            PdfPCell cell = createCell(value, cellStyle, tableScaleRatio);
            table.addCell(cell);
        }

        document.add(table);
    }

    private static PdfPCell createCell(Object value, PdfStyle style, float tableScaleRatio) {
        String text = formatValue(value);

        PdfPCell cell = new PdfPCell(new Phrase(text, style.font));
        cell.setHorizontalAlignment(style.horizontalAlignment);
        cell.setVerticalAlignment(style.verticalAlignment);
        cell.setPadding(Math.max(1f, 3f * tableScaleRatio));
        cell.setNoWrap(!style.wrapText);

        // Improve automatic height calculation around the font.
        cell.setUseAscender(true);
        cell.setUseDescender(true);

        if (style.backgroundColor != null) {
            cell.setBackgroundColor(style.backgroundColor);
        }

        if (style.noBorder) {
            cell.setBorder(Rectangle.NO_BORDER);
        } else {
            cell.setBorder(Rectangle.BOX);
            cell.setBorderWidth(style.borderWidth);
        }

        return cell;
    }

    private static PdfStyle buildCellStyleFromRow(
            Map<String, Object> row,
            String key,
            Object value,
            CellStyleConfig config,
            PdfStyle defaultStyle,
            float tableScaleRatio
    ) {
        PdfStyle style = extractPdfStyle(defaultStyle);

        if (config == null) {
            if (isNumericValue(value)) {
                style.horizontalAlignment = Element.ALIGN_RIGHT;
            }
            return style;
        }

        Object cellColor = getSuffixValue(row, key, config.cellColorSuffix);
        Object wrapText = getSuffixValue(row, key, config.cellWrapTextSuffix);
        Object horizontalAlignment = getSuffixValue(row, key, config.cellHorizontalAlignmentSuffix);
        Object verticalAlignment = getSuffixValue(row, key, config.cellVerticalAlignmentSuffix);
        Object fontName = getSuffixValue(row, key, config.fontNameSuffix);
        Object fontHeight = getSuffixValue(row, key, config.fontHeightSuffix);
        Object fontBold = getSuffixValue(row, key, config.fontBoldSuffix);
        Object fontItalic = getSuffixValue(row, key, config.fontItalicSuffix);
        Object fontColor = getSuffixValue(row, key, config.fontColorSuffix);

        if (cellColor != null) {
            style.backgroundColor = toColor(cellColor);
        }

        if (wrapText instanceof Boolean b) {
            style.wrapText = b;
        }

        if (horizontalAlignment != null) {
            style.horizontalAlignment = toHorizontalAlignment(horizontalAlignment);
        } else if (isNumericValue(row.get(key))) {
            style.horizontalAlignment = Element.ALIGN_RIGHT;
        }

        if (verticalAlignment != null) {
            style.verticalAlignment = toVerticalAlignment(verticalAlignment);
        }

        String resolvedFontName = fontName != null ? fontName.toString() : FontFactory.HELVETICA;
        float baseFontSize = fontHeight instanceof Number n ? n.floatValue() : style.font.getSize() / style.scaleRate;
        float resolvedFontSize = Math.max(5.5f, baseFontSize * style.scaleRate);

        int fontStyle = Font.NORMAL;
        if (Boolean.TRUE.equals(fontBold)) {
            fontStyle |= Font.BOLD;
        }
        if (Boolean.TRUE.equals(fontItalic)) {
            fontStyle |= Font.ITALIC;
        }

        Color resolvedFontColor = fontColor != null ? toColor(fontColor) : Color.BLACK;
        style.font = new Font(FontFactory.getFont(resolvedFontName).getFamily(), resolvedFontSize, fontStyle, resolvedFontColor);

        return style;
    }

    private static PdfStyle getPdfStyle(
            Map<String, Object> styles,
            String styleName,
            float tableScaleRatio
    ) {
        PdfStyle result = PdfStyle.defaultBody(tableScaleRatio);

        if (styles == null || styleName == null || styleName.isBlank()) {
            return result;
        }

        Object styleObj = styles.get(styleName);
        if (!(styleObj instanceof Map<?, ?> raw)) {
            return result;
        }

        Map<String, Object> map = (Map<String, Object>) raw;
        Object cellColor = map.get("cell_color");
        Object wrapText = map.get("cell_wrap_text");
        Object horizontalAlignment = map.get("cell_horizontal_alignment");
        Object verticalAlignment = map.get("cell_vertical_alignment");
        Object fontName = map.get("font_name");
        Object fontHeight = map.get("font_height");
        Object fontBold = map.get("font_bold");
        Object fontItalic = map.get("font_italic");
        Object fontColor = map.get("font_color");
        Object borderStyle = map.get("border_style");
        Object borderColor = map.get("border_color");
        Object scaleRate = map.get("scale_rate");

        if (cellColor != null) {
            result.backgroundColor = toColor(cellColor);
        }

        if (wrapText instanceof Boolean b) {
            result.wrapText = b;
        }

        if (horizontalAlignment != null) {
            result.horizontalAlignment = toHorizontalAlignment(horizontalAlignment);
        }

        if (verticalAlignment != null) {
            result.verticalAlignment = toVerticalAlignment(verticalAlignment);
        }

        String resolvedFontName = fontName != null ? fontName.toString() : FontFactory.HELVETICA;

        double scaleRateValue = scaleRate instanceof Number n ? n.doubleValue() : 1.0;
        result.scaleRate = (float) (scaleRateValue * tableScaleRatio);
        float resolvedFontSize = fontHeight instanceof Number n ? n.floatValue() : 9f;
        resolvedFontSize = Math.max(5.5f, resolvedFontSize * result.scaleRate);

        int resolvedFontStyle = Font.NORMAL;
        if (Boolean.TRUE.equals(fontBold)) {
            resolvedFontStyle |= Font.BOLD;
        }
        if (Boolean.TRUE.equals(fontItalic)) {
            resolvedFontStyle |= Font.ITALIC;
        }

        Color resolvedFontColor = fontColor != null ? toColor(fontColor) : Color.BLACK;
        result.font = new Font(
                FontFactory.getFont(resolvedFontName).getFamily(),
                resolvedFontSize,
                resolvedFontStyle,
                resolvedFontColor
        );

        if (borderColor != null) {
            Color color = toColor(borderColor);
            if (color != null) {
                result.borderColor = color;
            }
        }
        if (borderStyle instanceof Map<?, ?> rawBorderMap) {
            applyBorderStyle(result, (Map<String, Object>) rawBorderMap);
        }

        return result;
    }

    private static Object getSuffixValue(Map<String, Object> row, String key, String suffix) {
        if (row == null || key == null || suffix == null || suffix.isBlank()) {
            return null;
        }
        return row.get(key + suffix);
    }

    private static float[] toPdfColumnWidths(LinkedHashMap<String, Integer> bodyHeader) {
        float[] widths = new float[bodyHeader.size()];
        int i = 0;

        for (Integer contentWidth : bodyHeader.values()) {
            if (contentWidth == null) {
                widths[i++] = 10f;
            } else {
                widths[i++] = Math.max(1f, contentWidth / 256f);
            }
        }

        return widths;
    }

    private static float getTableScaleRatio(
            LinkedHashMap<String, Integer> bodyHeader,
            Rectangle pageSize,
            float leftMargin,
            float rightMargin
    ) {
        double totalContentWidth = 0;

        for (Integer width : bodyHeader.values()) {
            if (width != null) {
                totalContentWidth += width / 256.0;
            }
        }

        double availablePdfWidth = pageSize.getWidth() - leftMargin - rightMargin;
        double estimatedContentWidthInPoints = totalContentWidth * 7.0;

        if (estimatedContentWidthInPoints <= 0) {
            return 1.0f;
        }

        return (float) Math.min(1.0, availablePdfWidth / estimatedContentWidthInPoints);
    }

    private static String formatValue(Object value) {
        if (value == null) {
            return "";
        }

//        DecimalFormat decimalFormat = new DecimalFormat("#,##0.##");
//        DecimalFormat decimalFormat2 = new DecimalFormat("#,##0.00");
//        DecimalFormat decimalFormat2 = new DecimalFormat("#,##0.################");

        return switch (value) {
            case BigDecimal bd -> MathUtil.formatNumber(bd);
            case Integer i -> MathUtil.formatNumber(BigDecimal.valueOf(i));
            case Long l -> MathUtil.formatNumber(BigDecimal.valueOf(l));
            case Double d -> MathUtil.formatNumber(BigDecimal.valueOf(d));
            case Float f -> MathUtil.formatNumber(BigDecimal.valueOf(f));
            case LocalDateTime ldt -> ldt.toString();
            case String str -> {
                // Strings starting with ' are treated as text.
                String trimmed = str.trim();
                if (trimmed.startsWith("'")) {
                    yield trimmed.substring(1);
                }

                try {
                    BigDecimal bd = new BigDecimal(trimmed.replace(",", ""));
                    yield MathUtil.formatNumber(bd);
                } catch (NumberFormatException e) {
                    yield str;
                }
            }
            default -> String.valueOf(value);
        };
    }

    private static Color toColor(Object value) {
        if (value == null) {
            return null;
        }

        if (value instanceof Color color) {
            return color;
        }

        if (value instanceof Number n) {
            short index = n.shortValue();

            // Common POI IndexedColors mapping.
            if (index == 13) { // YELLOW
                return Color.YELLOW;
            }
            if (index == 22) { // GREY_25_PERCENT
                return new Color(217, 217, 217);
            }
            if (index == 55) { // GREY_40_PERCENT
                return new Color(153, 153, 153);
            }
            if (index == 23) { // GREY_50_PERCENT
                return new Color(128, 128, 128);
            }
            if (index == 48) { // LIGHT_BLUE
                return new Color(173, 216, 230);
            }
            if (index == 10) { // RED
                return Color.RED;
            }
            if (index == 12) { // BLUE
                return Color.BLUE;
            }
            if (index == 17) { // GREEN
                return new Color(0, 128, 0);
            }
            if (index == 9) { // WHITE
                return Color.WHITE;
            }

            return null;
        }

        if (value instanceof String s) {
            String hex = s.trim();

            if (hex.startsWith("#")) {
                hex = hex.substring(1);
            }

            if (hex.length() == 6) {
                return new Color(
                        Integer.parseInt(hex.substring(0, 2), 16),
                        Integer.parseInt(hex.substring(2, 4), 16),
                        Integer.parseInt(hex.substring(4, 6), 16)
                );
            }
        }

        return null;
    }

    private static class PdfStyle {
        Font font;
        Color backgroundColor;
        int horizontalAlignment = Element.ALIGN_LEFT;
        int verticalAlignment = Element.ALIGN_MIDDLE;
        boolean wrapText = true;
        boolean noBorder = true;

        Color borderColor = Color.BLACK;
        float borderWidth = 0.5f;

        boolean borderTop = false;
        boolean borderBottom = false;
        boolean borderLeft = false;
        boolean borderRight = false;

        float borderTopWidth = 0f;
        float borderBottomWidth = 0f;
        float borderLeftWidth = 0f;
        float borderRightWidth = 0f;

        float scaleRate = 1.0f;

        static PdfStyle defaultBody(float scaleRatio) {
            PdfStyle style = new PdfStyle();
            style.font = new Font(Font.HELVETICA, Math.max(5.5f, 9f * scaleRatio), Font.NORMAL, Color.BLACK);
            style.wrapText = true;
            style.noBorder = true;
            return style;
        }
    }

    private static class CellStyleConfig {
        String cellColorSuffix;
        String cellWrapTextSuffix;

        String cellHorizontalAlignmentSuffix;
        String cellVerticalAlignmentSuffix;

        String fontNameSuffix;
        String fontHeightSuffix;
        String fontBoldSuffix;
        String fontItalicSuffix;
        String fontColorSuffix;

        CellStyleConfig(Map<String, Object> styleMap) {
            if (styleMap == null) {
                return;
            }

            this.cellColorSuffix = asString(styleMap.get("cell_color_suffix"));
            this.cellWrapTextSuffix = asString(styleMap.get("cell_wrap_text_suffix"));

            this.cellHorizontalAlignmentSuffix = asString(styleMap.get("cell_horizontal_alignment_suffix"));
            this.cellVerticalAlignmentSuffix = asString(styleMap.get("cell_vertical_alignment_suffix"));

            this.fontNameSuffix = asString(styleMap.get("font_name_suffix"));
            this.fontHeightSuffix = asString(styleMap.get("font_height_suffix"));
            this.fontBoldSuffix = asString(styleMap.get("font_bold_suffix"));
            this.fontItalicSuffix = asString(styleMap.get("font_italic_suffix"));
            this.fontColorSuffix = asString(styleMap.get("font_color_suffix"));
        }

        private static String asString(Object value) {
            return value == null ? null : String.valueOf(value);
        }
    }

    private static void applyBorderStyle(PdfStyle style, Map<String, Object> borderMap) {
        if (style == null || borderMap == null || borderMap.isEmpty()) {
            return;
        }

        /*
         * If contains "all", apply same border to all sides.
         *
         * Example:
         * border_style: {
         *   "all": "thin"
         * }
         */
        Object allBorderObj = borderMap.get("all");
        if (allBorderObj != null) {
            float width = toBorderWidth(allBorderObj);

            if (width > 0f) {
                style.noBorder = false;

                style.borderTop = true;
                style.borderBottom = true;
                style.borderLeft = true;
                style.borderRight = true;

                style.borderTopWidth = width;
                style.borderBottomWidth = width;
                style.borderLeftWidth = width;
                style.borderRightWidth = width;
            }

            return;
        }

        /*
         * Otherwise apply side by side.
         *
         * Example:
         * border_style: {
         *   "left": "thin",
         *   "right": "thin",
         *   "bottom": "medium"
         * }
         */
        applySingleBorderSide(style, "top", borderMap.get("top"));
        applySingleBorderSide(style, "bottom", borderMap.get("bottom"));
        applySingleBorderSide(style, "left", borderMap.get("left"));
        applySingleBorderSide(style, "right", borderMap.get("right"));
    }

    private static void applySingleBorderSide(PdfStyle style, String side, Object borderValue) {
        if (style == null || borderValue == null) {
            return;
        }

        float width = toBorderWidth(borderValue);
        if (width <= 0f) {
            return;
        }

        style.noBorder = false;

        switch (side) {
            case "top" -> {
                style.borderTop = true;
                style.borderTopWidth = width;
            }
            case "bottom" -> {
                style.borderBottom = true;
                style.borderBottomWidth = width;
            }
            case "left" -> {
                style.borderLeft = true;
                style.borderLeftWidth = width;
            }
            case "right" -> {
                style.borderRight = true;
                style.borderRightWidth = width;
            }
            default -> {
                // Ignore unknown border side.
            }
        }
    }

    private static float toBorderWidth(Object value) {
        if (value == null) {
            return 0f;
        }
        if (value instanceof Number n) {
            return n.floatValue();
        }

        String border = value.toString().trim().toLowerCase(Locale.ROOT);
        return switch (border) {
            case "none", "no_border", "no-border", "false" -> 0f;
            case "hair" -> 0.25f;
            case "medium" -> 1.0f;
            case "thick" -> 1.5f;
            case "double" -> 1.2f;
            default -> 0.5f;
        };
    }

    private static float[] getKeyValueColumnWidths(LinkedHashMap<?, ?> data) {
        int maxKeyLength = 0;

        for (Object key : data.keySet()) {
            if (key == null) {
                continue;
            }

            maxKeyLength = Math.max(maxKeyLength, key.toString().length());
        }

        /*
         * Estimate key width based on text length.
         * Minimum 1.0f keeps short keys readable.
         * Maximum 4.0f prevents long keys from taking too much space.
         */
        float keyWidth = Math.max(1.0f, maxKeyLength / 8.0f);
        keyWidth = Math.min(keyWidth, 4.0f);

        /*
         * Value gets the remaining larger portion.
         * Total does not matter; only ratio matters.
         */
        float valueWidth = 12.0f - keyWidth;

        if (valueWidth < 4.0f) {
            valueWidth = 4.0f;
        }

        return new float[]{keyWidth, valueWidth};
    }

    private static int toHorizontalAlignment(Object value) {
        if (value == null) {
            return Element.ALIGN_LEFT;
        }
        if (value instanceof Number n) {
            return n.intValue();
        }

        String align = value.toString().trim().toLowerCase(Locale.ROOT);
        return switch (align) {
            case "center", "centre" -> Element.ALIGN_CENTER;
            case "right" -> Element.ALIGN_RIGHT;
            case "justify" -> Element.ALIGN_JUSTIFIED;
            default -> Element.ALIGN_LEFT;
        };
    }

    private static int toVerticalAlignment(Object value) {
        if (value == null) {
            return Element.ALIGN_MIDDLE;
        }
        if (value instanceof Number n) {
            return n.intValue();
        }

        String align = value.toString().trim().toLowerCase(Locale.ROOT);
        return switch (align) {
            case "top" -> Element.ALIGN_TOP;
            case "bottom" -> Element.ALIGN_BOTTOM;
            default -> Element.ALIGN_MIDDLE;
        };
    }

    private static boolean isNumericValue(Object value) {
        if (value instanceof Number) {return true;}
        if (!(value instanceof String str)) {return false;}
        if (str.isBlank()) {return false;}

        try {
            new BigDecimal(str.trim().replace(",", ""));
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static PdfStyle createBodyHeaderStyle(Map<String, Object> styles, String defaultBodyHeaderStyle, float tableScaleRatio) {
        /*
         * If default_body_header_style is provided,
         * use it as a style name and resolve from pdf_style.
         */
        if (defaultBodyHeaderStyle != null && !defaultBodyHeaderStyle.isBlank()) {
            return getPdfStyle(styles, defaultBodyHeaderStyle, tableScaleRatio);
        }

        /*
         * Fallback default body header style.
         */
        PdfStyle style = PdfStyle.defaultBody(tableScaleRatio);

        style.font = new Font(
                Font.HELVETICA,
                Math.max(5.5f, 13f * tableScaleRatio),
                Font.BOLD,
                Color.BLACK
        );
        style.backgroundColor = Color.YELLOW;
        style.horizontalAlignment = Element.ALIGN_LEFT;
        style.verticalAlignment = Element.ALIGN_MIDDLE;
        style.wrapText = true;
        style.noBorder = true;

        return style;
    }

    private static PdfStyle extractPdfStyle(PdfStyle source) {
        if (source == null) {
            return PdfStyle.defaultBody(1.0f);
        }

        PdfStyle resultStyle = new PdfStyle();

        resultStyle.font = source.font;
        resultStyle.backgroundColor = source.backgroundColor;
        resultStyle.horizontalAlignment = source.horizontalAlignment;
        resultStyle.verticalAlignment = source.verticalAlignment;
        resultStyle.wrapText = source.wrapText;
        resultStyle.noBorder = source.noBorder;

        resultStyle.borderColor = source.borderColor;
        resultStyle.borderWidth = source.borderWidth;

        resultStyle.borderTop = source.borderTop;
        resultStyle.borderBottom = source.borderBottom;
        resultStyle.borderLeft = source.borderLeft;
        resultStyle.borderRight = source.borderRight;

        resultStyle.borderTopWidth = source.borderTopWidth;
        resultStyle.borderBottomWidth = source.borderBottomWidth;
        resultStyle.borderLeftWidth = source.borderLeftWidth;
        resultStyle.borderRightWidth = source.borderRightWidth;

        resultStyle.scaleRate = source.scaleRate;

        return resultStyle;
    }
}