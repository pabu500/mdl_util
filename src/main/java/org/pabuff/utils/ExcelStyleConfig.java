package org.pabuff.utils;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.lang.reflect.Field;
import java.util.Map;

@Getter
@AllArgsConstructor
public class ExcelStyleConfig {
    private final String cellColorSuffix;
    private final String cellFillPatternSuffix;
    private final String cellWrapTextSuffix;

    private final String cellHorizontalAlignmentSuffix;
    private final String cellVerticalAlignmentSuffix;

    private final String fontColorSuffix;
    private final String fontNameSuffix;
    private final String fontHeightInPointsSuffix;
    private final String fontBoldSuffix;
    private final String fontItalicSuffix;

    public ExcelStyleConfig(Map<String, Object> map) {
        this.cellColorSuffix = getString(map, "cell_color");
        this.cellFillPatternSuffix = getString(map, "cell_fill_pattern");
        this.cellWrapTextSuffix = getString(map, "cell_wrap_text");

        this.cellHorizontalAlignmentSuffix = getString(map, "cell_horizontal_alignment");
        this.cellVerticalAlignmentSuffix = getString(map, "cell_vertical_alignment");

        this.fontColorSuffix = getString(map, "font_color");
        this.fontNameSuffix = getString(map, "font_name");
        this.fontHeightInPointsSuffix = getString(map, "font_height");
        this.fontBoldSuffix = getString(map, "font_bold");
        this.fontItalicSuffix = getString(map, "font_italic");
    }

    private static String getString(Map<String, Object> map, String key) {
        Object value = map.get(key);
        return value != null ? (String) value : null;
    }

    //check str contains any of the suffix
    public boolean containsAnySuffix(String str) {
        Field[] fields = this.getClass().getDeclaredFields();
        for (Field field : fields) {
            if (field.getType().equals(String.class)) {
                field.setAccessible(true); // Make private fields accessible
                try {
                    String value = (String) field.get(this);
                    if (value != null && str.contains(value)) {
                        return true;
                    }
                } catch (IllegalAccessException e) {
                    throw new RuntimeException("Failed to access field value", e);
                }
            }
        }
        return false;
    }
}

