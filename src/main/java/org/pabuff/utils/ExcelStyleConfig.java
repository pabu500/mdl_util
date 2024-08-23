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

    private final String fontColorSuffix;
    private final String fontNameSuffix;
    private final String fontHeightInPointsSuffix;
    private final String fontBoldSuffix;
    private final String fontItalicSuffix;

    public ExcelStyleConfig(Map<String, Object> map) {
        this.cellColorSuffix = map.get("cell_color") != null ? (String) map.get("cell_color") : null;
        this.cellFillPatternSuffix = map.get("cell_fill_pattern") != null ? (String) map.get("cell_fill_pattern") : null;
        this.cellWrapTextSuffix = map.get("cell_wrap_text") != null ? (String) map.get("cell_wrap_text") : null;

        this.fontColorSuffix = map.get("font_color") != null ? (String) map.get("font_color") : null;
        this.fontNameSuffix =  map.get("font_name") != null ? (String) map.get("font_name") : null;
        this.fontHeightInPointsSuffix = map.get("font_height") != null ? (String) map.get("font_height") : null;
        this.fontBoldSuffix = map.get("font_bold") != null ? (String) map.get("font_bold") : null;
        this.fontItalicSuffix = map.get("font_italic") != null ? (String) map.get("font_italic") : null;
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

