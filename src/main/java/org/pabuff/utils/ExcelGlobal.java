package org.pabuff.utils;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;

@Getter
@AllArgsConstructor
@Component
public class ExcelGlobal {
    private final String cellColorSuffix;
    private final String cellFillPatternSuffix;

    private final String fontColorSuffix;
    private final String fontNameSuffix;
    private final String fontHeightInPointsSuffix;
    private final String fontBoldSuffix;
    private final String fontItalicSuffix;

    public ExcelGlobal(Map<String, Object> map) {
        this.cellColorSuffix = map.get("cell_color") != null ? (String) map.get("cell_color") : null;
        this.cellFillPatternSuffix = map.get("cell_fill_pattern") != null ? (String) map.get("cell_fill_pattern") : null;

        this.fontColorSuffix = map.get("font_color") != null ? (String) map.get("font_color") : null;
        this.fontNameSuffix =  map.get("font_name") != null ? (String) map.get("font_name") : null;
        this.fontHeightInPointsSuffix = map.get("font_height") != null ? (String) map.get("font_height") : null;
        this.fontBoldSuffix = map.get("font_bold") != null ? (String) map.get("font_bold") : null;
        this.fontItalicSuffix = map.get("font_italic") != null ? (String) map.get("font_italic") : null;
    }
}

