package org.pabuff.utils;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@Getter
@AllArgsConstructor
public class ExcelGlobal {
    private final String cellColorSuffix;
    private final String fontColorSuffix;
    private final String fontNameSuffix;
    private final String fontSizeSuffix;
    private final String cellFillPatternSuffix;

    public ExcelGlobal() {
        this.cellColorSuffix = "_cell_clr";
        this.fontColorSuffix = "_font_clr";
        this.fontNameSuffix = "_font_name";
        this.fontSizeSuffix = "_font_size";
        this.cellFillPatternSuffix = "_cell_fill_pattern";
    }
}
