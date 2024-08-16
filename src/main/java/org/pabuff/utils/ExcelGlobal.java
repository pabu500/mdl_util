package org.pabuff.utils;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@Getter
@AllArgsConstructor
public class ExcelGlobal {
    private final String cellColorSuffix;
    private final String cellFillPatternSuffix;

    private final String fontColorSuffix;
    private final String fontNameSuffix;
    private final String fontSizeSuffix;
    private final String fontHeightInPointsSuffix;
    private final String fontBoldSuffix;

    public ExcelGlobal() {
        this.cellColorSuffix = "_cell_clr";
        this.cellFillPatternSuffix = "_cell_fill_pattern";
        this.fontColorSuffix = "_font_clr";
        this.fontNameSuffix = "_font_name";
        this.fontSizeSuffix = "_font_size";
        this.fontHeightInPointsSuffix = "_font_height_in_points";
        this.fontBoldSuffix = "_font_bold";
    }
}
