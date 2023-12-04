package com.xt.utils;

import lombok.*;

@Setter
@Getter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class XtStat {
    double min;
    double minNonZero;
    double max;
    double total;
    double avg;
    double median;
    long negativeCount;
    long positiveCount;
    long totalCount;
}
