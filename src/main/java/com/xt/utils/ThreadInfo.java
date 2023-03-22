package com.xt.utils;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor
public class ThreadInfo {
    private long id;
    private String name;
    private String status;
    private int priority;
    private boolean isDaemon;

}
