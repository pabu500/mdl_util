package com.xt.utils;

import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class DateTimeUtil {
    final static long fourHours = 14400000L;
    final static long twoHours = 7200000L;
    final static long oneHour = 3600000L;
    final static long halfHour = 1800000L;
    final static long fifteenMin = 900000L;
    final static long fiveMin = 300000L;

    private final static String desTimeStampFormat = "yyyy-MM-dd HH:mm:ss";
    private final static String desTimeStampFormatIncludingMs = "yyyy-MM-dd HH:mm:ss.SSS";

    public final static SimpleDateFormat sdf = new SimpleDateFormat(desTimeStampFormat);
    public final static DateTimeFormatter formatter = DateTimeFormatter.ofPattern(desTimeStampFormat);
    public final static SimpleDateFormat sdfms = new SimpleDateFormat(desTimeStampFormatIncludingMs);
    public final static DateTimeFormatter formatterMs = DateTimeFormatter.ofPattern(desTimeStampFormatIncludingMs);
    public static LocalDateTime getLocalDateTime(String dateTimeStr) {
        return LocalDateTime.parse(dateTimeStr, formatter);
    }

    public static String getLocalDateTimeStr(LocalDateTime dateTime) {
        return dateTime.format(formatter);
    }

}
