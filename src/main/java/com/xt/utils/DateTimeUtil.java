package com.xt.utils;

import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class DateTimeUtil {
    public final static long oneYear = 31536000000L;
    public final static long halfYear = 15768000000L;
    public final static long oneMonth = 2592000000L;
    public final static long halfMonth = 1296000000L;
    public final static long oneWeek = 604800000L;
    public final static long halfWeek = 302400000L;
    public final static long threeDays = 259200000L;
    public final static long twoDays = 172800000L;
    public final static long oneDay = 86400000L;
    public final static long halfDay = 43200000L;
    public final static long fourHours = 14400000L;
    public final static long twoHours = 7200000L;
    public final static long oneHour = 3600000L;
    public final static long halfHour = 1800000L;
    public final static long fifteenMin = 900000L;
    public final static long fiveMin = 300000L;
    public final static long oneMin = 60000L;
    public final static long halfMin = 30000L;
    public final static long tenSec = 10000L;

    public  final static String desTimeStampFormat = "yyyy-MM-dd HH:mm:ss";
    public  final static String desTimeStampFormatIncludingMs = "yyyy-MM-dd HH:mm:ss.SSS";

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

    public static Long getTimeStampFromLocalDateTime(LocalDateTime dateTime) {
        return dateTime.atZone(java.time.ZoneId.systemDefault()).toInstant().toEpochMilli();
    }
    public static Long getTimeStampFromLocalDateTimeStr(String dateTimeStr) {
        return getLocalDateTime(dateTimeStr).atZone(java.time.ZoneId.systemDefault()).toInstant().toEpochMilli();
    }
}
