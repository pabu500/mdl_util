package com.xt.utils;

import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Objects;

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
    public final static long fifteenMinutes = 900000L;
    public final static long fiveMinutes = 300000L;
    public final static long oneMinute = 60000L;
    public final static long halfMinute = 30000L;
    public final static long tenSec = 10000L;

    public  final static String desTimeStampFormat = "yyyy-MM-dd HH:mm:ss";
    public  final static String desTimeStampFormatIncludingMs = "yyyy-MM-dd HH:mm:ss.SSS";
    public  final static String desTimeStampFormatIncludingMs6 = "yyyy-MM-dd HH:mm:ss.SSSSSS";

    public final static SimpleDateFormat sdf = new SimpleDateFormat(desTimeStampFormat);
    public final static DateTimeFormatter formatter = DateTimeFormatter.ofPattern(desTimeStampFormat);
    public final static SimpleDateFormat sdfms = new SimpleDateFormat(desTimeStampFormatIncludingMs);
    public final static DateTimeFormatter formatterMs = DateTimeFormatter.ofPattern(desTimeStampFormatIncludingMs);
    public final static SimpleDateFormat sdfms6 = new SimpleDateFormat(desTimeStampFormatIncludingMs6);
    public final static DateTimeFormatter formatterMs6 = DateTimeFormatter.ofPattern(desTimeStampFormatIncludingMs6);
    public static LocalDateTime getLocalDateTime(String dateTimeStr) {
        try {
            return LocalDateTime.parse(dateTimeStr, formatter);
        }catch (Exception e) {
            try {
                return LocalDateTime.parse(dateTimeStr, formatterMs);
            } catch (Exception e1) {
                try {
                    return LocalDateTime.parse(dateTimeStr, formatterMs6);
                } catch (Exception e2) {
                    return null;
                }
            }
        }
    }

    public static ZonedDateTime getZonedDateTime(LocalDateTime dateTime, ZoneId zoneId) {
        return dateTime.atZone(zoneId);
    }

    public static String getZonedDateTimeStr(String localDateTimeStr, ZoneId zoneId) {
        try {
            return Objects.requireNonNull(getZonedDateTime(localDateTimeStr, zoneId)).format(formatter);
        }catch (Exception e) {
            try {
                return Objects.requireNonNull(getZonedDateTime(localDateTimeStr, zoneId)).format(formatterMs);
            } catch (Exception e1) {
                try {
                    return Objects.requireNonNull(getZonedDateTime(localDateTimeStr, zoneId)).format(formatterMs6);
                } catch (Exception e2) {
                    return null;
                }
            }
        }
    }

    public static ZonedDateTime getZonedDateTime(String dateTimeStr, ZoneId zoneId) {
        try {
            return ZonedDateTime.parse(dateTimeStr, formatter).withZoneSameInstant(zoneId);
        }catch (Exception e) {
            try {
                return ZonedDateTime.parse(dateTimeStr, formatterMs).withZoneSameInstant(zoneId);
            } catch (Exception e1) {
                try {
                    return ZonedDateTime.parse(dateTimeStr, formatterMs6).withZoneSameInstant(zoneId);
                } catch (Exception e2) {
                    return null;
                }
            }
        }
    }

    public static String getLocalDateTimeStr(LocalDateTime dateTime) {
        return dateTime.format(formatter);
    }

    public static String getZonedDateTimeStr(LocalDateTime dateTime, String format, ZoneId zoneId) {
        return dateTime.atZone(zoneId).format(DateTimeFormatter.ofPattern(format));
    }

    public static Long getTimeStampFromLocalDateTime(LocalDateTime dateTime) {
        return dateTime.atZone(java.time.ZoneId.systemDefault()).toInstant().toEpochMilli();
    }
    public static Long getTimeStampFromLocalDateTimeStr(String dateTimeStr) {
        return getLocalDateTime(dateTimeStr).atZone(java.time.ZoneId.systemDefault()).toInstant().toEpochMilli();
    }

    public static LocalDateTime getZonedLocalDateTimeFromSystemLocalDateTime(LocalDateTime systemLocalDateTime, ZoneId zoneId) {
        return systemLocalDateTime.atZone(java.time.ZoneId.systemDefault()).withZoneSameInstant(zoneId).toLocalDateTime();
    }
}
