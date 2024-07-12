package org.pabuff.utils;

import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
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

    public final static String desTimeStampFormat = "yyyy-MM-dd HH:mm:ss";
    public final static String desTimeStampFormatIncludingMs = "yyyy-MM-dd HH:mm:ss.SSS";
    public final static String desTimeStampFormatIncludingMs6 = "yyyy-MM-dd HH:mm:ss.SSSSSS";
    public final static String desTimeStampFormatIncludingMs6WithZone = "yyyy-MM-dd HH:mm:ss.SSSSSSXXX";

    public final static SimpleDateFormat sdf = new SimpleDateFormat(desTimeStampFormat);
    public final static DateTimeFormatter formatter = DateTimeFormatter.ofPattern(desTimeStampFormat);
    public final static SimpleDateFormat sdfms = new SimpleDateFormat(desTimeStampFormatIncludingMs);
    public final static DateTimeFormatter formatterMs = DateTimeFormatter.ofPattern(desTimeStampFormatIncludingMs);
    public final static SimpleDateFormat sdfms6 = new SimpleDateFormat(desTimeStampFormatIncludingMs6);
    public final static DateTimeFormatter formatterMs6 = DateTimeFormatter.ofPattern(desTimeStampFormatIncludingMs6);
    public static DateTimeFormatter formatterIso8601 = DateTimeFormatter.ISO_DATE_TIME;
    public static DateTimeFormatter formatterMs6WithZone = DateTimeFormatter.ofPattern(desTimeStampFormatIncludingMs6WithZone);

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
                    try {
                        return LocalDateTime.parse(dateTimeStr, formatterIso8601);
                    } catch (Exception e4) {
                        try {
                            return OffsetDateTime.parse(dateTimeStr, formatterMs6WithZone).toLocalDateTime();
                        } catch (Exception e5) {
                            return null;
                        }
                    }
                }
            }
        }
    }
    public static LocalDateTime getLocalDateTime2(String dateTimeStr, boolean tryZoneFormatterFirst) {
        if(tryZoneFormatterFirst){
            try {
                return OffsetDateTime.parse(dateTimeStr, formatterMs6WithZone).toLocalDateTime();
            }catch (Exception e) {
                try {
                    return LocalDateTime.parse(dateTimeStr, formatterMs6);
                } catch (Exception e1) {
                    try {
                        return LocalDateTime.parse(dateTimeStr, formatterMs);
                    } catch (Exception e2) {
                        try {
                            return LocalDateTime.parse(dateTimeStr, formatter);
                        } catch (Exception e3) {
                            try {
                                return LocalDateTime.parse(dateTimeStr, formatterIso8601);
                            } catch (Exception e4) {
                                return null;
                            }
                        }
                    }
                }
            }
        }else {
            return getLocalDateTime(dateTimeStr);
        }
    }
    public static ZonedDateTime getZonedDateTime(LocalDateTime dateTime, ZoneId zoneId) {
        ZonedDateTime zonedLocalDateTime = dateTime.atZone(ZoneId.systemDefault());
        return zonedLocalDateTime.withZoneSameInstant(zoneId);
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
    public static String getZonedDateTimeStr(LocalDateTime localDateTime, ZoneId zoneId) {
        try {
            return Objects.requireNonNull(getZonedDateTime(localDateTime, zoneId)).format(formatter);
        }catch (Exception e) {
            try {
                return Objects.requireNonNull(getZonedDateTime(localDateTime, zoneId)).format(formatterMs);
            } catch (Exception e1) {
                try {
                    return Objects.requireNonNull(getZonedDateTime(localDateTime, zoneId)).format(formatterMs6);
                } catch (Exception e2) {
                    return null;
                }
            }
        }
    }

    public static String getLocalDateTimeStr(LocalDateTime dateTime) {
        return dateTime.format(formatter);
    }

    public static String getLocalDateTimeStr(LocalDateTime dateTime, String format) {
        return dateTime.format(DateTimeFormatter.ofPattern(format));
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

    public static LocalDateTime getSgNow () {
//        return getZonedLocalDateTimeFromSystemLocalDateTime(LocalDateTime.now(), ZoneId.of("Asia/Singapore"));
        return LocalDateTime.now(ZoneId.of("Asia/Singapore"));
    }
    public static String getSgNowStr() {
        return getZonedDateTimeStr(LocalDateTime.now(), ZoneId.of("Asia/Singapore"));
    }

    public static String getSgNowStrMs() {
        return getZonedDateTime(LocalDateTime.now(), ZoneId.of("Asia/Singapore")).format(formatterMs);
    }

    public static int getDaysInMonth(LocalDateTime dateTime) {
        int year = dateTime.getYear();
        boolean isLeapYear = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
        return dateTime.getMonth().length(isLeapYear);
    }

    public static LocalDateTime alignToDayStart(LocalDateTime dateTime) {
        return dateTime.withHour(0).withMinute(0).withSecond(0).withNano(0);
    }
    public static LocalDateTime alignStrToDayStart(String dateTimeStr) {
        LocalDateTime dateTime = getLocalDateTime(dateTimeStr);
        if(dateTime == null) {
            return null;
        }
        return alignToDayStart(dateTime);
    }
    public static String alignStrToDayStartStr(String dateTimeStr) {
        LocalDateTime dateTime = getLocalDateTime(dateTimeStr);
        if(dateTime == null) {
            return null;
        }
        return alignToDayStart(dateTime).format(formatter);
    }
    public static LocalDateTime alignToDayEnd(LocalDateTime dateTime) {
        return dateTime.withHour(23).withMinute(59).withSecond(59).withNano(999999999);
    }
    public static String alignToDayEndStr(LocalDateTime dateTime) {
        return alignToDayEnd(dateTime).format(formatter);
    }
    public static LocalDateTime alignStrToDayEnd(String dateTimeStr) {
        LocalDateTime dateTime = getLocalDateTime(dateTimeStr);
        if(dateTime == null) {
            return null;
        }
        return alignToDayEnd(dateTime);
    }
    public static String alignStrToDayEndStr(String dateTimeStr) {
        LocalDateTime dateTime = getLocalDateTime(dateTimeStr);
        if(dateTime == null) {
            return null;
        }
        return alignToDayEnd(dateTime).format(formatter);
    }
}
