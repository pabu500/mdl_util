package com.xt.utils;

public class MathUtil {
    public static Double ObjToDouble(Object obj) {
        if (obj == null) {
            return null;
        }
        if (obj instanceof String) {
            return Double.parseDouble(obj.toString());
        }
        return ((Number) obj).doubleValue();
    }
    public static Long ObjToLong(Object obj) {
        if (obj == null) {
            return null;
        }
        if (obj instanceof String) {
            return Long.parseLong(obj.toString());
        }
        return ((Number) obj).longValue();
    }
    public static Float ObjToFloat(Object obj) {
        if (obj == null) {
            return null;
        }
        if (obj instanceof String) {
            return Float.parseFloat(obj.toString());
        }
        return ((Number) obj).floatValue();
    }
    public static Integer ObjToInteger(Object obj) {
        if (obj == null) {
            return null;
        }
        if (obj instanceof String) {
            return Integer.parseInt(obj.toString());
        }
        return ((Number) obj).intValue();
    }
}
