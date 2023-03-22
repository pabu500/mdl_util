package com.xt.utils;

public class MathUtil {
    static Double ObjToDouble(Object obj) {
        if (obj == null) {
            return null;
        }
        if (obj instanceof String) {
            return Double.parseDouble(obj.toString());
        }
        return ((Number) obj).doubleValue();
    }
}
