package com.xt.utils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    public static Long findDominantLong(List<Double> numbers) {
        Map<Long, Integer> frequencyMap = new HashMap<>();
        long dominantLong = Math.round(numbers.get(0));
        int dominantFrequency = 1;

        for (double number : numbers) {
            Long candidate = Math.round(number);
            int count = frequencyMap.getOrDefault(candidate, 0) + 1;
            frequencyMap.put(candidate, count);

            if (count > dominantFrequency) {
                dominantLong = candidate;
                dominantFrequency = count;
            }
        }
        return dominantLong;
    }
}
