package com.xt.utils;

import java.math.RoundingMode;
import java.text.DecimalFormat;
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

    public static Double setDecimalPlaces(Double value, int decimalPlaces, RoundingMode roundingMode) {
        if (value == null) {
            return null;
        }
        DecimalFormat df = new DecimalFormat("#." + "#".repeat(Math.max(0, decimalPlaces)));
        df.setRoundingMode(roundingMode);
        return Double.parseDouble(df.format(value));
    }

    public static double findMin(List<Double> numbers) {
        double min = numbers.get(0);
        for (double number : numbers) {
            if (number < min) {
                min = number;
            }
        }
        return min;
    }
    public static double findMinNonZero(List<Double> numbers) {
        double min = numbers.get(0);
        for (double number : numbers) {
            if (number < min && number != 0) {
                min = number;
            }
        }
        return min;
    }

    public static double findMax(List<Double> numbers) {
        double max = numbers.get(0);
        for (double number : numbers) {
            if (number > max) {
                max = number;
            }
        }
        return max;
    }
    public static double findTotal(List<Double> numbers) {
        double sum = 0;
        for (double number : numbers) {
            sum += number;
        }
        return sum;
    }
    public static long findPositiveCount(List<Double> numbers) {
        long count = 0;
        for (double number : numbers) {
            if (number > 0) {
                count++;
            }
        }
        return count;
    }
    public static double findAverage(List<Double> numbers) {
        double sum = 0;
        for (double number : numbers) {
            sum += number;
        }
        return sum / numbers.size();
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

    public static double findStandardDiv(List<Double> numbers) {
        double mean = findAverage(numbers);
        double sum = 0;
        for (double number : numbers) {
            sum += Math.pow(number - mean, 2);
        }
        return Math.sqrt(sum / numbers.size());
    }

    public static String genRandomNumStr(int len) {
        StringBuilder str = new StringBuilder();
        for (int i = 0; i < len; i++) {
            str.append((int) (Math.random() * 10));
        }
        return str.toString();
    }
}
