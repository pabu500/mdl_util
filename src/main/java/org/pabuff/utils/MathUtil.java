package org.pabuff.utils;

import java.math.BigDecimal;
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
    public static long findNegativeCount(List<Double> numbers) {
        long count = 0;
        for (double number : numbers) {
            if (number < 0) {
                count++;
            }
        }
        return count;
    }
    public static XtStat findStat(List<Double> numbers) {
        double min = Double.MAX_VALUE;
        double minNonZero = Double.MAX_VALUE;
        double max = Double.MIN_VALUE;
        double total = 0;
        double avg = 0;
        double median = 0;
        long negativeCount = 0;
        long positiveCount = 0;
        long totalCount = numbers.size();

        for (double number : numbers) {
            if (number < min) {
                min = number;
            }
            if (number < minNonZero && number != 0) {
                minNonZero = number;
            }
            if (number > max) {
                max = number;
            }
            total += number;
            if (number < 0) {
                negativeCount++;
            } else {
                positiveCount++;
            }
        }
        avg = total / totalCount;
        median = findMedian(numbers);

        return XtStat.builder()
                .min(min)
                .minNonZero(minNonZero)
                .max(max)
                .total(total)
                .avg(avg)
                .median(median)
                .negativeCount(negativeCount)
                .positiveCount(positiveCount)
                .totalCount(totalCount)
                .build();
    }
    public static double findAverage(List<Double> numbers) {
        double sum = 0;
        for (double number : numbers) {
            sum += number;
        }
        return sum / numbers.size();
    }
    public static double findMedian(List<Double> numbers) {
        numbers.sort(Double::compareTo);
        int size = numbers.size();
        if (size % 2 == 0) {
            return (numbers.get(size / 2 - 1) + numbers.get(size / 2)) / 2;
        } else {
            return numbers.get(size / 2);
        }
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

    public static Map<String, Double> findIntervalStat(List<Double> numbers, double threshold) {
        Map<Long, Integer> frequencyMap = new HashMap<>();
        long dominantLong = Math.round(numbers.get(0));
        int dominantFrequency = 1;
        long intervalOutlierCount = 0;
        double minNonZeroInterval = Double.MAX_VALUE;
        double maxInterval = Double.MIN_VALUE;

        for (double number : numbers) {
            Long candidate = Math.round(number);
            int count = frequencyMap.getOrDefault(candidate, 0) + 1;
            frequencyMap.put(candidate, count);

            if (count > dominantFrequency) {
                dominantLong = candidate;
                dominantFrequency = count;
            }
        }
        for (double number : numbers) {
            if (number > threshold * dominantLong) {
                intervalOutlierCount++;
            }
            //minNonZeroInterval
            if (number < minNonZeroInterval && number != 0) {
                minNonZeroInterval = number;
            }
            //maxInterval
            if (number > maxInterval) {
                maxInterval = number;
            }
        }
        Map<String, Double> intervalStat = new HashMap<>();
        intervalStat.put("dominant_interval", (double) dominantLong);
        intervalStat.put("outlier_count", (double) intervalOutlierCount);
        intervalStat.put("min_non_zero_interval", minNonZeroInterval);
        intervalStat.put("max_interval", maxInterval);

        return intervalStat;
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
    public static double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();

        BigDecimal bd = BigDecimal.valueOf(value);
        bd = bd.setScale(places, RoundingMode.HALF_UP);
        return bd.doubleValue();
    }
}

