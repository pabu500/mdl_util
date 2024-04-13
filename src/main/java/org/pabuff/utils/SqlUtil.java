package org.pabuff.utils;

import java.util.Map;

public class SqlUtil {
    public static Map<String, String> makeSelectSql(Map<String, Object> sqlMap) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ");
        if(sqlMap.get("select") != null) {
            sql.append(sqlMap.get("select"));
        } else {
            sql.append("*");
        }

        sql.append(" FROM ");
        if(sqlMap.get("from") != null) {
            sql.append(sqlMap.get("from"));
        } else {
            return Map.of("error", "Missing table name");
        }

        StringBuilder targetConstraint = new StringBuilder();
        if(sqlMap.get("target_key") != null) {
            targetConstraint = new StringBuilder(sqlMap.get("target_key") + " = '" + sqlMap.get("target_value") + "'");
        }else{
            //multiple target
            if(sqlMap.get("targets") != null){
                if(sqlMap.get("targets") instanceof Map<?,?>){
                    Map<String, String> targets = (Map<String, String>) sqlMap.get("targets");
                    if(!targets.keySet().isEmpty()) {
                        for (String key : targets.keySet()) {
                            targetConstraint.append(key).append(" = '").append(targets.get(key)).append("' AND ");
                        }
                        targetConstraint = new StringBuilder(targetConstraint.substring(0, targetConstraint.length() - 5));
                    }
                }
            }
        }

        if(sqlMap.get("is_not_null") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("is_not_null") + " IS NOT NULL");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("is_not_null")).append(" IS NOT NULL");
            }
        }

        if(sqlMap.get("is_not_empty") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("is_not_empty") + " != ''");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("is_not_empty")).append(" != ''");
            }
        }

        if(sqlMap.get("additional_constraint") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(String.valueOf(sqlMap.get("additional_constraint")));
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("additional_constraint"));
            }
        }

        String timeConstraint = "";
        if(sqlMap.get("time_key") != null) {
            if(sqlMap.get("start_datetime")!=null){
                timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime");
            }
            if(sqlMap.get("end_datetime")!=null){
                if(timeConstraint.isEmpty()){
                    timeConstraint = sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }else{
                    timeConstraint += "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }
            }
//            timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime") + "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime") + "'";
        }

        if(!targetConstraint.toString().isEmpty() && !timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint).append(" AND ").append(timeConstraint);
        } else if(!targetConstraint.toString().isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint);
        } else if(!timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(timeConstraint);
        }

        if(sqlMap.get("time_key")!=null){
            sql.append(" ORDER BY ").append(sqlMap.get("time_key")).append(" DESC");
        }

        String limit = sqlMap.get("limit").toString();
        if(limit != null) {
            sql.append(" LIMIT ").append(limit);
        }

        return Map.of("sql", sql.toString());
    }

    public static Map<String, String> makeSelectLikeSql(Map<String, Object> sqlMap) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ");
        if(sqlMap.get("select") != null) {
            sql.append(sqlMap.get("select"));
        } else {
            sql.append("*");
        }

        sql.append(" FROM ");
        if(sqlMap.get("from") != null) {
            sql.append(sqlMap.get("from"));
        } else {
            return Map.of("error", "Missing table name");
        }

        StringBuilder targetConstraint = new StringBuilder();
        boolean includeNullValue = sqlMap.get("include_null_value") != null && sqlMap.get("include_null_value").equals("true");

        if(sqlMap.get("target_key") != null) {
            targetConstraint = new StringBuilder(sqlMap.get("target_key") + " like '%" + sqlMap.get("target_value") + "%'");
        }else{
            //multiple target
            if(sqlMap.get("targets") != null){
                if(sqlMap.get("targets") instanceof Map<?,?>){
                    Map<String, Object> targets = (Map<String, Object>) sqlMap.get("targets");
                    if(!targets.keySet().isEmpty()) {
                        for (String key : targets.keySet()) {
                            Object value = targets.get(key);
                            if(value == null ){
                                if(includeNullValue){
                                    targetConstraint.append(key).append(" IS NULL AND ");
                                }else {
                                    continue;
                                }
                            }else {
                                if (value instanceof Integer || value instanceof Double) {
                                    targetConstraint.append(key).append(" = ").append(targets.get(key)).append(" AND ");
                                    continue;
                                }
                                if(value instanceof String){
                                    if(((String) value).isEmpty()) {
                                        targetConstraint.append(key).append(" = '' AND ");
                                        continue;
                                    }
                                }
                                targetConstraint.append(key).append(" like '%").append(targets.get(key)).append("%' AND ");
                            }
                        }
                        targetConstraint = new StringBuilder(targetConstraint.substring(0, targetConstraint.length() - 5));
                    }
                }
            }
        }

        if(sqlMap.get("is_not_null") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("is_not_null") + " IS NOT NULL");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("is_not_null")).append(" IS NOT NULL");
            }
        }

        if(sqlMap.get("is_not_empty") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("is_not_empty") + " != ''");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("is_not_empty")).append(" != ''");
            }
        }

        String timeConstraint = "";
        if(sqlMap.get("time_key") != null) {
            if(sqlMap.get("start_datetime")!=null){
                timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime");
            }
            if(sqlMap.get("end_datetime")!=null){
                if(timeConstraint.isEmpty()){
                    timeConstraint = sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }else{
                    timeConstraint += "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }
            }
//            timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime") + "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime") + "'";
        }

        if(!targetConstraint.toString().isEmpty() && !timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint).append(" AND ").append(timeConstraint);
        } else if(!targetConstraint.toString().isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint);
        } else if(!timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(timeConstraint);
        }

        if(sqlMap.get("time_key")!=null){
            sql.append(" ORDER BY ").append(sqlMap.get("time_key")).append(" DESC");
        }

        String limit = sqlMap.get("limit").toString();
        if(limit != null) {
            sql.append(" LIMIT ").append(limit);
        }

        return Map.of("sql", sql.toString());
    }

    //mix select '=' and select 'like'
    public static Map<String, String> makeSelectSql2(Map<String, Object> sqlMap) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ");
        if(sqlMap.get("select") != null) {
            sql.append(sqlMap.get("select"));
        } else {
            sql.append("*");
        }

        sql.append(" FROM ");
        if(sqlMap.get("from") != null) {
            sql.append(sqlMap.get("from"));
        } else {
            return Map.of("error", "Missing table name");
        }

        StringBuilder targetConstraint = new StringBuilder();
//        boolean isConstraintPresent = false;
        boolean includeNullValue = sqlMap.get("include_null_value") != null && sqlMap.get("include_null_value").equals("true");

        //by default, use key = value
        if(sqlMap.get("target_key") != null && sqlMap.get("target_value") != null) {
            targetConstraint = new StringBuilder(sqlMap.get("target_key") + " = '" + sqlMap.get("target_value") + "'");
//            isConstraintPresent = true;
        }else{
            //multiple target
            if(sqlMap.get("targets") != null){
                if(sqlMap.get("targets") instanceof Map<?,?>){
                    Map<String, Object> targets = (Map<String, Object>) sqlMap.get("targets");
                    if(!targets.keySet().isEmpty()) {
                        for (String key : targets.keySet()) {
//                            targetConstraint.append(key).append(" = '").append(targets.get(key)).append("' AND ");
                            Object value = targets.get(key);
                            if(value == null ) {
                                if (includeNullValue) {
//                                    targetConstraint.append(key).append(" IS NULL AND ");
//                                    if(isConstraintPresent){
//                                        targetConstraint.append(" AND ");
//                                    }
//                                    targetConstraint.append(" ( ").append(key).append(" IS NULL or ").append(key).append(" = '' ) AND ");
                                    targetConstraint.append(" AND ").append(" ( ").append(key).append(" IS NULL or ").append(key).append(" = '' ) ");
//                                    isConstraintPresent = true;
                                } else {
                                    continue;
                                }
                            }else {
//                                if(isConstraintPresent){
//                                    targetConstraint.append(" AND ");
//                                }
                                if (value instanceof Integer || value instanceof Double) {
                                    targetConstraint.append(" AND ").append(key).append(" = ").append(targets.get(key))/*.append(" AND ")*/;
                                    continue;
                                }
                                if(value instanceof String){
                                    if(((String) value).isEmpty()) {
//                                        targetConstraint.append(key).append(" = '' AND ");
                                        targetConstraint.append(" AND ").append(key).append(" = '' ");
                                        continue;
                                    }
                                }
//                                targetConstraint.append(key).append(" = '").append(targets.get(key)).append("' AND ");
                                targetConstraint.append(" AND ").append(key).append(" = '").append(targets.get(key)).append("' ");
//                                isConstraintPresent = true;
                            }
                        }
//                        targetConstraint = new StringBuilder(targetConstraint.substring(0, targetConstraint.length() - 5));
                    }
                }
            }
        }

        if(sqlMap.get("like_target_key") != null && sqlMap.get("like_target_value") != null) {
//            if(targetConstraint.toString().isEmpty()) {
//                targetConstraint = new StringBuilder(sqlMap.get("like_target_key") + " ilike '%" + sqlMap.get("like_target_value") + "%'");
//            } else {
//                targetConstraint.append(" AND ").append(sqlMap.get("like_target_key")).append(" ilike '%").append(sqlMap.get("like_target_value")).append("%'");
//            }
//            if(isConstraintPresent){
//                targetConstraint.append(" AND ");
//            }
            targetConstraint.append(" AND ").append(sqlMap.get("like_target_key")).append(" ilike '%").append(sqlMap.get("like_target_value")).append("%'");

//            isConstraintPresent = true;
        }else{
            //multiple target
            if(sqlMap.get("like_targets") != null){
                if(sqlMap.get("like_targets") instanceof Map<?,?>){
                    StringBuilder likeTargetConstraint = new StringBuilder();
                    Map<String, Object> likeTargets = (Map<String, Object>) sqlMap.get("like_targets");

                    if(!likeTargets.keySet().isEmpty()) {
                        for (String key : likeTargets.keySet()) {
                            Object value = likeTargets.get(key);
                            if(value == null ){
//                                if(isConstraintPresent){
//                                    likeTargetConstraint.append(" AND ");
//                                }
                                if(includeNullValue){
//                                    likeTargetConstraint.append(key).append(" IS NULL AND ");
                                    likeTargetConstraint.append(" AND ").append(key).append(" IS NULL ");
//                                    isConstraintPresent = true;
                                }else {
                                    continue;
                                }
                            }else {
//                                if(isConstraintPresent){
//                                    likeTargetConstraint.append(" AND ");
//                                }
                                if (value instanceof Integer || value instanceof Double) {
//                                    likeTargetConstraint.append(key).append(" = ").append(likeTargets.get(key)).append(" AND ");
                                    likeTargetConstraint.append(" AND ").append(key).append(" = ").append(likeTargets.get(key));
                                    continue;
                                }
                                if(value instanceof String){
                                    if(((String) value).isEmpty()) {
//                                        likeTargetConstraint.append(key).append(" = '' AND ");
                                        likeTargetConstraint.append(" AND ").append(key).append(" = '' ");
                                        continue;
                                    }
                                }
//                                likeTargetConstraint.append(key).append(" ilike '%").append(likeTargets.get(key)).append("%' AND ");
                                likeTargetConstraint.append(" AND ").append(key).append(" ilike '%").append(likeTargets.get(key)).append("%' ");
//                                isConstraintPresent = true;
                            }
                        }
//                        likeTargetConstraint = new StringBuilder(likeTargetConstraint.substring(0, likeTargetConstraint.length() - 5));
//                        if(targetConstraint.toString().isEmpty()) {
//                            targetConstraint = likeTargetConstraint;
//                        } else {
//                            targetConstraint.append(" AND ").append(likeTargetConstraint);
//                        }
//                        if(isConstraintPresent){
//                            targetConstraint.append(" AND ");
//                        }
                        targetConstraint.append(likeTargetConstraint);

//                        isConstraintPresent = true;
                    }
                }
            }
        }

        if(sqlMap.get("idInConstraint") != null) {
            String idInConstraint = (String) sqlMap.get("idInConstraint");
            if(!idInConstraint.isEmpty()) {
//                if(isConstraintPresent) {
                targetConstraint.append(" AND ").append(idInConstraint);
//                }
//                if (targetConstraint.toString().isEmpty()) {
//                    targetConstraint = new StringBuilder((String) sqlMap.get("idInConstraint"));
//                } else {
//                    targetConstraint.append(" AND ").append(sqlMap.get("idInConstraint"));
//                }
//                isConstraintPresent = true;
            }
        }

        if(sqlMap.get("is_not_null") != null) {
//            if(isConstraintPresent){
//                targetConstraint.append(" AND ");
//            }
            targetConstraint.append(" AND ").append(sqlMap.get("is_not_null")).append(" IS NOT NULL");
//            if(targetConstraint.toString().isEmpty()) {
//                targetConstraint = new StringBuilder(sqlMap.get("is_not_null") + " IS NOT NULL");
//            } else {
//                targetConstraint.append(" AND ").append(sqlMap.get("is_not_null")).append(" IS NOT NULL");
//            }
//            isConstraintPresent = true;
        }

        if(sqlMap.get("is_not_empty") != null) {
//            if(isConstraintPresent){
//                targetConstraint.append(" AND ");
//            }
            targetConstraint.append(" AND ").append(sqlMap.get("is_not_empty")).append(" != ''");
//            if(targetConstraint.toString().isEmpty()) {
//                targetConstraint = new StringBuilder(sqlMap.get("is_not_empty") + " != ''");
//            } else {
//                targetConstraint.append(" AND ").append(sqlMap.get("is_not_empty")).append(" != ''");
//            }
//            isConstraintPresent = true;
        }

        if(sqlMap.get("additional_constraint") != null) {
//            if(isConstraintPresent){
//                targetConstraint.append(" AND ");
//            }
            targetConstraint.append(" AND ").append(sqlMap.get("additional_constraint"));
//            if(targetConstraint.toString().isEmpty()) {
//                targetConstraint = new StringBuilder(String.valueOf(sqlMap.get("additional_constraint")));
//            } else {
//                targetConstraint.append(" AND ").append(sqlMap.get("additional_constraint"));
//            }
//            isConstraintPresent = true;
        }

        StringBuilder timeConstraint = new StringBuilder();
        if(sqlMap.get("time_key") != null) {
            if(sqlMap.get("start_datetime")!=null){
                timeConstraint.append(" AND ").append(sqlMap.get("time_key")).append(" >= '").append(sqlMap.get("start_datetime")).append("'");
//                isConstraintPresent = true;
            }
//            if(!timeConstraint.isEmpty()){
//                timeConstraint.append(" AND ");
//            }
            if(sqlMap.get("end_datetime")!=null){
                timeConstraint.append(" AND ").append(sqlMap.get("time_key")).append(" <= '").append(sqlMap.get("end_datetime")).append("'");
//                isConstraintPresent = true;
            }
        }

        if(sqlMap.get("time_constraint") != null) {
            Map<String, Object> timeConstraintMap = (Map<String, Object>) sqlMap.get("time_constraint");
            for (Map.Entry<String, Object> entry : timeConstraintMap.entrySet()) {
                String timeKey = entry.getKey();
                String startDateTime = (String) ((Map<String, Object>) entry.getValue()).get("start_datetime");
                String endDateTime = (String) ((Map<String, Object>) entry.getValue()).get("end_datetime");
//                timeConstraint = timeKey + " >= '" + startDateTime + "'";
                if(startDateTime != null){
//                    if(!timeConstraint.isEmpty()){
//                        timeConstraint.append(" AND ");
//                    }
                    timeConstraint.append(" AND ").append(timeKey).append(" >= '").append(startDateTime).append("'");
//                    isConstraintPresent = true;
                }
                if(endDateTime != null){
//                    if(!timeConstraint.isEmpty()){
//                        timeConstraint.append(" AND ");
//                    }
                    timeConstraint.append(" AND ").append(timeKey).append(" <= '").append(endDateTime).append("'");
//                    isConstraintPresent = true;
                }
            }
        }
        if(!timeConstraint.toString().isEmpty()){
//            if(isConstraintPresent){
//                targetConstraint.append(" AND ");
//            }
            targetConstraint.append(" AND ").append(timeConstraint);
        }

        if(!targetConstraint.toString().isEmpty() && (!timeConstraint.isEmpty())) {
            sql.append(" WHERE ").append(targetConstraint).append(" AND ").append(timeConstraint);
        } else if(!targetConstraint.toString().isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint);
        } else if(!timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(timeConstraint);
        }

        boolean hasSort = false;
        if(sqlMap.get("sort") != null && sqlMap.get("sort") instanceof Map<?,?>) {
            Map<String, Object> sort = (Map<String, Object>) sqlMap.get("sort");
            if(sort.containsKey("sort_by") && sort.containsKey("sort_order")){
                hasSort = true;
            }
        }
        if(hasSort){
            Map<String, Object> sort = (Map<String, Object>) sqlMap.get("sort");
            sql.append(" ORDER BY ").append(sort.get("sort_by"));
            if(sort.get("sort_order") != null){
                sql.append(" ").append(sort.get("sort_order"));
            }
            sql.append(" NULLS LAST");
        }else if(sqlMap.get("time_key")!=null){
            sql.append(" ORDER BY ").append(sqlMap.get("time_key")).append(" DESC");
            sql.append(" NULLS LAST");
        }

        if(sqlMap.get("limit") != null) {
            sql.append(" LIMIT ").append(sqlMap.get("limit"));
        }
        if(sqlMap.get("offset") != null) {
            sql.append(" OFFSET ").append(sqlMap.get("offset"));
        }

        //filter out duplicate 'AND'
        String sqlStr = sql.toString().replaceAll("AND\\s+AND", "AND");
        //filter out trailing 'AND'
        sqlStr = sqlStr.replaceAll("AND\\s*$", "");
        //filter out WHERE if no constraint
        sqlStr = sqlStr.replaceAll("WHERE\\s+AND", "WHERE");

        return Map.of("sql", sqlStr);
    }
    public static Map<String, String> makeJoinSelectLikeSql(Map<String, Object> sqlMap) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ");
        if(sqlMap.get("select") != null) {
            sql.append(sqlMap.get("select"));
        } else {
            sql.append("*");
        }

        sql.append(" FROM ");
        if(sqlMap.get("from") != null) {
            sql.append(sqlMap.get("from"));
        } else {
            return Map.of("error", "Missing table name");
        }

        if(sqlMap.get("join_table_names") != null) {
            sql.append(" JOIN ").append(sqlMap.get("join_table_names"));
        } else {
            return Map.of("error", "Missing join table name");
        }

        if(sqlMap.get("join_on") != null) {
            sql.append(" ON ").append(sqlMap.get("join_on"));
        } else {
            return Map.of("error", "Missing on condition");
        }

        StringBuilder targetConstraint = new StringBuilder();
        if(sqlMap.get("additional_constraint") != null) {
            if (targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(String.valueOf(sqlMap.get("additional_constraint")));
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("additional_constraint"));
            }
        }

        String timeConstraint = "";
        if(sqlMap.get("time_key") != null) {
            if(sqlMap.get("start_datetime")!=null){
                timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime");
            }
            if(sqlMap.get("end_datetime")!=null){
                if(timeConstraint.isEmpty()){
                    timeConstraint = sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }else{
                    timeConstraint += "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }
            }
        }

        if(!targetConstraint.toString().isEmpty() && !timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint).append(" AND ").append(timeConstraint);
        } else if(!targetConstraint.toString().isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint);
        } else if(!timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(timeConstraint);
        }

        boolean hasSort = false;
        if(sqlMap.get("sort") != null && sqlMap.get("sort") instanceof Map<?,?>) {
            Map<String, Object> sort = (Map<String, Object>) sqlMap.get("sort");
            if(sort.containsKey("sort_by") && sort.containsKey("sort_order")){
                hasSort = true;
            }
        }
        if(hasSort){
            Map<String, Object> sort = (Map<String, Object>) sqlMap.get("sort");
            sql.append(" ORDER BY ").append(sort.get("sort_by"));
            if(sort.get("sort_order") != null){
                sql.append(" ").append(sort.get("sort_order"));
            }
            sql.append(" NULLS LAST");
        }else if(sqlMap.get("time_key")!=null){
            sql.append(" ORDER BY ").append(sqlMap.get("time_key")).append(" DESC");
            sql.append(" NULLS LAST");
        }

        if(sqlMap.get("limit") != null) {
            sql.append(" LIMIT ").append(sqlMap.get("limit"));
        }
        if(sqlMap.get("offset") != null) {
            sql.append(" OFFSET ").append(sqlMap.get("offset"));
        }

        return Map.of("sql", sql.toString());
    }
    public static Map<String, String> makeJoinSelectLikeSql2(Map<String, Object> sqlMap) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ");
        if(sqlMap.get("select") != null) {
            sql.append(sqlMap.get("select"));
        } else {
            sql.append("*");
        }

        sql.append(" FROM ");
        if(sqlMap.get("from") != null) {
            sql.append(sqlMap.get("from"));
        } else {
            return Map.of("error", "Missing table name");
        }

        if(sqlMap.get("join") != null) {
            sql.append(" JOIN ").append(sqlMap.get("join"));
        } else {
            return Map.of("error", "Missing join table name");
        }

        if(sqlMap.get("on") != null) {
            sql.append(" ON ").append(sqlMap.get("on"));
        } else {
            return Map.of("error", "Missing on condition");
        }

        StringBuilder targetConstraint = new StringBuilder();
        boolean includeNullValue = sqlMap.get("include_null_value") != null && sqlMap.get("include_null_value").equals("true");

        //by default, use key = value
        if(sqlMap.get("target_key") != null && sqlMap.get("target_value") != null) {
            targetConstraint = new StringBuilder(sqlMap.get("target_key") + " = '" + sqlMap.get("target_value") + "'");
        }else{
            //multiple target
            if(sqlMap.get("targets") != null){
                if(sqlMap.get("targets") instanceof Map<?,?>){
                    Map<String, Object> targets = (Map<String, Object>) sqlMap.get("targets");
                    if(!targets.keySet().isEmpty()) {
                        for (String key : targets.keySet()) {
//                            targetConstraint.append(key).append(" = '").append(targets.get(key)).append("' AND ");
                            Object value = targets.get(key);
                            if(value == null ) {
                                if (includeNullValue) {
//                                    targetConstraint.append(key).append(" IS NULL AND ");
                                    targetConstraint.append(" ( ").append(key).append(" IS NULL or ").append(key).append(" = '' ) AND ");
                                } else {
                                    continue;
                                }
                            }else {
                                if (value instanceof Integer || value instanceof Double) {
                                    targetConstraint.append(key).append(" = ").append(targets.get(key)).append(" AND ");
                                    continue;
                                }
                                if(value instanceof String){
                                    if(((String) value).isEmpty()) {
                                        targetConstraint.append(key).append(" = '' AND ");
                                        continue;
                                    }
                                }
                                targetConstraint.append(key).append(" = '").append(targets.get(key)).append("' AND ");

                            }
                        }
                        targetConstraint = new StringBuilder(targetConstraint.substring(0, targetConstraint.length() - 5));
                    }
                }
            }
        }

        if(sqlMap.get("like_target_key") != null && sqlMap.get("like_target_value") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("like_target_key") + " ilike '%" + sqlMap.get("like_target_value") + "%'");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("like_target_key")).append(" ilike '%").append(sqlMap.get("like_target_value")).append("%'");
            }
        }else{
            //multiple target
            if(sqlMap.get("like_targets") != null){

                if(sqlMap.get("like_targets") instanceof Map<?,?>){
                    StringBuilder likeTargetConstraint = new StringBuilder();
                    Map<String, Object> likeTargets = (Map<String, Object>) sqlMap.get("like_targets");

                    if(!likeTargets.keySet().isEmpty()) {
                        for (String key : likeTargets.keySet()) {
                            Object value = likeTargets.get(key);
                            if(value == null ){
                                if(includeNullValue){
                                    likeTargetConstraint.append(key).append(" IS NULL AND ");
                                }else {
                                    continue;
                                }
                            }else {
                                if (value instanceof Integer || value instanceof Double) {
                                    likeTargetConstraint.append(key).append(" = ").append(likeTargets.get(key)).append(" AND ");
                                    continue;
                                }
                                if(value instanceof String){
                                    if(((String) value).isEmpty()) {
                                        likeTargetConstraint.append(key).append(" = '' AND ");
                                        continue;
                                    }
                                }
                                likeTargetConstraint.append(key).append(" ilike '%").append(likeTargets.get(key)).append("%' AND ");

                            }
                        }
                        likeTargetConstraint = new StringBuilder(likeTargetConstraint.substring(0, likeTargetConstraint.length() - 5));
                        if(targetConstraint.toString().isEmpty()) {
                            targetConstraint = likeTargetConstraint;
                        } else {
                            targetConstraint.append(" AND ").append(likeTargetConstraint);
                        }
                    }
                }
            }
        }

        if(sqlMap.get("idInConstraint") != null) {
            String idInConstraint = (String) sqlMap.get("idInConstraint");
            if(!idInConstraint.isEmpty()) {
                if (targetConstraint.toString().isEmpty()) {
                    targetConstraint = new StringBuilder((String) sqlMap.get("idInConstraint"));
                } else {
                    targetConstraint.append(" AND ").append(sqlMap.get("idInConstraint"));
                }
            }
        }

        if(sqlMap.get("is_not_null") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("is_not_null") + " IS NOT NULL");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("is_not_null")).append(" IS NOT NULL");
            }
        }

        if(sqlMap.get("is_not_empty") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(sqlMap.get("is_not_empty") + " != ''");
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("is_not_empty")).append(" != ''");
            }
        }

        if(sqlMap.get("additional_constraint") != null) {
            if(targetConstraint.toString().isEmpty()) {
                targetConstraint = new StringBuilder(String.valueOf(sqlMap.get("additional_constraint")));
            } else {
                targetConstraint.append(" AND ").append(sqlMap.get("additional_constraint"));
            }
        }

        String timeConstraint = "";
        if(sqlMap.get("time_key") != null) {
            if(sqlMap.get("start_datetime")!=null){
                timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime");
            }
            if(sqlMap.get("end_datetime")!=null){
                if(timeConstraint.isEmpty()){
                    timeConstraint = sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }else{
                    timeConstraint += "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }
            }
//            timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime") + "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime") + "'";

        }

        if(!targetConstraint.toString().isEmpty() && !timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint).append(" AND ").append(timeConstraint);
        } else if(!targetConstraint.toString().isEmpty()) {
            sql.append(" WHERE ").append(targetConstraint);
        } else if(!timeConstraint.isEmpty()) {
            sql.append(" WHERE ").append(timeConstraint);
        }

        boolean hasSort = false;
        if(sqlMap.get("sort") != null && sqlMap.get("sort") instanceof Map<?,?>) {
            Map<String, Object> sort = (Map<String, Object>) sqlMap.get("sort");
            if(sort.containsKey("sort_by") && sort.containsKey("sort_order")){
                hasSort = true;
            }
        }
        if(hasSort){
            Map<String, Object> sort = (Map<String, Object>) sqlMap.get("sort");
            sql.append(" ORDER BY ").append(sort.get("sort_by"));
            if(sort.get("sort_order") != null){
                sql.append(" ").append(sort.get("sort_order"));
            }
            sql.append(" NULLS LAST");
        }else if(sqlMap.get("time_key")!=null){
            sql.append(" ORDER BY ").append(sqlMap.get("time_key")).append(" DESC");
            sql.append(" NULLS LAST");
        }

        if(sqlMap.get("limit") != null) {
            sql.append(" LIMIT ").append(sqlMap.get("limit"));
        }
        if(sqlMap.get("offset") != null) {
            sql.append(" OFFSET ").append(sqlMap.get("offset"));
        }

        return Map.of("sql", sql.toString());

    }

    public static Map<String, String> makeInsertSql(Map<String, Object> sqlMap) {
        StringBuilder sql = new StringBuilder();
        sql.append("INSERT INTO ");
        if(sqlMap.get("table") != null) {
            sql.append(sqlMap.get("table"));
        } else {
            return Map.of("error", "Missing table name");
        }

        Map<String, Object> content = (Map<String, Object>) sqlMap.get("content");
        if(content == null) {
            return Map.of("error", "Missing content");
        }

        sql.append(" (");
        for(String key : content.keySet()) {
            sql.append(" ").append(key).append(",");
        }
        sql.deleteCharAt(sql.length() - 1);

        sql.append(") VALUES (");
        for(String key : content.keySet()) {
            if(content.get(key) == null) {
                sql.append(" null,");
                continue;
            }

            if(content.get(key) instanceof Integer || content.get(key) instanceof Double) {
                sql.append(" ").append(content.get(key)).append(",");
            } else {
                sql.append(" '").append(content.get(key)).append("',");
            }
            //sql.append(" '").append(content.get(key)).append("',");
        }
        sql.deleteCharAt(sql.length() - 1);

        sql.append(")");

        return Map.of("sql", sql.toString());
    }

    public static Map<String, String> makeUpdateSql(Map<String, Object> sqlMap){
        StringBuilder targetConstraint = new StringBuilder();
        if(sqlMap.get("target_key") != null) {
            targetConstraint = new StringBuilder(sqlMap.get("target_key") + " = '" + sqlMap.get("target_value") + "'");
        }else{
            //multiple target
            if(sqlMap.get("targets") != null){
                if(sqlMap.get("targets") instanceof Map<?,?>){
                    Map<String, String> targets = (Map<String, String>) sqlMap.get("targets");
                    if(!targets.keySet().isEmpty()) {
                        for (String key : targets.keySet()) {
                            targetConstraint.append(key).append(" = '").append(targets.get(key)).append("' AND ");
                        }
                        targetConstraint = new StringBuilder(targetConstraint.substring(0, targetConstraint.length() - 5));
                    }
                }
            }
        }
        //disable wildcard update for data security
        if(targetConstraint.toString().isEmpty()){
            return Map.of("error", "Missing target constraint");
        }

        StringBuilder sql = new StringBuilder();
        sql.append("UPDATE ");
        if(sqlMap.get("table") != null) {
            sql.append(sqlMap.get("table"));
        } else {
            return Map.of("error", "Missing table name");
        }

        Map<String, Object> content = (Map<String, Object>) sqlMap.get("content");
        if(content == null) {
            return Map.of("error", "Missing content");
        }

        sql.append(" SET ");
        for(String key : content.keySet()) {
            if(content.get(key) == null) {
                sql.append(" ").append(key).append(" = null,");
                continue;
            }

            if(content.get(key) instanceof Integer || content.get(key) instanceof Long || content.get(key) instanceof Double) {
                sql.append(" ").append(key).append(" = ").append(content.get(key)).append(",");
            } else {
                if((content.get(key)+"").isEmpty()) {
                    sql.append(" ").append(key).append(" = null,");
                }else {
                    sql.append(" ").append(key).append(" = '").append(content.get(key)).append("',");
                }
            }
        }
        sql.deleteCharAt(sql.length() - 1);

        sql.append(" WHERE ").append(targetConstraint);

        return Map.of("sql", sql.toString());
    }
}
