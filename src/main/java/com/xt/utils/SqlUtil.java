package com.xt.utils;

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

        String targetConstraint = "";
        if(sqlMap.get("target_key") != null) {
            targetConstraint = sqlMap.get("target_key") + " = '" + sqlMap.get("target_value") + "'";
        }else{
            //multiple target
            if(sqlMap.get("targets") != null){
                if(sqlMap.get("targets") instanceof Map<?,?>){
                    Map<String, String> targets = (Map<String, String>) sqlMap.get("targets");
                    for(String key : targets.keySet()){
                        targetConstraint += key + " = '" + targets.get(key) + "' AND ";
                    }
                    targetConstraint = targetConstraint.substring(0, targetConstraint.length()-5);
                }
            }
        }

        String timeConstraint = "";
        if(sqlMap.get("time_key") != null) {
            if(sqlMap.get("start_datetime")!=null){
                timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime");
            }
            if(sqlMap.get("end_datetime")!=null){
                if(timeConstraint.equals("")){
                    timeConstraint = sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }else{
                    timeConstraint += "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime");
                }
            }
//            timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime") + "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime") + "'";
        }

        if(!targetConstraint.equals("") && !timeConstraint.equals("")) {
            sql.append(" WHERE ").append(targetConstraint).append(" AND ").append(timeConstraint);
        } else if(!targetConstraint.equals("")) {
            sql.append(" WHERE ").append(targetConstraint);
        } else if(!timeConstraint.equals("")) {
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
}
