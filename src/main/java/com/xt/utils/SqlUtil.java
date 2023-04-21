package com.xt.utils;

import java.util.Map;

public class SqlUtil {
    public static Map<String, String> makeSelectSql(Map<String, String> sqlMap) {
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
        }

        String timeConstraint = "";
        if(sqlMap.get("time_key") != null) {
            timeConstraint = sqlMap.get("time_key") + " >= '" + sqlMap.get("start_datetime") + "' AND " + sqlMap.get("time_key") + " <= '" + sqlMap.get("end_datetime") + "'";
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

        String limit = sqlMap.get("limit");
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
            sql.append(" '").append(content.get(key)).append("',");
        }
        sql.deleteCharAt(sql.length() - 1);

        sql.append(")");

        return Map.of("sql", sql.toString());
    }
}
