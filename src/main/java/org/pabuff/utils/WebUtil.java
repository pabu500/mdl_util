package org.pabuff.utils;

//Spring Boot 3.0+
//import jakarta.servlet.http.HttpServletRequest;
//Spring Boot 2.5+
import javax.servlet.http.HttpServletRequest;

public class WebUtil {

    //get the site URL from HttpServletRequest
    public static String getSiteURL(HttpServletRequest request) {
        String siteURL = request.getRequestURL().toString();
        return siteURL.replace(request.getServletPath(), "");
    }
}
