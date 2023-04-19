package com.xt.utils;

import java.util.regex.Pattern;

public class StringValidate {
    // This file contains all the global variables and public staticants used in the app
    public static int maxLoginNameLength = 30;
    public static int maxFullNameLength = 50;
    public static int maxEmailLength = 250;
    public static int maxPhoneLength = 20;
    public static int maxPasswordLength = 30;
    public static int minPasswordLength = 8;

    public static String glb_reg_allkeyboard =
                "^[a-zA-Z0-9!@#\\$%^&*()_+\\-=\\[\\]{};':\\|,.<>\\/?\s]+$";

    // ^ matches the start of the string
    // [a-zA-Z]{2,} matches 2 or more consecutive letters (lowercase or uppercase)
    // [a-zA-Z\s]* matches zero or more consecutive letters (lowercase or uppercase) or whitespace characters
    // [a-zA-Z]{2,}$ matches 2 or more consecutive letters (lowercase or uppercase) and matches the end of the string
    public static String glb_reg_fullName = "^[a-zA-Z]{2,}[a-zA-Z\s]*[a-zA-Z]{2,}$";
    public static String glb_fullName_callout = "Please provide a valid full name";
    public static String glb_regNG_fullName = "[^a-zA-Z '\\.]|([.']| )(\\.|')|(\s)(\s)";
    // r"^
    // [a-zA-Z0-9-]  //matches any character from the alphabet (lowercase or uppercase), digits, or the hyphen symbol
    // {6,} //matches the preceding expression (a character from the alphabet or a hyphen symbol) at least 6 times
    // $
    public static String glb_reg_loginName = "^[a-zA-Z0-9-]{6,}$";
    public static String glb_loginName_callout =
                "Please provide a valid username (alphabets, numbers and - only, at least 6 characters)";
    public static String glb_regNG_loginName = "[^a-zA-Z0-9-]+";

    public static String glb_reg_email = "^[a-zA-Z0-9.]+@[a-zA-Z0-9]+\\.[a-zA-Z]+";
    public static String glb_email_callout = "Please provide a valid email address";
    public static String glb_regNG_email = "[^a-zA-Z0-9@. ]+|([@\\.])([@\\.])";

    /* for test https://regex101.com/
    95687951
    6595687951
    +6595687951
    +65 95687951
    +8613851610359
    +86 13851610359
    +60 1223568874
    */
    public static String glb_reg_phone = "^\\+?[0-9]{0,3}[\s]*[0-9]{0,3}[\s]*[0-9]{8,13}$";
    public static String glb_phone_callout = "Please provide a valid phone number";
    public static String glb_regNG_phone = "[^0-9+ ]|([0-9+ ])([+])|(\s)(\s)";

    // r'^
    //   (?=.*[A-Z])       // should contain at least one upper case
    //   (?=.*[a-z])       // should contain at least one lower case
    //   (?=.*?[0-9])      // should contain at least one digit
    //   (?=.*?[!@#\$&*~]) // should contain at least one Special character
    //   .{8,}             // Must be at least 8 characters in length
    // $
    public static String glb_reg_password =
                "^(?=.*?[A-Z])(?=.*?[a-z])(?=.*?[0-9])(?=.{8,}$)";
    public static String glb_password_callout =
                "Please provide a valid password (at least 8 characters, at lease 1 upper case, 1 lower case and 1 digit)";
    public static String glb_regNG_password = "[^a-zA-Z0-9!@#\\$&*~]+";

    public static String validate(String input, String reg, String regNG, String callout){
        if(input == null){
            return callout;
        }
        if(input.length() == 0){
            return callout;
        }
        if(!input.matches(reg)){
            return callout;
        }
        if(input.matches(regNG)){
            return callout;
        }
        return null;
    }
    public static String validateLoginName(String input){
        return validate(input, glb_reg_loginName, glb_regNG_loginName, glb_loginName_callout);
    }
    public static String validateFullName(String input){
        return validate(input, glb_reg_fullName, glb_regNG_fullName, glb_fullName_callout);
    }
    public static String validateEmail(String input){
        return validate(input, glb_reg_email, glb_regNG_email, glb_email_callout);
    }
    public static String validatePhone(String input){
        return validate(input, glb_reg_phone, glb_regNG_phone, glb_phone_callout);
    }
    public static String validatePassword(String input){
        //regex string not working, use break down method
//        return validate(input, glb_reg_password, glb_regNG_password, glb_password_callout);
        boolean lengthOK = input.length()>=8 && input.length()<=50;

        Pattern alphaCap = Pattern.compile("[A-Z]");
        boolean hasAlphaCap = alphaCap.matcher(input).find();
        Pattern alphaLow = Pattern.compile("[a-z]");
        boolean hasAlphaLow = alphaLow.matcher(input).find();
        Pattern digit = Pattern.compile("[0-9]");
        boolean hasDigit = digit.matcher(input).find();
//        Pattern special = Pattern.compile("[\\~\\!\\@\\#\\$\\%\\^\\&\\*\\(\\)_+\\{\\}\\[\\]\\?<>|_]");
//        boolean hasSpecial = special.matcher(input).find();
        return lengthOK && hasAlphaCap && hasAlphaLow && hasDigit ? null : glb_email_callout;

    }
}
