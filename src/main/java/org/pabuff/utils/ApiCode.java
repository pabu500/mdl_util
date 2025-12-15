package org.pabuff.utils;

public enum ApiCode {
    REQUEST_MISSING_PARAMETER,
    REQUEST_INVALID_PARAMETER,

    ACL_UNAUTHORIZED,

    RESULT_GENERIC_ERROR,
    RESULT_DATABASE_ERROR,
    RESULT_NOT_FOUND,
    RESULT_TIMEOUT,

    BL_GENERIC_ERROR, // business logic generic error
}
