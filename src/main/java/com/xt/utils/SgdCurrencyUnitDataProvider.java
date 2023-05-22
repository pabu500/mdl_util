package com.xt.utils;

import org.joda.money.CurrencyUnitDataProvider;

public class SgdCurrencyUnitDataProvider extends CurrencyUnitDataProvider {
    @Override
    protected void registerCurrencies() throws Exception {
        // The parameters are: code, numeric code, decimal places
        //ISO 4217 for SGD is 702
        registerCurrency("SGD", 702, 2);
    }
}
