package org.pabuff.utils;

import org.joda.money.CurrencyUnitDataProvider;

public class XtCurrencyUnitDataProvider extends CurrencyUnitDataProvider {
    @Override
    protected void registerCurrencies() throws Exception {
        // The parameters are: code, numeric code, decimal places
        //ISO 4217 for SGD is 702
        registerCurrency("SGD", 702, 2);
        //currency code for XBT is not standard, use 2009 for now
        registerCurrency("XBT", 2009, 8);
    }
}
