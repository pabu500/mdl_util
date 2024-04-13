package org.pabuff.utils;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class LocalDateTimeMultiDeserializer extends JsonDeserializer<LocalDateTime>{
    @Override
    public LocalDateTime deserialize(JsonParser jsonParser, DeserializationContext deserializationContext) throws IOException, IOException {
        String timestampString = jsonParser.readValueAs(String.class);
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss[.SSS]");
            return LocalDateTime.parse(timestampString, formatter);
        } catch (Exception e) {
            // If the custom pattern parsing fails, try ISO_LOCAL_DATE_TIME
            return LocalDateTime.parse(timestampString, DateTimeFormatter.ISO_LOCAL_DATE_TIME);
        }
    }
}
