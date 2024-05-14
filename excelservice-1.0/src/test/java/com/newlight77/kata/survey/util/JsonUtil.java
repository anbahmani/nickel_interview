package com.newlight77.kata.survey.util;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.ser.impl.SimpleFilterProvider;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;

public class JsonUtil {

    private static final JsonUtil instance = new JsonUtil();
    private final ObjectMapper mapper = initObjectMapper();

    private JsonUtil() { }

    public static JsonUtil instance() {
        return instance;
    }

    private ObjectMapper initObjectMapper() {
        final ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.configure(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS, false)
                .configure(JsonParser.Feature.ALLOW_SINGLE_QUOTES, true)
                .configure(SerializationFeature.INDENT_OUTPUT, true)
                .configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false)
                .enable(DeserializationFeature.READ_ENUMS_USING_TO_STRING)
                .setDateFormat(new SimpleDateFormat("yyyy-MM-dd hh:mm:ss"))
                .setFilterProvider(new SimpleFilterProvider().setFailOnUnknownId(false))
                .setSerializationInclusion(JsonInclude.Include.NON_NULL);
        objectMapper.findAndRegisterModules();
        return objectMapper;
    }

    public String toJson(final Object o) {
        try {
            return mapper.writeValueAsString(o);
        } catch (IOException ioe) {
            String msg = "Error convert to json from object " + o.toString();
            throw new IllegalStateException(msg, ioe);
        }
    }

    public <T> T fromJson(final String input, final Class<T> resourceClass) {
        try {
            return mapper.readValue(input, resourceClass);
        } catch (IOException ioe) {
            String msg = "Error converting from json {" + input + "} " + "to object " + resourceClass.getName();
            throw new IllegalStateException(msg, ioe);
        }
    }

    public <T> T fromJsonFile(final String path, final Class<T> resourceClass) {
        try {
            InputStream is = JsonUtil.class.getResourceAsStream(path);
            return mapper.readValue(is, resourceClass);
        } catch (IOException ioe) {
            String msg = "Error converting from json file " + path + " " + "to object " + resourceClass.getName();
            throw new IllegalStateException(msg, ioe);
        }
    }
}
