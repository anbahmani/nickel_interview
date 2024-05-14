package com.newlight77.kata.survey.model;

public enum Status {
    TODO,
    DONE
    ;

    public String value() {
        return name();
    }

    public static Status fromValue(String v) {
        return valueOf(v);
    }
}
