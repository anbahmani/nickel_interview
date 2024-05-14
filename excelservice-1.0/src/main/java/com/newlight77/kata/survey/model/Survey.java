package com.newlight77.kata.survey.model;

import lombok.Data;

import java.util.List;

@Data
public class Survey {
    private String id;
    private String sommary;
    private String client;
    private Address clientAddress;
    private List<Question> questions;
}

