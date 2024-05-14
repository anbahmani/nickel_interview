package com.newlight77.kata.survey.model;

import lombok.Data;

import java.util.List;

@Data
public class Campaign {
    private String id ;
    private String surveyId;
    private List<AddressStatus> addressStatuses;
}
