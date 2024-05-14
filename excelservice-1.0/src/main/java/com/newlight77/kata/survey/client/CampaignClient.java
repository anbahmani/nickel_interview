package com.newlight77.kata.survey.client;

import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Service;
import org.springframework.web.reactive.function.client.WebClient;

@Service
public class CampaignClient {

    private final WebClient webClient;

    public CampaignClient(@Value("${external.url}") final String externalUrl) {
        webClient = WebClient.builder()
                .baseUrl(externalUrl)
                .defaultHeader(HttpHeaders.CONTENT_TYPE, "application/json")
                .build();
    }

    public void createSurvey(final Survey survey) {
        webClient.post()
                .uri("/surveys")
                .bodyValue(survey)
                .retrieve();
    }

    public Survey getSurvey(final String id) {
        return webClient.get()
                .uri("/surveys/" + id)
                .retrieve()
                .bodyToMono(Survey.class).block();
    }

    public void createCampaign(final Campaign campaign) {
        webClient.post()
                .uri("/campaigns")
                .bodyValue(campaign);
    }

    public Campaign getCampaign(final String id) {
        return webClient.get()
                .uri("/campaigns/" + id)
                .retrieve()
                .bodyToMono(Campaign.class).block();
    }
}
