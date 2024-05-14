package com.newlight77.kata.survey.controller;

import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.service.ExportCampaignService;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/survey")
public class SurveyController {

    private final ExportCampaignService exportCampaignService;

    public SurveyController(final ExportCampaignService exportCampaignService) {
      this.exportCampaignService = exportCampaignService;
    }

    @PostMapping("/create")
    public void createSurvey(@RequestBody final Survey survey) {
        exportCampaignService.createSurvey(survey);
    }

    @GetMapping("/get")
    public ResponseEntity<Survey> getSurvey(@RequestParam final String id) {
        Survey survey = exportCampaignService.getSurvey(id);
        if (survey == null) {
            return ResponseEntity.notFound().build();
        }
        return ResponseEntity.ok(survey);
    }

    @PostMapping("/campaign/create")
    public void createCampaign(@RequestBody final Campaign campaign) {
        exportCampaignService.createCampaign(campaign);
    }

    @GetMapping("/campaign/get")
    public ResponseEntity<Campaign> getCampaign(@RequestParam final String id) {
        Campaign campaign = exportCampaignService.getCampaign(id);
        if (campaign == null) {
            return ResponseEntity.notFound().build();
        }
        return ResponseEntity.ok(campaign);
    }

    @PostMapping("/campaign/export")
    public void exportCampaign(@RequestParam final String campaignId) {
        this.exportCampaignService.exportCampaign(campaignId);
    }
}

