package com.newlight77.kata.survey.controler;

import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.service.ExportCampaignService;
import org.springframework.web.bind.annotation.*;

@RestController
public class SurveyController {

    private final ExportCampaignService exportCampaignService;

    public SurveyController(final ExportCampaignService exportCampaignService) {
      this.exportCampaignService = exportCampaignService;
    }

    @RequestMapping(value = "/api/survey/create", method = RequestMethod.POST)
    public void createSurvey(@RequestBody final Survey survey) {
        exportCampaignService.creerSurvey(survey);
    }

    @RequestMapping(value = "/api/survey/get", method = RequestMethod.GET)
    public Survey getSurvey(@RequestParam final String id) {
        return exportCampaignService.getSurvey(id);
    }

    @RequestMapping(value = "/api/survey/campaign/create", method = RequestMethod.POST)
    public void createCampaign(@RequestBody final Campaign campaign) {
        exportCampaignService.createCampaign(campaign);
    }

    @RequestMapping(value = "/api/survey/campaign/get", method = RequestMethod.GET)
    public Campaign getCampaign(@RequestParam final String id) {
        return exportCampaignService.getCampaign(id);
    }

    @RequestMapping(value = "/api/survey/campaign/export", method = RequestMethod.POST)
    public void exportCampaign(@RequestParam final String campaignId) {

        final Campaign campaign = exportCampaignService.getCampaign(campaignId);
        final Survey survey = exportCampaignService.getSurvey(campaign.getSurveyId());
        exportCampaignService.sendResults(campaign, survey);
        
    }
}

