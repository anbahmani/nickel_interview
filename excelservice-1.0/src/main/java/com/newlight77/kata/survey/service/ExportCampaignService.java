package com.newlight77.kata.survey.service;

import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.client.CampaignClient;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@Component
public class ExportCampaignService {

  private final CampaignClient campaignWebService;
  private final MailService mailService;
  private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
  
  public ExportCampaignService(final CampaignClient campaignWebService, final MailService mailService) {
    this.campaignWebService = campaignWebService;
    this.mailService = mailService;
  }

  public void creerSurvey(final Survey survey) {
    campaignWebService.createSurvey(survey);
  }

  public Survey getSurvey(final String id) {
    return campaignWebService.getSurvey(id);
  }

  public void createCampaign(final Campaign campaign) {
    campaignWebService.createCampaign(campaign);
  }

  public Campaign getCampaign(final String id) {
    return campaignWebService.getCampaign(id);
  }

  public void sendResults(final Campaign campaign, final Survey survey) {
    final Workbook workbook = new XSSFWorkbook();

    final Sheet sheet = workbook.createSheet("Survey");
    sheet.setColumnWidth(0, 10500);
    for (int i = 1; i <= 18; i++) {
      sheet.setColumnWidth(i, 6000);
    }

    // 1ere ligne =  l'entête
    final Row header = sheet.createRow(0);

    final CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    final XSSFFont font = ((XSSFWorkbook) workbook).createFont();
    font.setFontName("Arial");
    font.setFontHeightInPoints((short) 14);
    font.setBold(true);
    headerStyle.setFont(font);
    headerStyle.setWrapText(false);

    final Cell headerCell = header.createCell(0);
    headerCell.setCellValue("Survey");
    headerCell.setCellStyle(headerStyle);

    final CellStyle titleStyle = workbook.createCellStyle();
    titleStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
    titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    final XSSFFont titleFont = ((XSSFWorkbook) workbook).createFont();
    titleFont.setFontName("Arial");
    titleFont.setFontHeightInPoints((short) 12);
    titleFont.setUnderline(FontUnderline.SINGLE);
    titleStyle.setFont(titleFont);

    final CellStyle style = workbook.createCellStyle();
    style.setWrapText(true);

    // section client
    Row row = sheet.createRow(2);
    Cell cell = row.createCell(0);
    cell.setCellValue("Client");
    cell.setCellStyle(titleStyle);

    final Row clientRow = sheet.createRow(3);
    final Cell nomClientRowLabel = clientRow.createCell(0);
    nomClientRowLabel.setCellValue(survey.getClient());
    nomClientRowLabel.setCellStyle(style);

    final String clientAddress = survey.getClientAddress().getStreetNumber() + " "
            + survey.getClientAddress().getStreetName() + survey.getClientAddress().getPostalCode() + " "
            + survey.getClientAddress().getCity();

    final Row clientAddressLabelRow = sheet.createRow(4);
    final Cell clientAddressCell = clientAddressLabelRow.createCell(0);
    clientAddressCell.setCellValue(clientAddress);
    clientAddressCell.setCellStyle(style);

    row = sheet.createRow(6);
    cell = row.createCell(0);
    cell.setCellValue("Number of surveys");
    cell = row.createCell(1);
    cell.setCellValue(campaign.getAddressStatuses().size());

    final Row surveyLabelRow = sheet.createRow(8);
    Cell surveyLabelCell = surveyLabelRow.createCell(0);
    surveyLabelCell.setCellValue("N° street");
    surveyLabelCell.setCellStyle(style);

    surveyLabelCell = surveyLabelRow.createCell(1);
    surveyLabelCell.setCellValue("streee");
    surveyLabelCell.setCellStyle(style);

    surveyLabelCell = surveyLabelRow.createCell(2);
    surveyLabelCell.setCellValue("Postal code");
    surveyLabelCell.setCellStyle(style);

    surveyLabelCell = surveyLabelRow.createCell(3);
    surveyLabelCell.setCellValue("City");
    surveyLabelCell.setCellStyle(style);

    surveyLabelCell = surveyLabelRow.createCell(4);
    surveyLabelCell.setCellValue("Status");
    surveyLabelCell.setCellStyle(style);

    final int startIndex = 9;
    int currentIndex = 0;

    for (final AddressStatus addressStatus : campaign.getAddressStatuses()) {

      final Row surveyRow = sheet.createRow(startIndex + currentIndex);
      Cell surveyRowCell = surveyRow.createCell(0);
      surveyRowCell.setCellValue(addressStatus.getAddress().getStreetNumber());
      surveyRowCell.setCellStyle(style);

      surveyRowCell = surveyRow.createCell(1);
      surveyRowCell.setCellValue(addressStatus.getAddress().getStreetName());
      surveyRowCell.setCellStyle(style);

      surveyRowCell = surveyRow.createCell(2);
      surveyRowCell.setCellValue(addressStatus.getAddress().getPostalCode());
      surveyRowCell.setCellStyle(style);

      surveyRowCell = surveyRow.createCell(3);
      surveyRowCell.setCellValue(addressStatus.getAddress().getCity());
      surveyRowCell.setCellStyle(style);

      surveyRowCell = surveyRow.createCell(4);
      surveyRowCell.setCellValue(addressStatus.getStatus().toString());
      surveyRowCell.setCellStyle(style);

      currentIndex++;

    }

    writeFileAndSend(survey, workbook);

  }

  protected void writeFileAndSend(final Survey survey, final Workbook workbook) {
    try {
      final File resultFile = new File(System.getProperty("java.io.tmpdir"), "survey-" + survey.getId() + "-" + dateTimeFormatter.format(LocalDate.now()) + ".xlsx");
      final FileOutputStream outputStream = new FileOutputStream(resultFile);
      workbook.write(outputStream);

      mailService.send(resultFile);
      resultFile.deleteOnExit();
    } catch(final Exception ex) {
        throw new RuntimeException("Errorr while trying to send email", ex);
    } finally {
      try {
        workbook.close();
      } catch(final Exception e) {
        // CANT HAPPEN
      }
    }
  }

}
