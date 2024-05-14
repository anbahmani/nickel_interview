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
  private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd");
  
  public ExportCampaignService(final CampaignClient campaignWebService, final MailService mailService) {
    this.campaignWebService = campaignWebService;
    this.mailService = mailService;
  }

  public void exportCampaign(String campaignId) {
    final Campaign campaign = getCampaign(campaignId);
    final Survey survey = getSurvey(campaign.getSurveyId());
    sendResults(campaign, survey);
  }

  public void createSurvey(final Survey survey) {
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

  public void sendResults(Campaign campaign, Survey survey) {
    Workbook workbook = new XSSFWorkbook();
    try {
      createSurveySheet(workbook, campaign, survey);
      writeFileAndSend(survey, workbook);
    } finally {
      try {
        workbook.close();
      } catch (Exception e) {
        throw new RuntimeException("Error in sending survey results", e);
      }
    }
  }

  private void createSurveySheet(Workbook workbook, Campaign campaign, Survey survey) {
    Sheet sheet = workbook.createSheet("Survey");
    setColumnWidths(sheet);
    createHeaderRow(workbook, sheet);
    createClientSection(workbook, sheet, survey, campaign);
    createSurveySection(workbook, sheet, campaign);
  }

  private void createHeaderRow(Workbook workbook, Sheet sheet) {
    Row header = sheet.createRow(0);
    Cell headerCell = header.createCell(0);
    headerCell.setCellValue("Survey");
    headerCell.setCellStyle(createHeaderStyle(workbook));
  }

  private CellStyle createHeaderStyle(Workbook workbook) {
    CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    XSSFFont font = ((XSSFWorkbook) workbook).createFont();
    font.setFontName("Arial");
    font.setFontHeightInPoints((short) 14);
    font.setBold(true);
    headerStyle.setFont(font);
    headerStyle.setWrapText(false);
    return headerStyle;
  }

  private CellStyle createTitleStyle(Workbook workbook) {
    CellStyle titleStyle = workbook.createCellStyle();
    titleStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
    titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    XSSFFont titleFont = ((XSSFWorkbook) workbook).createFont();
    titleFont.setFontName("Arial");
    titleFont.setFontHeightInPoints((short) 12);
    titleFont.setUnderline(FontUnderline.SINGLE);
    titleStyle.setFont(titleFont);
    return titleStyle;
  }

  private void createClientSection(Workbook workbook, Sheet sheet, Survey survey, Campaign campaign) {
    CellStyle titleStyle = createTitleStyle(workbook);
    CellStyle style = workbook.createCellStyle();
    style.setWrapText(true);

    Row clientRow = sheet.createRow(2);
    Cell clientLabelCell = clientRow.createCell(0);
    clientLabelCell.setCellValue("Client");
    clientLabelCell.setCellStyle(titleStyle);

    Row clientDataRow = sheet.createRow(3);
    Cell clientNameCell = clientDataRow.createCell(0);
    clientNameCell.setCellValue(survey.getClient());
    clientNameCell.setCellStyle(style);

    String clientAddress = survey.getClientAddress().getStreetNumber() + " "
            + survey.getClientAddress().getStreetName() + survey.getClientAddress().getPostalCode() + " "
            + survey.getClientAddress().getCity();

    Row clientAddressLabelRow = sheet.createRow(4);
    Cell clientAddressCell = clientAddressLabelRow.createCell(0);
    clientAddressCell.setCellValue(clientAddress);
    clientAddressCell.setCellStyle(style);

    clientRow = sheet.createRow(6);
    clientLabelCell = clientRow.createCell(0);
    clientLabelCell.setCellValue("Number of surveys");
    clientLabelCell = clientRow.createCell(1);
    clientLabelCell.setCellValue(campaign.getAddressStatuses().size());
  }

  private void createSurveySection(Workbook workbook, Sheet sheet, Campaign campaign) {
    CellStyle style = createWrapTextStyle(workbook);
    createHeaderRowForSurveyData(sheet, style);
    populateSurveyData(sheet, campaign, style);
  }

  private CellStyle createWrapTextStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setWrapText(true);
    return style;
  }

  private void createHeaderRowForSurveyData(Sheet sheet, CellStyle style) {
    String[] headers = {"N° street", "street", "Postal code", "City", "Status"};
    Row headerRow = sheet.createRow(8);
    for (int i = 0; i < headers.length; i++) {
      Cell cell = headerRow.createCell(i);
      cell.setCellValue(headers[i]);
      cell.setCellStyle(style);
    }
  }

  private void populateSurveyData(Sheet sheet, Campaign campaign, CellStyle style) {
    int rowIndex = 9;
    for (AddressStatus addressStatus : campaign.getAddressStatuses()) {
      Row row = sheet.createRow(rowIndex++);
      populateRowWithData(row, addressStatus, style);
    }
  }

  private void populateRowWithData(Row row, AddressStatus addressStatus, CellStyle style) {
    String[] data = {
            addressStatus.getAddress().getStreetNumber(),
            addressStatus.getAddress().getStreetName(),
            addressStatus.getAddress().getPostalCode(),
            addressStatus.getAddress().getCity(),
            addressStatus.getStatus().toString()
    };

    for (int i = 0; i < data.length; i++) {
      Cell cell = row.createCell(i);
      cell.setCellValue(data[i]);
      cell.setCellStyle(style);
    }
  }

  private void setColumnWidths(Sheet sheet) {
    sheet.setColumnWidth(0, 10500);
    for (int i = 1; i <= 18; i++) {
      sheet.setColumnWidth(i, 6000);
    }
  }

  protected void writeFileAndSend(Survey survey, Workbook workbook) {
    File resultFile = new File(System.getProperty("java.io.tmpdir"), "survey-" + survey.getId() + "-" + DATE_TIME_FORMATTER.format(LocalDate.now()) + ".xlsx");
    try (FileOutputStream outputStream = new FileOutputStream(resultFile)) {
      workbook.write(outputStream);
      mailService.send(resultFile);
    } catch (Exception ex) {
      throw new RuntimeException("Error while trying to send email", ex);
    } finally {
      resultFile.deleteOnExit();
    }
  }

}
