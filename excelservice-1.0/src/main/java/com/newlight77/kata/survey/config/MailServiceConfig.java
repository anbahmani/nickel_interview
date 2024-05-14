package com.newlight77.kata.survey.config;

import lombok.Data;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;

@Configuration
@Data
public class MailServiceConfig {

  @Value("${service.mail.from}")
  private String from;

  @Value("${service.mail.to}")
  private String[] to;

}
