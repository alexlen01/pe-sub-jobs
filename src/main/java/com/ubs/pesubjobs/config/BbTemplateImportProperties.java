package com.ubs.pesubjobs.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.DefaultValue;

import java.time.Duration;

@ConfigurationProperties(prefix = "bb-template-import")
public record BbTemplateImportProperties(
        @DefaultValue("true") boolean enabled,
        @DefaultValue("data/bb-templates") String directory,
        String apiBaseUrl,
        @DefaultValue("30s") Duration scanInterval,
        @DefaultValue("2s") Duration stableAge) {}
