package com.ubs.pesubjobs.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

@ConfigurationProperties(prefix = "ingest")
public record IngestProperties(
        String facilityFile,
        String lpMasterFile,
        String lpFacilitySeedsFile) {}