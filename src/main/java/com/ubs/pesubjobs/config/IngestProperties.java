package com.ubs.pesubjobs.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.DefaultValue;

@ConfigurationProperties(prefix = "ingest")
public record IngestProperties(
        String facilityFile,
        String lpMasterFile,
        String lpFacilitySeedsFile,
        // When false, JobStartupRunner does not launch the seed jobs on boot. Defaults true to
        // preserve production behaviour; tests set it false so jobs don't run against the empty
        // embedded schema (the business tables are owned by pe-sub-api's migrations).
        @DefaultValue("true") boolean runOnStartup) {}