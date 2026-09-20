package com.ubs.pesubjobs.health;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.config.IngestProperties;
import org.springframework.boot.health.contributor.Health;
import org.springframework.boot.health.contributor.HealthIndicator;
import org.springframework.stereotype.Component;

@Component("peSubApi")
public class PeSubApiReadinessHealthIndicator implements HealthIndicator {

    private final PeSubApiClient apiClient;
    private final IngestProperties ingestProperties;

    public PeSubApiReadinessHealthIndicator(PeSubApiClient apiClient, IngestProperties ingestProperties) {
        this.apiClient = apiClient;
        this.ingestProperties = ingestProperties;
    }

    @Override
    public Health health() {
        Health.Builder builder = apiClient.isApiReady() ? Health.up() : Health.down();
        return builder.withDetail("apiBaseUrl", ingestProperties.apiBaseUrl()).build();
    }
}