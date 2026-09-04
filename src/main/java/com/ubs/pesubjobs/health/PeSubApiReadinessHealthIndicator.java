package com.ubs.pesubjobs.health;

import com.ubs.pesubjobs.client.PeSubApiClient;
import org.springframework.boot.health.contributor.Health;
import org.springframework.boot.health.contributor.HealthIndicator;
import org.springframework.stereotype.Component;

@Component("peSubApi")
public class PeSubApiReadinessHealthIndicator implements HealthIndicator {

    private final PeSubApiClient apiClient;

    public PeSubApiReadinessHealthIndicator(PeSubApiClient apiClient) {
        this.apiClient = apiClient;
    }

    @Override
    public Health health() {
        return apiClient.isApiReady() ? Health.up().build() : Health.down().build();
    }
}