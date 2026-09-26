package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.health.PeSubApiReadinessHealthIndicator;
import org.junit.jupiter.api.Test;

import java.time.Duration;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class PeSubApiReadinessHealthIndicatorTest {

    private static final String API_BASE_URL = "http://pe-sub-api:9062/pe-sub-api";

    private final PeSubApiClient apiClient = mock(PeSubApiClient.class);
    private final IngestProperties ingestProperties = new IngestProperties(
            "facilities.csv", "umbrellas.csv", "lp_master.csv", "lp_facility_seeds.csv", "",
            API_BASE_URL, "data/out", 10, Duration.ofSeconds(30), Duration.ofSeconds(2), true);
    private final PeSubApiReadinessHealthIndicator indicator =
            new PeSubApiReadinessHealthIndicator(apiClient, ingestProperties);

    @Test
    void apiReady_reportsUpWithApiBaseUrlDetail() {
        when(apiClient.isApiReady()).thenReturn(true);

        assertThat(indicator.health().getStatus().getCode()).isEqualTo("UP");
        assertThat(indicator.health().getDetails()).containsEntry("apiBaseUrl", API_BASE_URL);
    }

    @Test
    void apiUnavailable_reportsDownWithApiBaseUrlDetail() {
        when(apiClient.isApiReady()).thenReturn(false);

        assertThat(indicator.health().getStatus().getCode()).isEqualTo("DOWN");
        assertThat(indicator.health().getDetails()).containsEntry("apiBaseUrl", API_BASE_URL);
    }
}