package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.health.PeSubApiReadinessHealthIndicator;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class PeSubApiReadinessHealthIndicatorTest {

    private final PeSubApiClient apiClient = mock(PeSubApiClient.class);
    private final PeSubApiReadinessHealthIndicator indicator =
            new PeSubApiReadinessHealthIndicator(apiClient);

    @Test
    void apiReady_reportsUpWithoutDetails() {
        when(apiClient.isApiReady()).thenReturn(true);

        assertThat(indicator.health().getStatus().getCode()).isEqualTo("UP");
        assertThat(indicator.health().getDetails()).isEmpty();
    }

    @Test
    void apiUnavailable_reportsDownWithoutDetails() {
        when(apiClient.isApiReady()).thenReturn(false);

        assertThat(indicator.health().getStatus().getCode()).isEqualTo("DOWN");
        assertThat(indicator.health().getDetails()).isEmpty();
    }
}