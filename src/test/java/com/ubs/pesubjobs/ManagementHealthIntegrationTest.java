package com.ubs.pesubjobs;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.boot.resttestclient.TestRestTemplate;
import org.springframework.boot.resttestclient.autoconfigure.AutoConfigureTestRestTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.web.server.LocalServerPort;
import org.springframework.http.ResponseEntity;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.when;

@AutoConfigureTestRestTemplate
class ManagementHealthIntegrationTest extends IntegrationTestBase {

    @LocalServerPort
    int port;

    @Autowired
    TestRestTemplate restTemplate;

    @BeforeEach
    void apiIsUnavailable() {
        when(apiClient.isApiReady()).thenReturn(false);
    }

    @Test
    void managementHealthIsAnonymousAndReportsLivenessUpReadinessDown() {
        ResponseEntity<String> liveness = restTemplate.getForEntity(
                "http://localhost:{port}/pe-sub-jobs/actuator/health/liveness", String.class, port);
        ResponseEntity<String> readiness = restTemplate.getForEntity(
                "http://localhost:{port}/pe-sub-jobs/actuator/health/readiness", String.class, port);

        assertThat(liveness.getStatusCode().value())
                .as("liveness response: %s %s", liveness.getStatusCode(), liveness.getBody())
                .isEqualTo(200);
        assertThat(liveness.getBody()).containsPattern("\"status\"\\s*:\\s*\"UP\"");
        assertThat(readiness.getStatusCode().value()).isEqualTo(503);
        assertThat(readiness.getBody()).containsPattern("\"status\"\\s*:\\s*\"DOWN\"");
    }
}