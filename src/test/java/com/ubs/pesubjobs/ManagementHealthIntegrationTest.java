package com.ubs.pesubjobs;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.boot.resttestclient.TestRestTemplate;
import org.springframework.boot.resttestclient.autoconfigure.AutoConfigureTestRestTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.web.server.LocalManagementPort;
import org.springframework.http.ResponseEntity;
import org.springframework.test.context.TestPropertySource;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.when;

@TestPropertySource(properties = "management.server.port=0")
@AutoConfigureTestRestTemplate
class ManagementHealthIntegrationTest extends IntegrationTestBase {

    @LocalManagementPort
    int managementPort;

    @Autowired
    TestRestTemplate restTemplate;

    @BeforeEach
    void apiIsUnavailable() {
        when(apiClient.isApiReady()).thenReturn(false);
    }

    @Test
    void managementHealthIsAnonymousAndReportsLivenessUpReadinessDown() {
        ResponseEntity<String> liveness = restTemplate.getForEntity(
                "http://localhost:{port}/manage/health/liveness", String.class, managementPort);
        ResponseEntity<String> readiness = restTemplate.getForEntity(
                "http://localhost:{port}/manage/health/readiness", String.class, managementPort);

        assertThat(liveness.getStatusCode().value())
                .as("liveness response: %s %s", liveness.getStatusCode(), liveness.getBody())
                .isEqualTo(200);
        assertThat(liveness.getBody()).containsPattern("\"status\"\\s*:\\s*\"UP\"");
        assertThat(readiness.getStatusCode().value()).isEqualTo(503);
        assertThat(readiness.getBody()).containsPattern("\"status\"\\s*:\\s*\"DOWN\"");
    }
}