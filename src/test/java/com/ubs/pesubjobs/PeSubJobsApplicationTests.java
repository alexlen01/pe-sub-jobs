package com.ubs.pesubjobs;

import org.junit.jupiter.api.Test;

// Boots the full DB-less Spring context (ResourcelessJobRepository, mocked PeSubApiClient via
// IntegrationTestBase) — no database and no running pe-sub-api required.
class PeSubJobsApplicationTests extends IntegrationTestBase {

    @Test
    void contextLoads() {
    }
}
