package com.ubs.pesubjobs;

import org.junit.jupiter.api.Test;

// Boots the full Spring context against Zonky's in-process embedded PostgreSQL (via
// IntegrationTestBase) — no running localhost:5432 instance required.
class PeSubJobsApplicationTests extends IntegrationTestBase {

    @Test
    void contextLoads() {
    }
}
