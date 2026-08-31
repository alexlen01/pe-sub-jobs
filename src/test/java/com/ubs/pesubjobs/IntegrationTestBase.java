package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient;
import org.springframework.boot.webmvc.test.autoconfigure.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;
import org.springframework.test.context.bean.override.mockito.MockitoBean;

// pe-sub-jobs is DB-less: every feed writes through pe-sub-api's REST endpoints and Spring
// Batch runs on an in-memory ResourcelessJobRepository, so tests need no database at all.
// The PeSubApiClient is replaced with a Mockito mock — job tests run the real reader/processor
// pipeline against temp CSV files and assert the payloads handed to the API client.
//
// Startup ingest is disabled here: context-load tests validate wiring only; job behaviour is
// covered by dedicated tests that launch jobs explicitly with their own feed files.
@SpringBootTest(webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
@AutoConfigureMockMvc
@TestPropertySource(properties = {
    "ingest.run-on-startup=false",
    "bb-template-import.enabled=false",
    // Caller-named feed files resolve against a committed fixture directory rather than data/out,
    // which is generated output and absent on a clean checkout. Set here, not per test, so every
    // test shares one application context.
    "ingest.import-root=src/test/resources/feeds"
})
public abstract class IntegrationTestBase {

    @MockitoBean
    protected PeSubApiClient apiClient;
}
