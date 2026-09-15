package com.ubs.pesubjobs;

import com.ubs.pesubjobs.model.ProcessedLpMaster;
import org.junit.jupiter.api.AfterEach;
import org.mockito.ArgumentMatchers;
import org.junit.jupiter.api.Test;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.http.HttpStatus;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.web.client.HttpClientErrorException;
import org.springframework.web.client.HttpServerErrorException;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.ArgumentMatchers.anyList;
import static org.mockito.Mockito.doThrow;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

/**
 * pe-sub-api can refuse a whole chunk rather than individual rows — it is down, it rejects the
 * payload, or the service identity is not accepted. A non-2xx there throws in the client, and what
 * matters is that the throw ends the run visibly instead of being absorbed into a COMPLETED that
 * an operator would stop looking at.
 */
class ApiRefusalDuringFeedTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("facilityIngestJob") Job facilityIngestJob;
    @Autowired @Qualifier("lpMasterIngestJob") Job lpMasterIngestJob;
    @Autowired MockMvc mockMvc;

    private Path feedFile;

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void aServerErrorFromTheApi_failsTheRun() throws Exception {
        when(apiClient.ingestFacilities(anyList()))
                .thenThrow(HttpServerErrorException.create(
                        HttpStatus.INTERNAL_SERVER_ERROR, "Server Error", null, null, null));

        JobExecution execution = runFacilityFeed(3);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
    }

    /**
     * A rejected identity is the failure most likely to look like success from a distance: nothing
     * is written and every row is "sent". The run has to fail on it.
     */
    @Test
    void aRejectedServiceIdentity_failsTheRun() throws Exception {
        when(apiClient.ingestFacilities(anyList()))
                .thenThrow(HttpClientErrorException.create(
                        HttpStatus.UNAUTHORIZED, "Unauthorized", null, null, null));

        JobExecution execution = runFacilityFeed(3);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
    }

    /** The caller who triggered the feed is told, not just the service log. */
    @Test
    void theTriggerResponseReportsTheFailureRatherThanAnEmptySuccess() throws Exception {
        when(apiClient.ingestFacilities(anyList()))
                .thenThrow(HttpServerErrorException.create(
                        HttpStatus.SERVICE_UNAVAILABLE, "Service Unavailable", null, null, null));

        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "facilities.csv"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.status").value("FAILED"))
                .andExpect(jsonPath("$.createdCount").value(0))
                .andExpect(jsonPath("$.updatedCount").value(0));
    }

    /**
     * The LP Master feed clears before it repopulates. If the clear itself is refused, the run must
     * stop there — attempting the repopulate against an API that just rejected the caller would
     * report progress against a table nobody has established the state of.
     */
    @Test
    void aRefusedClear_stopsBeforeAnyRepopulate() throws Exception {
        doThrow(HttpClientErrorException.create(HttpStatus.FORBIDDEN, "Forbidden", null, null, null))
                .when(apiClient).clearLpMaster();

        JobExecution execution = runLpMasterFeed(2);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
        verify(apiClient, never()).ingestLpMaster(anyLpMasterRows());
    }

    /**
     * The clear succeeded and the repopulate did not: LP Master is now empty or partial. That is
     * the worst outcome the feed can produce, so it must be reported as a failure and never as a
     * completed run with a low write count.
     */
    @Test
    void aRefusedRepopulateAfterAClear_failsTheRunLoudly() throws Exception {
        when(apiClient.ingestLpMaster(anyLpMasterRows()))
                .thenThrow(HttpServerErrorException.create(
                        HttpStatus.BAD_GATEWAY, "Bad Gateway", null, null, null));

        JobExecution execution = runLpMasterFeed(2);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
        verify(apiClient).clearLpMaster();
    }

    /** Any chunk of rows — these assertions are about whether the call happened, not its contents. */
    private static List<ProcessedLpMaster> anyLpMasterRows() {
        return ArgumentMatchers.any();
    }

    private JobExecution runFacilityFeed(int rows) throws Exception {
        StringBuilder csv = new StringBuilder(
                "\"agent_bank\",\"name\",\"account_number\",\"facility_size\",\"maturity_date\","
                + "\"status\",\"status_date\",\"ubs_participation\",\"collateral_date\"\n");
        for (int i = 1; i <= rows; i++) {
            csv.append("\"Bank of America\",\"Facility %d\",\"ACC%d\",\"1000\",\"2026-01-01\",\"Active\",\"2026-01-01\",\"1\",\"2026-01-01\"\n"
                    .formatted(i, i));
        }
        return run(facilityIngestJob, "refusal-facilities", csv.toString());
    }

    private JobExecution runLpMasterFeed(int rows) throws Exception {
        StringBuilder csv = new StringBuilder(
                "\"investor_name\",\"parent\",\"spv\",\"investor_type\",\"institutional_or_hnw\","
                + "\"region_location\",\"investment_grade\",\"sp_rating\",\"moodys_rating\",\"fitch_rating\","
                + "\"aum\",\"nav\",\"pension_assets\",\"funding_ratio\",\"ubs_lp_category\","
                + "\"ubs_default_advance_rate\",\"ubs_default_concentration_limit\",\"notes\"\n");
        for (int i = 1; i <= rows; i++) {
            csv.append("\"Investor ").append(i).append("\",\"\",\"FALSE\",\"Pension Fund\",")
               .append("\"Institutional\",\"United States\",\"TRUE\",\"AA\",\"Aa2\",\"AA\",")
               .append("\"$10B\",\"$8B\",\"$9B\",\"105%\",\"Rated Investor\",\"90%\",\"5%\",\"\"\n");
        }
        return run(lpMasterIngestJob, "refusal-lp-master", csv.toString());
    }

    private JobExecution run(Job job, String prefix, String csv) throws Exception {
        feedFile = Files.createTempFile(prefix, ".csv");
        Files.writeString(feedFile, csv);
        return jobOperator.start(job, new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters());
    }
}
