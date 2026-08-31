package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient.ApiIngestSummary;
import com.ubs.pesubjobs.config.IngestTallyListener;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.test.web.servlet.MockMvc;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.ArgumentMatchers.anyList;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

/**
 * A feed's write count says what was posted, not what the platform holds. pe-sub-api owns the
 * constraints, so it can accept a chunk and still refuse rows inside it — and those rows are the
 * ones that matter, because an operator who sees {@code COMPLETED} stops looking. The run has to
 * report what landed, and fail outright once too much of the feed did not.
 */
class IngestTallyTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("facilityIngestJob") Job facilityIngestJob;
    @Autowired MockMvc mockMvc;

    private Path feedFile;

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void whatTheApiReported_isCarriedOntoTheRun() throws Exception {
        when(apiClient.ingestFacilities(anyList())).thenReturn(new ApiIngestSummary(2, 1, 0));

        JobExecution execution = runFeed(3);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        assertThat(execution.getExecutionContext().getLong(IngestTallyListener.CREATED_KEY)).isEqualTo(2);
        assertThat(execution.getExecutionContext().getLong(IngestTallyListener.UPDATED_KEY)).isEqualTo(1);
        assertThat(execution.getExecutionContext().getLong(IngestTallyListener.NOT_LANDED_KEY)).isZero();
    }

    /**
     * A few refused rows are a data-quality problem for the next feed, not a reason to call the run
     * a failure — but they are still counted and reported rather than absorbed.
     */
    @Test
    void aFewRefusedRows_areReportedWithoutFailingTheRun() throws Exception {
        when(apiClient.ingestFacilities(anyList())).thenReturn(new ApiIngestSummary(9, 0, 3));

        JobExecution execution = runFeed(12);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        assertThat(execution.getExecutionContext().getLong(IngestTallyListener.NOT_LANDED_KEY)).isEqualTo(3);
    }

    @Test
    void pastTheThreshold_theRunFails() throws Exception {
        // The whole feed was posted and the write count will say so; the API kept two rows of it.
        when(apiClient.ingestFacilities(anyList())).thenReturn(new ApiIngestSummary(2, 0, 11));

        JobExecution execution = runFeed(13);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
        assertThat(execution.getExitStatus().getExitDescription()).contains("did not land");
        assertThat(execution.getExecutionContext().getLong(IngestTallyListener.NOT_LANDED_KEY)).isEqualTo(11);
    }

    /** The counts reach whoever triggered the job, not only the service log. */
    @Test
    void theTriggerResponseStatesWhatLanded() throws Exception {
        when(apiClient.ingestFacilities(anyList())).thenReturn(new ApiIngestSummary(1, 0, 0));

        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "facilities.csv"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.createdCount").value(1))
                .andExpect(jsonPath("$.updatedCount").value(0))
                .andExpect(jsonPath("$.rowsNotLanded").value(0));
    }

    private JobExecution runFeed(int rows) throws Exception {
        StringBuilder csv = new StringBuilder(
                "\"agent_bank\",\"name\",\"account_number\",\"loan_amount\",\"maturity_date\","
                + "\"bank_status\",\"bank_status_date\",\"ubs_participation\",\"collateral_date\"\n");
        for (int i = 1; i <= rows; i++) {
            csv.append("\"Bank of America\",\"Facility %d\",\"ACC%d\",\"1000\",\"2026-01-01\",\"Active\",\"2026-01-01\",\"1\",\"2026-01-01\"\n"
                    .formatted(i, i));
        }
        feedFile = Files.createTempFile("tally-feed", ".csv");
        Files.writeString(feedFile, csv.toString());
        return jobOperator.start(facilityIngestJob, new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters());
    }
}
