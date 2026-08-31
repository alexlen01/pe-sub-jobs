package com.ubs.pesubjobs;

import com.ubs.pesubjobs.model.ProcessedLpMaster;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.mockito.ArgumentMatchers;
import org.mockito.InOrder;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.inOrder;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;

/**
 * LP Master is replaced wholesale, and the clear that precedes the load is a remote call the batch
 * transaction manager cannot roll back. So what these tests pin is when the clear happens: only
 * after the entire replacement has been read and processed, and never when the feed cannot supply
 * one. A cleared table with nothing to put back is the loss this job's ordering exists to prevent.
 */
class LpMasterIngestJobTest extends IntegrationTestBase {

    private static final String HEADER =
            "\"investor_name\",\"parent\",\"spv\",\"investor_type\",\"institutional_or_hnw\","
            + "\"region_location\",\"investment_grade\",\"sp_rating\",\"moodys_rating\",\"fitch_rating\","
            + "\"aum\",\"nav\",\"pension_assets\",\"funding_ratio\",\"ubs_lp_category\","
            + "\"ubs_default_advance_rate\",\"ubs_default_concentration_limit\",\"notes\"\n";

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("lpMasterIngestJob") Job lpMasterIngestJob;

    private Path feedFile;

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void clearsOnlyAfterTheWholeReplacementHasBeenStaged() throws Exception {
        JobExecution execution = runFeed(HEADER
                + "\"Acme Pension Fund\",\"Acme Holdings\",\"FALSE\",\"Pension Fund\",\"Institutional\","
                + "\"United States\",\"TRUE\",\"AA\",\"Aa2\",\"AA\",\"$10B\",\"$8B\",\"$9B\",\"105%\","
                + "\"Rated Investor\",\"90%\",\"5%\",\"master note\"\n"
                + "\"Beta Capital LLC\",\"\",\"TRUE\",\"Family Office\",\"HNW\",\"United States\","
                + "\"FALSE\",\"\",\"\",\"\",\"$1B\",\"$900M\",\"\",\"\",\"Designated\",\"50%\",\"15%\",\"\"\n");

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        InOrder order = inOrder(apiClient);
        order.verify(apiClient).clearLpMaster();
        order.verify(apiClient).ingestLpMaster(anyRows());

        ArgumentCaptor<List<ProcessedLpMaster>> captor = ArgumentCaptor.captor();
        verify(apiClient).ingestLpMaster(captor.capture());
        List<ProcessedLpMaster> rows = captor.getValue();
        assertThat(rows).hasSize(2);
        assertThat(rows.getFirst().investorName()).isEqualTo("Acme Pension Fund");
        assertThat(rows.getFirst().parent()).isEqualTo("Acme Holdings");
        assertThat(rows.getFirst().ubsLpCategory()).isEqualTo("Rated Investor");
        assertThat(rows.getLast().investorName()).isEqualTo("Beta Capital LLC");
        assertThat(rows.getLast().parent()).isNull();
        assertThat(rows.getLast().spv()).isTrue();
    }

    @Test
    void unreadableFeed_failsWithoutClearingTheExistingTable() throws Exception {
        JobExecution execution = jobOperator.start(lpMasterIngestJob, new JobParametersBuilder()
                .addString("filePath", "src/test/resources/feeds/no-such-lp-master.csv")
                .addLong("runId", System.nanoTime())
                .toJobParameters());

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
        // The whole point of the ordering: a feed that cannot be read costs nothing.
        verify(apiClient, never()).clearLpMaster();
        verify(apiClient, never()).ingestLpMaster(anyRows());
    }

    @Test
    void feedWithNoRows_failsWithoutClearingTheExistingTable() throws Exception {
        // A header-only feed would otherwise replace the entire table with nothing, which is the
        // same loss as a failed read.
        JobExecution execution = runFeed(HEADER);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.FAILED);
        verify(apiClient, never()).clearLpMaster();
        verify(apiClient, never()).ingestLpMaster(anyRows());
    }

    /** Any chunk of rows — these assertions are about whether the call happened, not its contents. */
    private static List<ProcessedLpMaster> anyRows() {
        return ArgumentMatchers.any();
    }

    private JobExecution runFeed(String csv) throws Exception {
        feedFile = Files.createTempFile("lp-master-feed", ".csv");
        Files.writeString(feedFile, csv);
        return jobOperator.start(lpMasterIngestJob, new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters());
    }
}
