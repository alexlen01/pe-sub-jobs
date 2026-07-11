package com.ubs.pesubjobs;

import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
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
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.verify;

/**
 * The LP record seed job passes the raw feed rows to pe-sub-api verbatim — name resolution,
 * LP Master merge and classification normalization are the API's job now. The cls-conc job
 * parses/normalizes locally and merges one map per chunk.
 */
class SeedAndConfigJobsTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("lpRecordsSeedJob") Job lpRecordsSeedJob;
    @Autowired @Qualifier("clsConcLimitIngestJob") Job clsConcLimitIngestJob;

    private Path feedFile;

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void lpRecordsSeed_postsRawFeedRowsVerbatim() throws Exception {
        JobExecution execution = run(lpRecordsSeedJob, """
                "facility_name","investor_name","cap_commit","uncalled","agent_cls","agent_rate","agent_conc"
                "Carlyle Buyout Umbrella","Acme Pension Fund","$250M","$75M","Rated Included","90%","5%"
                "KKR Ascendant","Beta Capital LLC","$100M","$40M","Designated","50%","15%"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ArgumentCaptor<List<LpFacilitySeedRow>> captor = ArgumentCaptor.captor();
        verify(apiClient).seedLpRecords(captor.capture());
        List<LpFacilitySeedRow> rows = captor.getValue();
        assertThat(rows).hasSize(2);
        assertThat(rows.getFirst().facilityName()).isEqualTo("Carlyle Buyout Umbrella");
        assertThat(rows.getFirst().investorName()).isEqualTo("Acme Pension Fund");
        assertThat(rows.getFirst().capCommit()).isEqualTo("$250M");
        assertThat(rows.getFirst().uncalled()).isEqualTo("$75M");
        assertThat(rows.getFirst().agentCls()).isEqualTo("Rated Included");
        assertThat(rows.getLast().facilityName()).isEqualTo("KKR Ascendant");
        assertThat(rows.getLast().agentRate()).isEqualTo("50%");
    }

    @Test
    void clsConcLimits_mergesParsedMap_droppingInvalidRows() throws Exception {
        JobExecution execution = run(clsConcLimitIngestJob, """
                "classification","limit_pct"
                "Rated Investor","15%"
                "Unrated NAV > $1Bn","7.5"
                "Bad Row","not-a-number"
                "","5"
                "Excluded","250"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        // Blank classification, unparseable percent, and out-of-range percent (250) are dropped.
        ArgumentCaptor<Map<String, Double>> captor = ArgumentCaptor.captor();
        verify(apiClient).mergeClsConcLimitDefaults(captor.capture());
        assertThat(captor.getValue()).containsExactly(
                Map.entry("Rated Investor", 15.0),
                Map.entry("Unrated NAV > $1Bn", 7.5));
    }

    private JobExecution run(Job job, String csv) throws Exception {
        feedFile = Files.createTempFile("jobs-feed", ".csv");
        Files.writeString(feedFile, csv);
        return jobOperator.start(job, new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters());
    }
}
