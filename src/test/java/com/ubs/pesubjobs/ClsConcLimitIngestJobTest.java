package com.ubs.pesubjobs;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.jdbc.core.JdbcTemplate;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The cls-conc-limits feed must merge rows into the single {@code cls_conc_limit_defaults}
 * config row: fed classes are overwritten, unfed classes survive, and malformed rows are
 * skipped. The config table is owned by pe-sub-api's migrations, so this test provisions
 * a matching table on the embedded database first.
 */
class ClsConcLimitIngestJobTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("clsConcLimitIngestJob") Job clsConcLimitIngestJob;
    @Autowired JdbcTemplate jdbc;

    private Path feedFile;

    @BeforeEach
    void provisionConfigTable() {
        // Mirrors pe-sub-api V1_1__schema.sql. TEST ONLY
        jdbc.execute("""
                CREATE TABLE IF NOT EXISTS config (
                    key        VARCHAR(100) PRIMARY KEY,
                    value      JSONB        NOT NULL,
                    updated_at TIMESTAMP    NOT NULL DEFAULT NOW()
                )
                """);
        jdbc.update("DELETE FROM config WHERE key = 'cls_conc_limit_defaults'");
    }

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void feedCreatesConfigRowAndSkipsMalformedRows() throws Exception {
        JobExecution execution = runFeed("""
                classification,limit_pct
                Rated Investor,5
                Other Institutional,10%
                ,3
                Corp Pension > $5Bn Assets,not-a-number
                Excluded,250
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        // Blank classification, unparseable percent, and out-of-range percent are dropped.
        assertThat(readPct("Rated Investor")).isEqualTo(5.0);
        assertThat(readPct("Other Institutional")).isEqualTo(10.0);
        assertThat(keyCount()).isEqualTo(2);
    }

    @Test
    void feedMergesIntoExistingMapWithoutTouchingUnfedClasses() throws Exception {
        jdbc.update("""
                INSERT INTO config (key, value)
                VALUES ('cls_conc_limit_defaults', '{"Rated Investor": 7.5, "Excluded": 7.5}'::jsonb)
                """);

        JobExecution execution = runFeed("""
                classification,limit_pct
                Rated Investor,5
                FoF & Other > $10Bn AUM,6.5
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        // Fed rows are overwritten, while unfed classifications remain untouched.
        assertThat(readPct("Rated Investor")).isEqualTo(5.0);
        assertThat(readPct("Excluded")).isEqualTo(7.5);
        assertThat(readPct("FoF & Other > $10Bn AUM")).isEqualTo(6.5);
        assertThat(keyCount()).isEqualTo(3);
    }

    private JobExecution runFeed(String csv) throws Exception {
        feedFile = Files.createTempFile("cls-conc-feed", ".csv");
        Files.writeString(feedFile, csv);
        JobParameters params = new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters();
        return jobOperator.start(clsConcLimitIngestJob, params);
    }

    private Double readPct(String cls) {
        return jdbc.queryForObject(
                "SELECT (value->>?)::float8 FROM config WHERE key = 'cls_conc_limit_defaults'",
                Double.class, cls);
    }

    private Integer keyCount() {
        return jdbc.queryForObject("""
                SELECT count(*)::int FROM config, jsonb_object_keys(value)
                WHERE key = 'cls_conc_limit_defaults'
                """, Integer.class);
    }
}
