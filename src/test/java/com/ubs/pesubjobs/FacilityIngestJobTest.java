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

import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The facility feed carries UBS Participation (dollar amount) and Collateral As Of Date
 * alongside the core columns. Both must land on the {@code facilities} row so the UI edit
 * form/Dashboard render a real participation rate and collateral date rather than the
 * "defaults to 100%" / empty placeholders. The facilities table is owned by pe-sub-api's
 * migrations, so this test provisions a matching table on the embedded database first.
 */
class FacilityIngestJobTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("facilityIngestJob") Job facilityIngestJob;
    @Autowired JdbcTemplate jdbc;

    private Path feedFile;

    @BeforeEach
    void provisionFacilitiesTable() {
        // Mirrors the columns the writer touches in pe-sub-api V1_1__schema.sql. TEST ONLY
        jdbc.execute("""
                CREATE TABLE IF NOT EXISTS facilities (
                    id                SERIAL PRIMARY KEY,
                    name              VARCHAR(255)   NOT NULL UNIQUE,
                    agent_bank        VARCHAR(255)   NOT NULL,
                    status            VARCHAR(50)    NOT NULL DEFAULT 'Not Started',
                    conc_limit_m      NUMERIC(10, 2) NOT NULL DEFAULT 25,
                    account_number    VARCHAR(20),
                    loan_amount       NUMERIC(15, 2),
                    maturity_date     DATE,
                    bank_status       VARCHAR(50),
                    bank_status_date  DATE,
                    ubs_participation NUMERIC(15, 2),
                    collateral_date   DATE,
                    created_at        TIMESTAMP      NOT NULL DEFAULT NOW(),
                    updated_at        TIMESTAMP      NOT NULL DEFAULT NOW()
                )
                """);
        jdbc.update("DELETE FROM facilities");
    }

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void ingestsUbsParticipationAndCollateralDate() throws Exception {
        // Column order matches facilities.csv: agent_bank, name, account_number, loan_amount,
        // maturity_date, bank_status, bank_status_date, ubs_participation, collateral_date.
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        assertThat(jdbc.queryForObject(
                "SELECT ubs_participation FROM facilities WHERE name = 'HIG LBO IV'", BigDecimal.class))
                .isEqualByComparingTo("9502500.00");
        assertThat(jdbc.queryForObject(
                "SELECT collateral_date FROM facilities WHERE name = 'HIG LBO IV'", LocalDate.class))
                .isEqualTo(LocalDate.of(2026, 6, 9));
    }

    @Test
    void reingestOverwritesUbsParticipationAndCollateralDate() throws Exception {
        // A facility created before the feed carried these columns (both NULL) must be
        // back-filled on the next run via ON CONFLICT (name) DO UPDATE.
        jdbc.update("""
                INSERT INTO facilities (name, agent_bank, loan_amount)
                VALUES ('HIG LBO IV', 'Bank of America', 75000000)
                """);

        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        assertThat(jdbc.queryForObject(
                "SELECT ubs_participation FROM facilities WHERE name = 'HIG LBO IV'", BigDecimal.class))
                .isEqualByComparingTo("9502500.00");
        assertThat(jdbc.queryForObject(
                "SELECT collateral_date FROM facilities WHERE name = 'HIG LBO IV'", LocalDate.class))
                .isEqualTo(LocalDate.of(2026, 6, 9));
    }

    private JobExecution runFeed(String csv) throws Exception {
        feedFile = Files.createTempFile("facility-feed", ".csv");
        Files.writeString(feedFile, csv);
        JobParameters params = new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters();
        return jobOperator.start(facilityIngestJob, params);
    }
}
