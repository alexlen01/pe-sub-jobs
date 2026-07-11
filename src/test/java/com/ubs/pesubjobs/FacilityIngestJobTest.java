package com.ubs.pesubjobs;

import com.ubs.pesubjobs.model.ProcessedFacility;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;

import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.verify;

/**
 * The facility feed carries UBS Participation (dollar amount) and Collateral As Of Date
 * alongside the core columns, in mixed date formats (M/d/yyyy and ISO). The job must parse
 * them and hand the typed rows to pe-sub-api's facility ingest endpoint — the API owns the
 * upsert; this app owns only the CSV parsing.
 */
class FacilityIngestJobTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("facilityIngestJob") Job facilityIngestJob;

    private Path feedFile;

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void parsesFeedAndPostsTypedRowsToApi() throws Exception {
        // Column order matches facilities.csv: agent_bank, name, account_number, loan_amount,
        // maturity_date, bank_status, bank_status_date, ubs_participation, collateral_date.
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        List<ProcessedFacility> rows = capturedFacilityRows();
        assertThat(rows).hasSize(1);
        ProcessedFacility row = rows.getFirst();
        assertThat(row.name()).isEqualTo("HIG LBO IV");
        assertThat(row.agentBank()).isEqualTo("Bank of America");
        assertThat(row.accountNumber()).isEqualTo("5VX1796");
        assertThat(row.loanAmount()).isEqualByComparingTo(new BigDecimal("75000000"));
        assertThat(row.maturityDate()).isEqualTo(LocalDate.of(2026, 10, 26));   // M/d/yyyy
        assertThat(row.bankStatusDate()).isEqualTo(LocalDate.of(2026, 5, 21));  // M/d/yyyy
        assertThat(row.ubsParticipation()).isEqualByComparingTo(new BigDecimal("9502500.00"));
        assertThat(row.collateralDate()).isEqualTo(LocalDate.of(2026, 6, 9));   // ISO
    }

    @Test
    void rowsWithoutNameOrAgentBank_areFilteredBeforeTheApiCall() throws Exception {
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09"
                "","Nameless Agent Facility","X","1","2026-01-01","Active","2026-01-01","1","2026-01-01"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        assertThat(capturedFacilityRows())
                .extracting((ProcessedFacility f) -> f.name())
                .containsExactly("HIG LBO IV");
    }

    private List<ProcessedFacility> capturedFacilityRows() {
        ArgumentCaptor<List<ProcessedFacility>> captor = ArgumentCaptor.captor();
        verify(apiClient).ingestFacilities(captor.capture());
        return captor.getValue();
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
