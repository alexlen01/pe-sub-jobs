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
    void rowsWithoutNameOrAgentBank_stillReachTheApi() throws Exception {
        // Neither a missing agent bank nor a missing name is filtered out here; both are filled
        // in by the API, which owns the constraints.
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09"
                "","Nameless Agent Facility","X","1","2026-01-01","Active","2026-01-01","1","2026-01-01"
                "Bank of America","","5VX9999","1","2026-01-01","Active","2026-01-01","1","2026-01-01"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        List<ProcessedFacility> rows = capturedFacilityRows();
        assertThat(rows)
                .extracting((ProcessedFacility f) -> f.name())
                .containsExactly("HIG LBO IV", "Nameless Agent Facility", null);
        assertThat(rows.get(1).agentBank()).isNull();
    }

    @Test
    void carriesTheUmbrellaTheFeedStates() throws Exception {
        // Two funds reported on one account number is an umbrella subscription facility. The feed
        // states the grouping outright; this job never works it out, because it reads the file in
        // chunks and so never has enough of it in view to notice two rows sharing an account.
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date","umbrella_name"
                "Bank of America","Carlyle Buyout Umbrella","5VZ8873","100000000","2028-01-12","Active","2026-05-13","","2026-04-28","Umbrella 5VZ8873"
                "Bank of America","Oaktree Opportunities Fund Xb","5VZ8873","112000000","2027-07-21","Active","2025-07-23","","2025-11-24","Umbrella 5VZ8873"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09",""
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        List<ProcessedFacility> rows = capturedFacilityRows();
        assertThat(rows)
                .extracting((ProcessedFacility f) -> f.umbrellaName())
                // A fund that borrows alone states no umbrella, and blank arrives as null rather
                // than as an empty group name.
                .containsExactly("Umbrella 5VZ8873", "Umbrella 5VZ8873", null);
    }

    @Test
    void carriesTheGroupKeyAndSharedBaseTheFeedStates() throws Exception {
        // The sleeves of a multi-tranche facility hold DIFFERENT account numbers — each tranche is
        // certified separately — so the account number cannot resolve the group and the feed states
        // a key of its own. It also states that the sleeves stand on one borrowing base, which for
        // a tranche set it can know: same borrower, same collateral, split only by commitment.
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date","umbrella_name","umbrella_key","cross_collateralized"
                "Bank of America","Ares Direct Lending (Committed)","5VZ1001","60000000","2028-01-12","Active","2026-05-13","","2026-04-28","Ares Direct Lending","Ares Direct Lending","true"
                "Bank of America","Ares Direct Lending (Uncommitted)","5VZ1002","40000000","2028-01-12","Active","2026-05-13","","2026-04-28","Ares Direct Lending","Ares Direct Lending","true"
                "Bank of America","Carlyle Buyout Umbrella","5VZ8873","100000000","2028-01-12","Active","2026-05-13","","2026-04-28","Umbrella 5VZ8873","5VZ8873",""
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        List<ProcessedFacility> rows = capturedFacilityRows();
        assertThat(rows)
                .extracting((ProcessedFacility f) -> f.umbrellaKey())
                .containsExactly("Ares Direct Lending", "Ares Direct Lending", "5VZ8873");
        // Blank is "not stated", not "no shared base": an account umbrella's terms are a matter for
        // an analyst, and null is what stops the ingest reading silence as an answer.
        assertThat(rows)
                .extracting((ProcessedFacility f) -> f.umbrellaCrossCollateralized())
                .containsExactly(true, true, null);
    }

    @Test
    void feedWrittenBeforeTheUmbrellaColumnExisted_stillLoads() throws Exception {
        // The umbrella columns are last and the tokenizer is lenient, so a nine-column feed loads
        // and simply states no umbrella — an older feed file is never a failed run.
        JobExecution execution = runFeed("""
                "agent_bank","name","account_number","loan_amount","maturity_date","bank_status","bank_status_date","ubs_participation","collateral_date"
                "Bank of America","HIG LBO IV","5VX1796","75000000","10/26/2026","Active","5/21/2026","9502500.00","2026-06-09"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        List<ProcessedFacility> rows = capturedFacilityRows();
        assertThat(rows).hasSize(1);
        assertThat(rows.getFirst().name()).isEqualTo("HIG LBO IV");
        assertThat(rows.getFirst().umbrellaName()).isNull();
        assertThat(rows.getFirst().umbrellaKey()).isNull();
        assertThat(rows.getFirst().umbrellaCrossCollateralized()).isNull();
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
