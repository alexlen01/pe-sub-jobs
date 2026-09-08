package com.ubs.pesubjobs;

import com.ubs.pesubjobs.model.ProcessedUmbrella;
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
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.verify;

/**
 * The umbrella feed carries the group layer: one row per credit agreement, keyed by whatever
 * resolves it across re-runs. The job parses it and hands the typed rows to pe-sub-api's umbrella
 * ingest endpoint — the API owns the upsert and the naming rules; this app owns only the parsing.
 */
class UmbrellaIngestJobTest extends IntegrationTestBase {

    @Autowired JobOperator jobOperator;
    @Autowired @Qualifier("umbrellaIngestJob") Job umbrellaIngestJob;

    private Path feedFile;

    @AfterEach
    void cleanup() throws Exception {
        if (feedFile != null) Files.deleteIfExists(feedFile);
    }

    @Test
    void parsesTheGroupRowAndPostsItToTheApi() throws Exception {
        // The shape the report actually prints for an account umbrella: one row naming the credit
        // agreement, over member funds the report never names at all.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","loan_amount","cross_collateralized"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000",""
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        List<ProcessedUmbrella> rows = capturedUmbrellaRows();
        assertThat(rows).hasSize(1);
        ProcessedUmbrella row = rows.getFirst();
        assertThat(row.key()).isEqualTo("5VZ8873");
        assertThat(row.name()).isEqualTo("Carlyle Buyout Umbrella");
        assertThat(row.agentBank()).isEqualTo("Wells Fargo");
        assertThat(row.accountNumber()).isEqualTo("5VZ8873");
        // The whole agreement's line, stated once over every member fund.
        assertThat(row.facilitySize()).isEqualByComparingTo(new BigDecimal("1500000000"));
        // Blank is "not stated", never "the members hold separate collateral" — null is what stops
        // the API reading silence as an answer and turning an analyst's shared base off.
        assertThat(row.crossCollateralized()).isNull();
    }

    @Test
    void carriesTheSharedBaseATrancheSetStates() throws Exception {
        // A tranche set is the one group shape the feed can speak to collateral for: same borrower,
        // same collateral pool, split only by commitment. Its sleeves hold different account
        // numbers, so the key is the agreement and the account column stays blank.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","loan_amount","cross_collateralized"
                "ares direct lending","Ares Direct Lending","Bank of America","","100000000","true"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.key()).isEqualTo("ares direct lending");
        assertThat(row.crossCollateralized()).isTrue();
        // Nothing is invented for a column the report said nothing in: the sleeves share no
        // account, so the account column is silence rather than a number borrowed from elsewhere.
        assertThat(row.accountNumber()).isNull();
    }

    @Test
    void silentColumnsArriveAsNullRatherThanAsBlanks() throws Exception {
        // A group the report printed almost nothing about must reach the API as silence, so the
        // ingest leaves whatever an analyst recorded alone instead of blanking it.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","loan_amount","cross_collateralized"
                "5VT5929","Emberly Umbrella XII, LP","","5VT5929","",""
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.name()).isEqualTo("Emberly Umbrella XII, LP");
        assertThat(row.agentBank()).isNull();
        assertThat(row.facilitySize()).isNull();
        assertThat(row.crossCollateralized()).isNull();
    }

    @Test
    void everyGroupInTheFeedReachesTheApi() throws Exception {
        // A run loads the whole group layer: a group left behind here is a group the facilities
        // that name it would go on to create under a minted name, losing the printed one.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","loan_amount","cross_collateralized"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000",""
                "5VT5929","Emberly Umbrella XII, LP","BFBH","5VT5929","200000000",""
                "ares direct lending","Ares Direct Lending","Bank of America","","100000000","true"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);
        assertThat(capturedUmbrellaRows())
                .extracting((ProcessedUmbrella u) -> u.key())
                .containsExactly("5VZ8873", "5VT5929", "ares direct lending");
    }

    @Test
    void carriesTheAgreementReferenceWhenTheFeedPrintsIt() throws Exception {
        // What the agent actually signed: a reference for the agreement itself. It is what the API
        // groups on ahead of the account number, because an account is how a bank administers an
        // agreement and can be re-papered without the agreement changing.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","loan_amount","cross_collateralized","agreement_ref"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000","","CA-2021-4471"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.agreementRef()).isEqualTo("CA-2021-4471");
    }

    @Test
    void aFeedWrittenBeforeTheAgreementColumnsExisted_stillLoads() throws Exception {
        // The column is appended, never inserted, and the tokenizer is lenient — so last quarter's
        // six-column extract still parses and simply states no agreement reference. The API then
        // groups it by its key exactly as it did before the columns existed, which is what keeps the
        // legacy book whole through the rename.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","loan_amount","cross_collateralized"
                "5VT5929","Emberly Umbrella XII, LP","BFBH","5VT5929","200000000",""
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.key()).isEqualTo("5VT5929");
        assertThat(row.agreementRef()).isNull();
    }

    private List<ProcessedUmbrella> capturedUmbrellaRows() {
        ArgumentCaptor<List<ProcessedUmbrella>> captor = ArgumentCaptor.captor();
        verify(apiClient).ingestUmbrellas(captor.capture());
        return captor.getValue();
    }

    private JobExecution runFeed(String csv) throws Exception {
        feedFile = Files.createTempFile("umbrella-feed", ".csv");
        Files.writeString(feedFile, csv);
        JobParameters params = new JobParametersBuilder()
                .addString("filePath", feedFile.toString())
                .addLong("runId", System.nanoTime())
                .toJobParameters();
        return jobOperator.start(umbrellaIngestJob, params);
    }
}
