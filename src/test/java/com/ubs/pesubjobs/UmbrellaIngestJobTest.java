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
import java.time.LocalDate;
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
                "key","name","agent_bank","account_number","facility_size","cross_collateralized"
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
                "key","name","agent_bank","account_number","facility_size","cross_collateralized"
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
                "key","name","agent_bank","account_number","facility_size","cross_collateralized"
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
                "key","name","agent_bank","account_number","facility_size","cross_collateralized"
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
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000","","CA-2021-4471"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.creditAgreementRef()).isEqualTo("CA-2021-4471");
    }

    @Test
    void aFeedWrittenBeforeTheAgreementColumnsExisted_stillLoads() throws Exception {
        // The column is appended, never inserted, and the tokenizer is lenient — so last quarter's
        // six-column extract still parses and simply states no agreement reference. The API then
        // groups it by its key exactly as it did before the columns existed, which is what keeps the
        // legacy book whole through the rename.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized"
                "5VT5929","Emberly Umbrella XII, LP","BFBH","5VT5929","200000000",""
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.key()).isEqualTo("5VT5929");
        assertThat(row.creditAgreementRef()).isNull();
    }

    @Test
    void carriesTheAgreementsOwnTermsAndStanding() throws Exception {
        // The terms a group hands down to every member fund while it is Active. They are parsed here
        // so an agreement whose maturity and collateral date are printed in the file does not land on
        // the platform empty and governing nothing until somebody types them back in.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref","maturity_date","collateral_date","facility_status"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000","","CA-2021-4471","2029-03-31","2026-06-25","Active"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.maturityDate()).isEqualTo(LocalDate.of(2029, 3, 31));
        assertThat(row.collateralDate()).isEqualTo(LocalDate.of(2026, 6, 25));
        // The one switch that makes the group govern its members at all.
        assertThat(row.facilityStatus()).isEqualTo("Active");
    }

    @Test
    void carriesUbsSliceOfTheAgreementsLineAlongsideTheWholeOfIt() throws Exception {
        // Two amounts, two facts: the whole syndicated line every lender in the group holds, and the
        // slice of it UBS is exposed for. Both are stated once over the agreement, so both are parsed
        // here — deriving one from the other would put a guess where the agent printed a figure.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref","maturity_date","collateral_date","facility_status","ubs_participation"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000","","CA-2021-4471","2029-03-31","2026-06-25","Active","240000000"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.facilitySize()).isEqualByComparingTo(new BigDecimal("1500000000"));
        // A part of the whole above, never the whole restated under another name.
        assertThat(row.ubsParticipation()).isEqualByComparingTo(new BigDecimal("240000000"));
    }

    @Test
    void aFeedWrittenBeforeTheParticipationColumnExisted_statesNoParticipation() throws Exception {
        // The column is appended after the status, so a ten-column extract is short one cell rather
        // than misaligned. The group loads with its size intact and says nothing about UBS's slice,
        // which is what stops a re-run of an older extract blanking a participation on record.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref","maturity_date","collateral_date","facility_status"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","1500000000","","CA-2021-4471","2029-03-31","2026-06-25","Active"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.facilitySize()).isEqualByComparingTo(new BigDecimal("1500000000"));
        assertThat(row.ubsParticipation()).isNull();
    }

    @Test
    void readsTheSlashedDateShapeTheFacilityFeedUses() throws Exception {
        // Both feeds state the same agreement's dates, so both read the same shapes. A group feed
        // that rejected M/d/yyyy would stop a run over a date its sibling file accepts.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref","maturity_date","collateral_date","facility_status"
                "5VZ8873","Carlyle Buyout Umbrella","Wells Fargo","5VZ8873","","","","3/31/2029","6/25/2026","Active"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.maturityDate()).isEqualTo(LocalDate.of(2029, 3, 31));
        assertThat(row.collateralDate()).isEqualTo(LocalDate.of(2026, 6, 25));
    }

    @Test
    void aGroupStatingNoTermsReachesTheApiAsSilence() throws Exception {
        // The columns are present and empty — an inferred group whose members disagreed on their
        // terms. Silence must stay silence all the way to the API, which is what stops a re-run
        // blanking a maturity an analyst recorded.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref","maturity_date","collateral_date","facility_status"
                "5VZ9100","Umbrella 5VZ9100","Ashford Bank","5VZ9100","","","","","","Active"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.maturityDate()).isNull();
        assertThat(row.collateralDate()).isNull();
        assertThat(row.facilityStatus()).isEqualTo("Active");
    }

    @Test
    void aFeedWrittenBeforeTheTermsColumnsExisted_statesNoTerms() throws Exception {
        // Seven columns, from before the agreement's terms were carried. Appended-never-inserted
        // means the row is short rather than misaligned: the lenient tokenizer pads it, the group
        // loads exactly as it did before, and it simply states nothing about its own terms.
        JobExecution execution = runFeed("""
                "key","name","agent_bank","account_number","facility_size","cross_collateralized","credit_agreement_ref"
                "5VT5929","Emberly Umbrella XII, LP","BFBH","5VT5929","200000000","","CA-2019-8812"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ProcessedUmbrella row = capturedUmbrellaRows().getFirst();
        assertThat(row.creditAgreementRef()).isEqualTo("CA-2019-8812");
        assertThat(row.facilitySize()).isEqualByComparingTo(new BigDecimal("200000000"));
        assertThat(row.maturityDate()).isNull();
        assertThat(row.collateralDate()).isNull();
        // No standing invented for a file that states none — the API leaves the group's alone.
        assertThat(row.facilityStatus()).isNull();
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
