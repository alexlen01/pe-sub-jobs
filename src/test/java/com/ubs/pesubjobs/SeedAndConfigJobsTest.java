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
        // Row 1 is a full 32-column row; row 2 is a legacy 7-column row, which the non-strict
        // tokenizer pads with blanks (the API then falls back to LP Master for those fields).
        // high_quality is gone from the feed (the LP DB Export dropped it) and the two
        // excess-concentration columns arrived with the 2026-08-18 format.
        JobExecution execution = run(lpRecordsSeedJob, """
                "facility_name","investor_name","capital_commitment","uncalled_capital","agent_lp_category","agent_advance_rate","agent_concentration_limit","parent","spv","investor_type","institutional_or_hnw","region_location","investment_grade","ubs_lp_category","sp_rating","moodys_rating","fitch_rating","aum","nav","pension_assets","funding_ratio","pct_of_fund_commitments","called_capital","pct_of_fund_uncalled","pct_lp_called","ubs_concentration_limit","ubs_advance_rate","agent_excess_concentration","ubs_excess_concentration","agent_borrowing_base","ubs_borrowing_base","notes"
                "Carlyle Buyout Umbrella","Acme Pension Fund","$250M","$75M","Rated Included","90%","5%","Acme Holdings","FALSE","Pension Fund","Institutional","United States","TRUE","Rated Investor","AA","Aa2","AA","$10B","$8B","$9B","105%","3%","$175M","2%","70%","4%","90%","$2.5M","$1.5M","$67.5M","$67.5M","seed note"
                "KKR Ascendant","Beta Capital LLC","$100M","$40M","Designated","50%","15%"
                """);

        assertThat(execution.getStatus()).isEqualTo(BatchStatus.COMPLETED);

        ArgumentCaptor<List<LpFacilitySeedRow>> captor = ArgumentCaptor.captor();
        verify(apiClient).seedLpRecords(captor.capture());
        List<LpFacilitySeedRow> rows = captor.getValue();
        assertThat(rows).hasSize(2);
        LpFacilitySeedRow full = rows.getFirst();
        assertThat(full.facilityName()).isEqualTo("Carlyle Buyout Umbrella");
        assertThat(full.investorName()).isEqualTo("Acme Pension Fund");
        assertThat(full.capitalCommitment()).isEqualTo("$250M");
        assertThat(full.uncalledCapital()).isEqualTo("$75M");
        assertThat(full.agentLpCategory()).isEqualTo("Rated Included");
        assertThat(full.parent()).isEqualTo("Acme Holdings");
        assertThat(full.spv()).isEqualTo("FALSE");
        assertThat(full.investorType()).isEqualTo("Pension Fund");
        assertThat(full.institutionalOrHnw()).isEqualTo("Institutional");
        assertThat(full.regionLocation()).isEqualTo("United States");
        assertThat(full.investmentGrade()).isEqualTo("TRUE");
        assertThat(full.ubsLpCategory()).isEqualTo("Rated Investor");
        assertThat(full.spRating()).isEqualTo("AA");
        assertThat(full.moodysRating()).isEqualTo("Aa2");
        assertThat(full.fitchRating()).isEqualTo("AA");
        assertThat(full.aum()).isEqualTo("$10B");
        assertThat(full.nav()).isEqualTo("$8B");
        assertThat(full.pensionAssets()).isEqualTo("$9B");
        assertThat(full.fundingRatio()).isEqualTo("105%");
        assertThat(full.pctOfFundCommitments()).isEqualTo("3%");
        assertThat(full.calledCapital()).isEqualTo("$175M");
        assertThat(full.pctOfFundUncalled()).isEqualTo("2%");
        assertThat(full.pctLpCalled()).isEqualTo("70%");
        assertThat(full.ubsConcentrationLimit()).isEqualTo("4%");
        assertThat(full.ubsAdvanceRate()).isEqualTo("90%");
        assertThat(full.agentExcessConcentration()).isEqualTo("$2.5M");
        assertThat(full.ubsExcessConcentration()).isEqualTo("$1.5M");
        assertThat(full.agentBorrowingBase()).isEqualTo("$67.5M");
        assertThat(full.ubsBorrowingBase()).isEqualTo("$67.5M");
        assertThat(full.notes()).isEqualTo("seed note");

        LpFacilitySeedRow legacy = rows.getLast();
        assertThat(legacy.facilityName()).isEqualTo("KKR Ascendant");
        assertThat(legacy.agentAdvanceRate()).isEqualTo("50%");
        assertThat(legacy.parent()).isEmpty();     // 7-column row -> new columns blank-padded
        assertThat(legacy.ubsLpCategory()).isEmpty();
        assertThat(legacy.notes()).isEmpty();
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
