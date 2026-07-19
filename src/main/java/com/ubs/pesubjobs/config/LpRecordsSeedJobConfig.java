package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.Step;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.infrastructure.item.ItemWriter;
import org.springframework.batch.infrastructure.item.file.FlatFileItemReader;
import org.springframework.batch.infrastructure.item.file.builder.FlatFileItemReaderBuilder;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.FileSystemResource;
import org.springframework.transaction.PlatformTransactionManager;

import java.util.List;

/**
 * Seeds facility LP records from the lp_facility_seeds feed. Rows are posted verbatim to
 * pe-sub-api's seed endpoint, which resolves the facility and LP Master references by name,
 * merges the LP Master profile, normalizes the classifications, and inserts only pairs that
 * do not already exist (lp_records intentionally has no unique constraint on
 * facility+investor, so idempotency is application-level and server-side).
 */
@Configuration
public class LpRecordsSeedJobConfig {

    @Bean
    public Job lpRecordsSeedJob(JobRepository jobRepository,
                                 @Qualifier("lpRecordsSeedStep") Step lpRecordsSeedStep) {
        return new JobBuilder("lpRecordsSeedJob", jobRepository)
                .start(lpRecordsSeedStep)
                .build();
    }

    @Bean("lpRecordsSeedStep")
    public Step lpRecordsSeedStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("lpFacilitySeedReader") FlatFileItemReader<LpFacilitySeedRow> reader,
                                   @Qualifier("lpRecordsSeedWriter") ItemWriter<LpFacilitySeedRow> writer) {
        return new StepBuilder("lpRecordsSeedStep", jobRepository)
                .<LpFacilitySeedRow, LpFacilitySeedRow>chunk(50)
                .transactionManager(txManager)
                .reader(reader)
                .writer(writer)
                .faultTolerant().skip(Exception.class).skipLimit(10)
                .build();
    }

    @Bean("lpFacilitySeedReader")
    @org.springframework.batch.core.configuration.annotation.StepScope
    public FlatFileItemReader<LpFacilitySeedRow> lpFacilitySeedReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<LpFacilitySeedRow>()
                .name("lpFacilitySeedReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .lineTokenizer(CsvLineTokenizers.lenientQuotedCsvTokenizer(
                        "facilityName", "investorName", "capCommit", "uncalled",
                        "agentCls", "agentRate", "agentConc",
                        "parent", "spv", "highQty", "investorType", "instVsHnw",
                        "regionLocation", "investmentGrade", "ubsCls", "sp", "mdy", "fitch",
                        "aum", "nav", "pension", "pensionFunded", "pctCapCommit", "calledCap",
                        "pctUncalled", "pctCalled", "ubsConc", "ubsRate", "agentBb", "ubsBb",
                        "notes"))
                .fieldSetMapper(fs -> new LpFacilitySeedRow(
                        fs.readString("facilityName"),
                        fs.readString("investorName"),
                        fs.readString("capCommit"),
                        fs.readString("uncalled"),
                        fs.readString("agentCls"),
                        fs.readString("agentRate"),
                        fs.readString("agentConc"),
                        fs.readString("parent"),
                        fs.readString("spv"),
                        fs.readString("highQty"),
                        fs.readString("investorType"),
                        fs.readString("instVsHnw"),
                        fs.readString("regionLocation"),
                        fs.readString("investmentGrade"),
                        fs.readString("ubsCls"),
                        fs.readString("sp"),
                        fs.readString("mdy"),
                        fs.readString("fitch"),
                        fs.readString("aum"),
                        fs.readString("nav"),
                        fs.readString("pension"),
                        fs.readString("pensionFunded"),
                        fs.readString("pctCapCommit"),
                        fs.readString("calledCap"),
                        fs.readString("pctUncalled"),
                        fs.readString("pctCalled"),
                        fs.readString("ubsConc"),
                        fs.readString("ubsRate"),
                        fs.readString("agentBb"),
                        fs.readString("ubsBb"),
                        fs.readString("notes")))
                .build();
    }

    @Bean("lpRecordsSeedWriter")
    public ItemWriter<LpFacilitySeedRow> lpRecordsSeedWriter(PeSubApiClient apiClient) {
        return chunk -> apiClient.seedLpRecords(List.copyOf(chunk.getItems()));
    }
}
