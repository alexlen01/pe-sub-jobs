package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.model.LpMasterRow;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import com.ubs.pesubjobs.processor.LpMasterRowProcessor;
import org.springframework.batch.core.configuration.annotation.StepScope;
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

@Configuration
public class LpMasterIngestJobConfig {

    @Bean
    public Job lpMasterIngestJob(JobRepository jobRepository,
                                 @Qualifier("lpMasterIngestStep") Step lpMasterIngestStep) {
        return new JobBuilder("lpMasterIngestJob", jobRepository)
                .start(lpMasterIngestStep)
                .build();
    }

    @Bean("lpMasterIngestStep")
    public Step lpMasterIngestStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("lpMasterReader") FlatFileItemReader<LpMasterRow> lpMasterReader,
                                   LpMasterRowProcessor lpMasterProcessor,
                                   @Qualifier("lpMasterWriter") ItemWriter<ProcessedLpMaster> lpMasterWriter) {
        return new StepBuilder("lpMasterIngestStep", jobRepository)
                .<LpMasterRow, ProcessedLpMaster>chunk(50)
                .transactionManager(txManager)
                .reader(lpMasterReader)
                .processor(lpMasterProcessor)
                .writer(lpMasterWriter)
                .faultTolerant()
                .skip(Exception.class)
                .skipLimit(10)
                .build();
    }

    @Bean("lpMasterReader")
    @StepScope
    public FlatFileItemReader<LpMasterRow> lpMasterReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<LpMasterRow>()
                .name("lpMasterReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .lineTokenizer(CsvLineTokenizers.lenientQuotedCsvTokenizer(
                        "investorName", "parent", "spv", "highQty", "investorType",
                        "instVsHnw", "regionLocation", "investmentGrade", "sp", "mdy",
                        "fitch", "aum", "nav", "pension", "pensionFunded",
                        "ubsClassification", "ubsDefaultAdvRate", "ubsDefaultConcLimit", "notes"))
                .fieldSetMapper(fs -> new LpMasterRow(
                        fs.readString("investorName"),
                        fs.readString("parent"),
                        fs.readString("spv"),
                        fs.readString("highQty"),
                        fs.readString("investorType"),
                        fs.readString("instVsHnw"),
                        fs.readString("regionLocation"),
                        fs.readString("investmentGrade"),
                        fs.readString("sp"),
                        fs.readString("mdy"),
                        fs.readString("fitch"),
                        fs.readString("aum"),
                        fs.readString("nav"),
                        fs.readString("pension"),
                        fs.readString("pensionFunded"),
                        fs.readString("ubsClassification"),
                        fs.readString("ubsDefaultAdvRate"),
                        fs.readString("ubsDefaultConcLimit"),
                        fs.readString("notes")
                ))
                .build();
    }

    @Bean
    public LpMasterRowProcessor lpMasterProcessor() {
        return new LpMasterRowProcessor();
    }

    /**
     * Posts each chunk to pe-sub-api's LP Master ingest endpoint, which upserts by investor
     * name. pe-sub-api owns the lp_master schema — this app issues no SQL against it.
     */
    @Bean("lpMasterWriter")
    public ItemWriter<ProcessedLpMaster> lpMasterWriter(PeSubApiClient apiClient) {
        return chunk -> apiClient.ingestLpMaster(List.copyOf(chunk.getItems()));
    }
}
