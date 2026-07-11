package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.model.FacilityRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import com.ubs.pesubjobs.processor.FacilityRowProcessor;
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
public class FacilityIngestJobConfig {

    @Bean
    public Job facilityIngestJob(JobRepository jobRepository,
                                 @Qualifier("facilityIngestStep") Step facilityIngestStep) {
        return new JobBuilder("facilityIngestJob", jobRepository)
                .start(facilityIngestStep)
                .build();
    }

    @Bean("facilityIngestStep")
    public Step facilityIngestStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("facilityReader") FlatFileItemReader<FacilityRow> facilityReader,
                                   FacilityRowProcessor facilityProcessor,
                                   @Qualifier("facilityWriter") ItemWriter<ProcessedFacility> facilityWriter) {
        return new StepBuilder("facilityIngestStep", jobRepository)
                .<FacilityRow, ProcessedFacility>chunk(50)
                .transactionManager(txManager)
                .reader(facilityReader)
                .processor(facilityProcessor)
                .writer(facilityWriter)
                .faultTolerant()
                .skip(Exception.class)
                .skipLimit(10)
                .build();
    }

    @Bean("facilityReader")
    @StepScope
    public FlatFileItemReader<FacilityRow> facilityReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<FacilityRow>()
                .name("facilityReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .lineTokenizer(CsvLineTokenizers.lenientQuotedCsvTokenizer(
                        "agentBank", "name", "accountNumber", "loanAmount",
                        "maturityDate", "bankStatus", "bankStatusDate",
                        "ubsParticipation", "collateralDate"))
                .fieldSetMapper(fs -> new FacilityRow(
                        fs.readString("agentBank"),
                        fs.readString("name"),
                        fs.readString("accountNumber"),
                        fs.readString("loanAmount"),
                        fs.readString("maturityDate"),
                        fs.readString("bankStatus"),
                        fs.readString("bankStatusDate"),
                        fs.readString("ubsParticipation"),
                        fs.readString("collateralDate")
                ))
                .build();
    }

    @Bean
    public FacilityRowProcessor facilityProcessor() {
        return new FacilityRowProcessor();
    }

    /**
     * Posts each chunk to pe-sub-api's facility ingest endpoint, which upserts by name.
     * pe-sub-api owns the facilities schema — this app issues no SQL against it.
     */
    @Bean("facilityWriter")
    public ItemWriter<ProcessedFacility> facilityWriter(PeSubApiClient apiClient) {
        return chunk -> apiClient.ingestFacilities(List.copyOf(chunk.getItems()));
    }
}
