package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.model.ProcessedUmbrella;
import com.ubs.pesubjobs.model.UmbrellaRow;
import com.ubs.pesubjobs.processor.UmbrellaRowProcessor;
import org.springframework.batch.core.configuration.annotation.StepScope;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.Step;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.infrastructure.item.ItemWriter;
import org.springframework.batch.infrastructure.item.file.FlatFileItemReader;
import org.springframework.batch.infrastructure.item.file.builder.FlatFileItemReaderBuilder;
import com.ubs.pesubjobs.storage.FeedFileStore;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.transaction.PlatformTransactionManager;

import java.util.List;

/**
 * The group layer, loaded ahead of the facilities that belong to it: a fund names its group on its
 * own row, so the group has to exist under the name the agent printed before its members arrive, or
 * the first member through creates it under a name minted from an account number instead.
 */
@Configuration
public class UmbrellaIngestJobConfig {

    @Bean
    public Job umbrellaIngestJob(JobRepository jobRepository,
                                 @Qualifier("umbrellaIngestStep") Step umbrellaIngestStep) {
        return new JobBuilder("umbrellaIngestJob", jobRepository)
                .start(umbrellaIngestStep)
                .build();
    }

    @Bean("umbrellaIngestStep")
    public Step umbrellaIngestStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("umbrellaReader") FlatFileItemReader<UmbrellaRow> umbrellaReader,
                                   UmbrellaRowProcessor umbrellaProcessor,
                                   @Qualifier("umbrellaWriter") ItemWriter<ProcessedUmbrella> umbrellaWriter,
                                   IngestTallyListener tallyListener) {
        return new StepBuilder("umbrellaIngestStep", jobRepository)
                .<UmbrellaRow, ProcessedUmbrella>chunk(50)
                .transactionManager(txManager)
                .reader(umbrellaReader)
                .processor(umbrellaProcessor)
                .writer(umbrellaWriter)
                // No skip policy: a group that cannot be written fails the run rather than
                // vanishing, which would leave its member funds to be grouped by account number
                // under a minted name and the agreement's own name lost.
                .listener(tallyListener)
                .build();
    }

    @Bean("umbrellaReader")
    @StepScope
    public FlatFileItemReader<UmbrellaRow> umbrellaReader(
            @Value("#{jobParameters['filePath']}") String filePath,
            FeedFileStore feedFileStore) {
        return new FlatFileItemReaderBuilder<UmbrellaRow>()
                .name("umbrellaReader")
                .resource(feedFileStore.open(filePath))
                .linesToSkip(1)
                // In the order the extract writes them. The agreement's terms are APPENDED, so a
                // file written before they existed is short a few cells rather than misaligned; the
                // lenient tokenizer pads it and the group loads stating nothing about them. UBS's
                // participation is the newest of those appendices, which is why it trails the status
                // rather than sitting beside the facility size it is a slice of.
                .lineTokenizer(CsvLineTokenizers.lenientQuotedCsvTokenizer(
                        "key", "name", "agentBank", "accountNumber",
                        "facilitySize", "crossCollateralized", "creditAgreementRef",
                        "maturityDate", "collateralDate", "facilityStatus", "ubsParticipation"))
                .fieldSetMapper(fs -> new UmbrellaRow(
                        fs.readString("key"),
                        fs.readString("name"),
                        fs.readString("agentBank"),
                        fs.readString("accountNumber"),
                        fs.readString("facilitySize"),
                        fs.readString("crossCollateralized"),
                        fs.readString("creditAgreementRef"),
                        fs.readString("maturityDate"),
                        fs.readString("collateralDate"),
                        fs.readString("facilityStatus"),
                        fs.readString("ubsParticipation")
                ))
                .build();
    }

    @Bean
    public UmbrellaRowProcessor umbrellaProcessor() {
        return new UmbrellaRowProcessor();
    }

    /**
     * Posts each chunk to pe-sub-api's umbrella ingest endpoint, which upserts by group key.
     * pe-sub-api owns the umbrellas schema — this app issues no SQL against it.
     */
    @Bean("umbrellaWriter")
    public ItemWriter<ProcessedUmbrella> umbrellaWriter(PeSubApiClient apiClient, IngestTally tally) {
        return chunk -> tally.record(apiClient.ingestUmbrellas(List.copyOf(chunk.getItems())));
    }
}
