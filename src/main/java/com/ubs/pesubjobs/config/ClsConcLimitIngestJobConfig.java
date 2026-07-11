package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.model.ClsConcLimitRow;
import com.ubs.pesubjobs.model.ProcessedClsConcLimit;
import com.ubs.pesubjobs.processor.ClsConcLimitRowProcessor;
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

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Classification concentration-limit defaults feed. CSV columns:
 * {@code classification,limit_pct} — the default per-LP concentration limit as a
 * percent of total uncalled capital. Each chunk is merged into pe-sub-api's
 * {@code cls_conc_limit_defaults} config map via its SERVICE endpoint, so a feed updates
 * only the classes it includes and leaves the rest untouched. The API persists and refreshes
 * its in-memory cache in the same call — no follow-up config reload is needed.
 */
@Configuration
public class ClsConcLimitIngestJobConfig {

    @Bean
    public Job clsConcLimitIngestJob(JobRepository jobRepository,
                                     @Qualifier("clsConcLimitIngestStep") Step clsConcLimitIngestStep) {
        return new JobBuilder("clsConcLimitIngestJob", jobRepository)
                .start(clsConcLimitIngestStep)
                .build();
    }

    @Bean("clsConcLimitIngestStep")
    public Step clsConcLimitIngestStep(JobRepository jobRepository,
                                       PlatformTransactionManager txManager,
                                       @Qualifier("clsConcLimitReader") FlatFileItemReader<ClsConcLimitRow> clsConcLimitReader,
                                       ClsConcLimitRowProcessor clsConcLimitProcessor,
                                       @Qualifier("clsConcLimitWriter") ItemWriter<ProcessedClsConcLimit> clsConcLimitWriter) {
        return new StepBuilder("clsConcLimitIngestStep", jobRepository)
                .<ClsConcLimitRow, ProcessedClsConcLimit>chunk(50)
                .transactionManager(txManager)
                .reader(clsConcLimitReader)
                .processor(clsConcLimitProcessor)
                .writer(clsConcLimitWriter)
                .faultTolerant()
                .skip(Exception.class)
                .skipLimit(10)
                .build();
    }

    @Bean("clsConcLimitReader")
    @StepScope
    public FlatFileItemReader<ClsConcLimitRow> clsConcLimitReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<ClsConcLimitRow>()
                .name("clsConcLimitReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .delimited()
                .quoteCharacter('"')
                .names("classification", "limitPct")
                .fieldSetMapper(fs -> new ClsConcLimitRow(
                        fs.readString("classification"),
                        fs.readString("limitPct")
                ))
                .build();
    }

    @Bean
    public ClsConcLimitRowProcessor clsConcLimitProcessor() {
        return new ClsConcLimitRowProcessor();
    }

    /**
     * Merges the chunk's classifications into the API's defaults map. Later rows for the same
     * classification within a chunk win (LinkedHashMap keeps feed order semantics).
     */
    @Bean("clsConcLimitWriter")
    public ItemWriter<ProcessedClsConcLimit> clsConcLimitWriter(PeSubApiClient apiClient) {
        return chunk -> {
            Map<String, Double> limits = new LinkedHashMap<>();
            chunk.getItems().forEach(item -> limits.put(item.classification(), item.limitPct()));
            if (!limits.isEmpty()) {
                apiClient.mergeClsConcLimitDefaults(limits);
            }
        };
    }
}
