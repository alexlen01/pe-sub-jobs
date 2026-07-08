package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.model.ClsConcLimitRow;
import com.ubs.pesubjobs.model.ProcessedClsConcLimit;
import com.ubs.pesubjobs.processor.ClsConcLimitRowProcessor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.configuration.annotation.StepScope;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.listener.JobExecutionListener;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.Step;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.infrastructure.item.database.JdbcBatchItemWriter;
import org.springframework.batch.infrastructure.item.database.builder.JdbcBatchItemWriterBuilder;
import org.springframework.batch.infrastructure.item.file.FlatFileItemReader;
import org.springframework.batch.infrastructure.item.file.builder.FlatFileItemReaderBuilder;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.FileSystemResource;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.web.client.RestClient;

import javax.sql.DataSource;

/**
 * Classification concentration-limit defaults feed. CSV columns:
 * {@code classification,limit_pct} — the default per-LP concentration limit as a
 * percent of total uncalled capital. Rows merge into the single
 * {@code cls_conc_limit_defaults} config row (jsonb map keyed by classification),
 * so a feed updates only the classes it includes and leaves the rest untouched.
 */
@Configuration
public class ClsConcLimitIngestJobConfig {

    private static final Logger log = LoggerFactory.getLogger(ClsConcLimitIngestJobConfig.class);

    @Bean
    public Job clsConcLimitIngestJob(JobRepository jobRepository,
                                     @Qualifier("clsConcLimitIngestStep") Step clsConcLimitIngestStep,
                                     @Qualifier("configReloadListener") JobExecutionListener configReloadListener) {
        return new JobBuilder("clsConcLimitIngestJob", jobRepository)
                .start(clsConcLimitIngestStep)
                .listener(configReloadListener)
                .build();
    }

    @Bean("clsConcLimitIngestStep")
    public Step clsConcLimitIngestStep(JobRepository jobRepository,
                                       PlatformTransactionManager txManager,
                                       @Qualifier("clsConcLimitReader") FlatFileItemReader<ClsConcLimitRow> clsConcLimitReader,
                                       ClsConcLimitRowProcessor clsConcLimitProcessor,
                                       @Qualifier("clsConcLimitWriter") JdbcBatchItemWriter<ProcessedClsConcLimit> clsConcLimitWriter) {
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

    @Bean("clsConcLimitWriter")
    public JdbcBatchItemWriter<ProcessedClsConcLimit> clsConcLimitWriter(DataSource dataSource) {
        // jsonb || merges per classification: feed rows overwrite matching keys and
        // preserve every class the feed does not mention.
        String sql = """
                INSERT INTO config (key, value)
                VALUES ('cls_conc_limit_defaults', jsonb_build_object(:classification::text, :limitPct::numeric))
                ON CONFLICT (key) DO UPDATE SET value = config.value || EXCLUDED.value
                """;

        return new JdbcBatchItemWriterBuilder<ProcessedClsConcLimit>()
                .dataSource(dataSource)
                .sql(sql)
                .itemSqlParameterSourceProvider(item -> {
                    MapSqlParameterSource params = new MapSqlParameterSource();
                    params.addValue("classification", item.classification());
                    params.addValue("limitPct",       item.limitPct());
                    return params;
                })
                .build();
    }

    /**
     * pe-sub-api holds the config table in an in-memory cache loaded at startup, so a
     * DB-side feed is invisible to it until reloaded. Best-effort: a failed reload only
     * logs a warning — the fed values are picked up on the next API restart regardless.
     */
    @Bean("configReloadListener")
    public JobExecutionListener configReloadListener(IngestProperties ingestProperties) {
        String baseUrl = ingestProperties.apiBaseUrl().replaceAll("/+$", "");
        RestClient rest = RestClient.builder().baseUrl(baseUrl).build();
        return new JobExecutionListener() {
            @Override
            public void afterJob(JobExecution jobExecution) {
                if (jobExecution.getStatus() != BatchStatus.COMPLETED) return;
                try {
                    rest.post().uri("/api/config/reload").retrieve().toBodilessEntity();
                    log.info("pe-sub-api config cache reloaded after cls-conc-limits feed");
                } catch (Exception e) {
                    log.warn("pe-sub-api config reload failed ({}); fed values apply on the next API restart",
                            e.getMessage());
                }
            }
        };
    }
}
