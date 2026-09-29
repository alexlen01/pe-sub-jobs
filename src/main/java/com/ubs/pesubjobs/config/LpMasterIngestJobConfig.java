package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.model.LpMasterRow;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import com.ubs.pesubjobs.processor.LpMasterRowProcessor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.configuration.annotation.JobScope;
import org.springframework.batch.core.configuration.annotation.StepScope;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.Step;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.core.step.tasklet.Tasklet;
import org.springframework.batch.infrastructure.item.ExecutionContext;
import org.springframework.batch.infrastructure.item.ItemReader;
import org.springframework.batch.infrastructure.item.ItemWriter;
import org.springframework.batch.infrastructure.repeat.RepeatStatus;
import org.springframework.batch.infrastructure.item.file.FlatFileItemReader;
import org.springframework.batch.infrastructure.item.file.builder.FlatFileItemReaderBuilder;
import org.springframework.batch.infrastructure.item.support.IteratorItemReader;
import com.ubs.pesubjobs.storage.FeedFileStore;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.transaction.PlatformTransactionManager;

import java.util.List;

@Configuration
public class LpMasterIngestJobConfig {

    private static final Logger log = LoggerFactory.getLogger(LpMasterIngestJobConfig.class);

    @Bean
    public Job lpMasterIngestJob(JobRepository jobRepository,
                                 @Qualifier("lpMasterStageStep") Step lpMasterStageStep,
                                 @Qualifier("lpMasterClearStep") Step lpMasterClearStep,
                                 @Qualifier("lpMasterIngestStep") Step lpMasterIngestStep) {
        // LP Master is repopulated wholesale from the extract feed ("override, do not preserve"),
        // so the load is preceded by a table-wide clear the batch transaction manager cannot roll
        // back. Stage first: the replacement is read and processed in full before anything is
        // deleted, so an unreadable feed or an unprocessable row leaves the existing table intact.
        return new JobBuilder("lpMasterIngestJob", jobRepository)
                .start(lpMasterStageStep)
                .next(lpMasterClearStep)
                .next(lpMasterIngestStep)
                .build();
    }

    /**
     * One replacement per run. Job-scoped rather than a singleton so a staged replacement cannot
     * outlive the run that read it, or leak into the next one.
     */
    @Bean
    @JobScope
    public LpMasterStagingArea lpMasterStagingArea() {
        return new LpMasterStagingArea();
    }

    @Bean("lpMasterStageStep")
    public Step lpMasterStageStep(JobRepository jobRepository,
                                  PlatformTransactionManager txManager,
                                  @Qualifier("lpMasterReader") FlatFileItemReader<LpMasterRow> lpMasterReader,
                                  LpMasterRowProcessor lpMasterProcessor,
                                  LpMasterStagingArea staging) {
        Tasklet tasklet = (contribution, chunkContext) -> {
            lpMasterReader.open(new ExecutionContext());
            try {
                LpMasterRow row;
                while ((row = lpMasterReader.read()) != null) {
                    staging.stage(lpMasterProcessor.process(row));
                    contribution.incrementReadCount();
                }
            } finally {
                lpMasterReader.close();
            }

            // A feed that yields no rows would replace the whole table with nothing. That is the
            // same outcome as the failure this step exists to prevent, so it is refused too — a
            // genuine emptying is a curation action, not a feed.
            if (staging.isEmpty()) {
                throw new IllegalStateException(
                        "LP Master feed contained no rows; refusing to clear the existing table");
            }

            log.info("[lp-master-ingest] staged {} rows for replacement", staging.size());
            return RepeatStatus.FINISHED;
        };
        return new StepBuilder("lpMasterStageStep", jobRepository)
                .tasklet(tasklet, txManager)
                .build();
    }

    @Bean("lpMasterClearStep")
    public Step lpMasterClearStep(JobRepository jobRepository,
                                  PlatformTransactionManager txManager,
                                  PeSubApiClient apiClient) {
        Tasklet tasklet = (contribution, chunkContext) -> {
            apiClient.clearLpMaster();
            return RepeatStatus.FINISHED;
        };
        return new StepBuilder("lpMasterClearStep", jobRepository)
                .tasklet(tasklet, txManager)
                .build();
    }

    @Bean("lpMasterIngestStep")
    public Step lpMasterIngestStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("lpMasterStagedReader") ItemReader<ProcessedLpMaster> lpMasterStagedReader,
                                   @Qualifier("lpMasterWriter") ItemWriter<ProcessedLpMaster> lpMasterWriter,
                                   IngestTallyListener tallyListener) {
        return new StepBuilder("lpMasterIngestStep", jobRepository)
                .<ProcessedLpMaster, ProcessedLpMaster>chunk(50)
                .transactionManager(txManager)
                .reader(lpMasterStagedReader)
                .writer(lpMasterWriter)
                // No skip policy: a row that cannot be written fails the job rather than
                // vanishing from the load. Rows the API itself refuses are counted by the
                // listener, which fails the run past the configured threshold — it matters most
                // here, where the table was cleared before the load began.
                .listener(tallyListener)
                .build();
    }

    /** Reads what the staging step already validated — the feed file is not touched again. */
    @Bean("lpMasterStagedReader")
    @StepScope
    public ItemReader<ProcessedLpMaster> lpMasterStagedReader(LpMasterStagingArea staging) {
        return new IteratorItemReader<>(staging.rows());
    }

    @Bean("lpMasterReader")
    @StepScope
    public FlatFileItemReader<LpMasterRow> lpMasterReader(
            @Value("#{jobParameters['filePath']}") String filePath,
            FeedFileStore feedFileStore) {
        return new FlatFileItemReaderBuilder<LpMasterRow>()
                .name("lpMasterReader")
                .resource(feedFileStore.open(filePath))
                .linesToSkip(1)
                .lineTokenizer(CsvLineTokenizers.lenientQuotedCsvTokenizer(
                        "investorName", "parent", "spv", "investorType",
                        "institutionalOrHnw", "regionLocation", "investmentGrade", "spRating", "moodysRating",
                        "fitchRating", "aum", "nav", "pensionAssets", "fundingRatio",
                        "ubsLpCategory", "ubsDefaultAdvanceRate", "ubsDefaultConcentrationLimit", "notes"))
                .fieldSetMapper(fs -> new LpMasterRow(
                        fs.readString("investorName"),
                        fs.readString("parent"),
                        fs.readString("spv"),
                        fs.readString("investorType"),
                        fs.readString("institutionalOrHnw"),
                        fs.readString("regionLocation"),
                        fs.readString("investmentGrade"),
                        fs.readString("spRating"),
                        fs.readString("moodysRating"),
                        fs.readString("fitchRating"),
                        fs.readString("aum"),
                        fs.readString("nav"),
                        fs.readString("pensionAssets"),
                        fs.readString("fundingRatio"),
                        fs.readString("ubsLpCategory"),
                        fs.readString("ubsDefaultAdvanceRate"),
                        fs.readString("ubsDefaultConcentrationLimit"),
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
    public ItemWriter<ProcessedLpMaster> lpMasterWriter(PeSubApiClient apiClient, IngestTally tally) {
        return chunk -> tally.record(apiClient.ingestLpMaster(List.copyOf(chunk.getItems())));
    }
}
