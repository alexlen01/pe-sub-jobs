package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.config.IngestTallyListener;
import com.ubs.pesubjobs.storage.FeedFileStore;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.batch.infrastructure.item.ExecutionContext;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.time.Duration;
import java.time.Instant;
import java.util.OptionalLong;

@Component
public class JobStartupRunner implements ApplicationRunner {

    private static final Logger log = LoggerFactory.getLogger(JobStartupRunner.class);

    private final JobOperator jobOperator;
    private final Job umbrellaIngestJob;
    private final Job facilityIngestJob;
    private final Job lpMasterIngestJob;
    private final Job lpRecordsSeedJob;
    private final Job clsConcLimitIngestJob;
    private final IngestProperties ingestProperties;
    private final PeSubApiClient apiClient;
    private final FeedFileStore feedFileStore;

    public JobStartupRunner(JobOperator jobOperator,
                            @Qualifier("umbrellaIngestJob") Job umbrellaIngestJob,
                            @Qualifier("facilityIngestJob") Job facilityIngestJob,
                            @Qualifier("lpMasterIngestJob") Job lpMasterIngestJob,
                            @Qualifier("lpRecordsSeedJob") Job lpRecordsSeedJob,
                            @Qualifier("clsConcLimitIngestJob") Job clsConcLimitIngestJob,
                            IngestProperties ingestProperties,
                            PeSubApiClient apiClient,
                            FeedFileStore feedFileStore) {
        this.jobOperator = jobOperator;
        this.umbrellaIngestJob = umbrellaIngestJob;
        this.facilityIngestJob = facilityIngestJob;
        this.lpMasterIngestJob = lpMasterIngestJob;
        this.lpRecordsSeedJob = lpRecordsSeedJob;
        this.clsConcLimitIngestJob = clsConcLimitIngestJob;
        this.ingestProperties = ingestProperties;
        this.apiClient = apiClient;
        this.feedFileStore = feedFileStore;
    }

    @Override
    public void run(ApplicationArguments args) {
        if (!ingestProperties.runOnStartup()) {
            log.info("Startup ingest disabled (ingest.run-on-startup=false) - skipping seed jobs");
            return;
        }
        if (!waitForApi()) {
            log.warn("Startup ingest skipped because pe-sub-api did not become reachable within {}. Start pe-sub-api first or run /jobs once it is up.",
                    ingestProperties.schemaWaitTimeout());
            return;
        }

        PopulationState state = populationState();
        if (state == PopulationState.UNKNOWN) {
            // The feeds below replace LP Master wholesale. Running them on an unanswered question
            // risks emptying a populated platform, so an uncertain check is a stop, not a go.
            log.warn("Startup ingest skipped because the platform's populated state could not be established. Fix the API and restart, or run /jobs once the counts are readable.");
            return;
        }
        if (state == PopulationState.POPULATED) {
            log.info("Database already populated with seed data - skipping startup ingest");
            return;
        }

        // Groups first: a facility names its umbrella on its own row, so the group must already be
        // onboarded under the name the agent printed. A feed set written before umbrellas.csv
        // existed simply has no file — the facilities then fall back to grouping by account number,
        // which is the behaviour that feed set was extracted for.
        String umbrellaFile = ingestProperties.umbrellaFile();
        if (umbrellaFile == null || umbrellaFile.isBlank()) {
            log.info("[umbrella-ingest] skipped - no feed file configured (ingest.umbrella-file)");
        } else if (!feedFileStore.exists(umbrellaFile)) {
            log.warn("[umbrella-ingest] skipped - no readable feed file at {}. Facility groups will be named from what the facility feed carries.",
                    umbrellaFile);
        } else {
            runJob("umbrella-ingest", umbrellaIngestJob, umbrellaFile);
        }
        runJob("facility-ingest", facilityIngestJob, ingestProperties.facilityFile());
        runJob("lp-master-ingest", lpMasterIngestJob, ingestProperties.lpMasterFile());
        runJob("lp-records-seed", lpRecordsSeedJob, ingestProperties.lpFacilitySeedsFile());
        // Optional feed: the API seeds class defaults on schema creation, so this only runs
        // when a feed file is explicitly configured (CLS_CONC_LIMITS_FILE).
        String clsConcFile = ingestProperties.clsConcLimitsFile();
        if (clsConcFile != null && !clsConcFile.isBlank()) {
            runJob("cls-conc-limits-ingest", clsConcLimitIngestJob, clsConcFile);
        } else {
            log.info("[cls-conc-limits-ingest] skipped - no feed file configured (ingest.cls-conc-limits-file)");
        }
    }

    // All feeds go through pe-sub-api's endpoints, so the API answering /api/ping IS the
    // readiness signal — it only serves once its own Flyway migrations have run.
    private boolean waitForApi() {
        Duration timeout = ingestProperties.schemaWaitTimeout();
        Duration interval = ingestProperties.schemaWaitInterval();
        Instant deadline = Instant.now().plus(timeout);

        while (true) {
            if (apiClient.isApiReady()) {
                return true;
            }
            if (!Instant.now().isBefore(deadline)) {
                return false;
            }
            try {
                Thread.sleep(Math.max(250L, interval.toMillis()));
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                return false;
            }
        }
    }

    /** What the row counts say about the platform — including that they say nothing. */
    private enum PopulationState { EMPTY, POPULATED, UNKNOWN }

    /**
     * Answers only from counts the API actually returned. One unreadable count makes the whole
     * answer UNKNOWN: "no rows" and "could not ask" lead to opposite actions, and the destructive
     * one must never be reached by default.
     */
    private PopulationState populationState() {
        OptionalLong facilityCount = apiClient.getFacilityCount();
        OptionalLong lpMasterCount = apiClient.getLpMasterCount();
        OptionalLong lpRecordCount = apiClient.getLpRecordCount();

        if (facilityCount.isEmpty() || lpMasterCount.isEmpty() || lpRecordCount.isEmpty()) {
            log.warn("Data existence check inconclusive: facilities={} lpMaster={} lpRecords={}",
                    describe(facilityCount), describe(lpMasterCount), describe(lpRecordCount));
            return PopulationState.UNKNOWN;
        }

        boolean hasData = facilityCount.getAsLong() > 0
                || lpMasterCount.getAsLong() > 0
                || lpRecordCount.getAsLong() > 0;
        if (hasData) {
            log.info("Data existence check: facilities={} lpMaster={} lpRecords={}",
                    facilityCount.getAsLong(), lpMasterCount.getAsLong(), lpRecordCount.getAsLong());
        }
        return hasData ? PopulationState.POPULATED : PopulationState.EMPTY;
    }

    private String describe(OptionalLong count) {
        return count.isPresent() ? Long.toString(count.getAsLong()) : "unknown";
    }

    private void runJob(String name, Job job, String filePath) {
        log.info("[{}] starting - filePath={}", name, filePath);
        try {
            JobParameters params = new JobParametersBuilder()
                    .addString("filePath", filePath)
                    .addLong("runId", System.currentTimeMillis())
                    .toJobParameters();

            JobExecution execution = jobOperator.start(job, params);

            long reads = execution.getStepExecutions().stream().mapToLong(step -> step != null ? step.getReadCount() : 0).sum();
            long writes = execution.getStepExecutions().stream().mapToLong(step -> step != null ? step.getWriteCount() : 0).sum();

            ExecutionContext ctx = execution.getExecutionContext();
            long created = ctx.containsKey(IngestTallyListener.CREATED_KEY)
                    ? ctx.getLong(IngestTallyListener.CREATED_KEY) : 0;
            long updated = ctx.containsKey(IngestTallyListener.UPDATED_KEY)
                    ? ctx.getLong(IngestTallyListener.UPDATED_KEY) : 0;
            long notLanded = ctx.containsKey(IngestTallyListener.NOT_LANDED_KEY)
                    ? ctx.getLong(IngestTallyListener.NOT_LANDED_KEY) : 0;

            log.info("[{}] finished - status={} reads={} writes={} created={} updated={} rowsNotLanded={}",
                    name, execution.getStatus(), reads, writes, created, updated, notLanded);
        } catch (Exception e) {
            log.error("[{}] failed - filePath={} error={}", name, filePath, e.getMessage(), e);
        }
    }
}
