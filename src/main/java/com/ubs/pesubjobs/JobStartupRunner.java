package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.config.IngestProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.time.Duration;
import java.time.Instant;

@Component
public class JobStartupRunner implements ApplicationRunner {

    private static final Logger log = LoggerFactory.getLogger(JobStartupRunner.class);

    private final JobOperator jobOperator;
    private final Job facilityIngestJob;
    private final Job lpMasterIngestJob;
    private final Job lpRecordsSeedJob;
    private final Job clsConcLimitIngestJob;
    private final IngestProperties ingestProperties;
    private final PeSubApiClient apiClient;

    public JobStartupRunner(JobOperator jobOperator,
                            @Qualifier("facilityIngestJob") Job facilityIngestJob,
                            @Qualifier("lpMasterIngestJob") Job lpMasterIngestJob,
                            @Qualifier("lpRecordsSeedJob") Job lpRecordsSeedJob,
                            @Qualifier("clsConcLimitIngestJob") Job clsConcLimitIngestJob,
                            IngestProperties ingestProperties,
                            PeSubApiClient apiClient) {
        this.jobOperator = jobOperator;
        this.facilityIngestJob = facilityIngestJob;
        this.lpMasterIngestJob = lpMasterIngestJob;
        this.lpRecordsSeedJob = lpRecordsSeedJob;
        this.clsConcLimitIngestJob = clsConcLimitIngestJob;
        this.ingestProperties = ingestProperties;
        this.apiClient = apiClient;
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

        if (dataAlreadyPopulated()) {
            log.info("Database already populated with seed data - skipping startup ingest");
            return;
        }

        runJob("facility-ingest", facilityIngestJob, ingestProperties.facilityFile());
        runJob("lp-master-ingest", lpMasterIngestJob, ingestProperties.lpMasterFile());
        runJob("lp-records-seed", lpRecordsSeedJob, ingestProperties.lpFacilitySeedsFile());
        // Optional feed: the API's V1_5 migration seeds class defaults, so this only runs
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

    private boolean dataAlreadyPopulated() {
        long facilityCount = apiClient.getFacilityCount();
        long lpMasterCount = apiClient.getLpMasterCount();
        long lpRecordCount = apiClient.getLpRecordCount();

        boolean hasData = facilityCount > 0 || lpMasterCount > 0 || lpRecordCount > 0;
        if (hasData) {
            log.info("Data existence check: facilities={} lpMaster={} lpRecords={}",
                    facilityCount, lpMasterCount, lpRecordCount);
        }
        return hasData;
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
            long skips = execution.getStepExecutions().stream().mapToLong(step -> step != null ? step.getSkipCount() : 0).sum();

            log.info("[{}] finished - status={} reads={} writes={} skips={}",
                    name, execution.getStatus(), reads, writes, skips);
        } catch (Exception e) {
            log.error("[{}] failed - filePath={} error={}", name, filePath, e.getMessage(), e);
        }
    }
}
