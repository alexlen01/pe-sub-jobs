package com.ubs.pesubjobs;

import com.ubs.pesubjobs.config.IngestProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.Job;
import org.springframework.batch.core.JobExecution;
import org.springframework.batch.core.JobParameters;
import org.springframework.batch.core.JobParametersBuilder;
import org.springframework.batch.core.StepExecution;
import org.springframework.batch.core.launch.JobLauncher;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

@Component
public class JobStartupRunner implements ApplicationRunner {

    private static final Logger log = LoggerFactory.getLogger(JobStartupRunner.class);

    private final JobLauncher jobLauncher;
    private final Job facilityIngestJob;
    private final Job lpMasterIngestJob;
    private final Job lpRecordsSeedJob;
    private final IngestProperties ingestProperties;

    public JobStartupRunner(JobLauncher jobLauncher,
                            @Qualifier("facilityIngestJob")   Job facilityIngestJob,
                            @Qualifier("lpMasterIngestJob")   Job lpMasterIngestJob,
                            @Qualifier("lpRecordsSeedJob")    Job lpRecordsSeedJob,
                            IngestProperties ingestProperties) {
        this.jobLauncher       = jobLauncher;
        this.facilityIngestJob = facilityIngestJob;
        this.lpMasterIngestJob = lpMasterIngestJob;
        this.lpRecordsSeedJob  = lpRecordsSeedJob;
        this.ingestProperties  = ingestProperties;
    }

    @Override
    public void run(ApplicationArguments args) {
        runJob("facility-ingest",   facilityIngestJob,  ingestProperties.facilityFile());
        runJob("lp-master-ingest",  lpMasterIngestJob,  ingestProperties.lpMasterFile());
        runJob("lp-records-seed",   lpRecordsSeedJob,   ingestProperties.lpFacilitySeedsFile());
    }

    private void runJob(String name, Job job, String filePath) {
        log.info("[{}] starting — filePath={}", name, filePath);
        try {
            JobParameters params = new JobParametersBuilder()
                    .addString("filePath", filePath)
                    .addLong("runId", System.currentTimeMillis())
                    .toJobParameters();

            JobExecution execution = jobLauncher.run(job, params);

            long reads  = execution.getStepExecutions().stream().mapToLong(StepExecution::getReadCount).sum();
            long writes = execution.getStepExecutions().stream().mapToLong(StepExecution::getWriteCount).sum();
            long skips  = execution.getStepExecutions().stream().mapToLong(StepExecution::getSkipCount).sum();

            log.info("[{}] finished — status={} reads={} writes={} skips={}",
                    name, execution.getStatus(), reads, writes, skips);
        } catch (Exception e) {
            log.error("[{}] failed — filePath={} error={}", name, filePath, e.getMessage(), e);
        }
    }
}