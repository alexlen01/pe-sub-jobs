package com.ubs.pesubjobs.controller;

import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.batch.core.step.StepExecution;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.server.ResponseStatusException;

@RestController
@RequestMapping("/jobs")
public class JobController {

    private final JobOperator jobOperator;
    private final Job facilityIngestJob;
    private final Job lpMasterIngestJob;
    private final Job lpRecordsSeedJob;

    public JobController(JobOperator jobOperator,
                         @Qualifier("facilityIngestJob")  Job facilityIngestJob,
                         @Qualifier("lpMasterIngestJob")  Job lpMasterIngestJob,
                         @Qualifier("lpRecordsSeedJob")   Job lpRecordsSeedJob) {
        this.jobOperator       = jobOperator;
        this.facilityIngestJob = facilityIngestJob;
        this.lpMasterIngestJob = lpMasterIngestJob;
        this.lpRecordsSeedJob  = lpRecordsSeedJob;
    }

    @PostMapping("/{jobName}")
    public ResponseEntity<JobRunResponse> trigger(
            @PathVariable String jobName,
            @RequestParam String filePath) throws Exception {

        Job job = switch (jobName) {
            case "facility-ingest"  -> facilityIngestJob;
            case "lp-master-ingest" -> lpMasterIngestJob;
            case "lp-records-seed"  -> lpRecordsSeedJob;
            default -> throw new ResponseStatusException(HttpStatus.NOT_FOUND, "Unknown job: " + jobName);
        };

        JobParameters params = new JobParametersBuilder()
                .addString("filePath", filePath)
                .addLong("runId", System.currentTimeMillis())
                .toJobParameters();

        JobExecution execution = jobOperator.start(job, params);

        long reads = 0, writes = 0, skips = 0;
        for (StepExecution se : execution.getStepExecutions()) {
            reads  += se.getReadCount();
            writes += se.getWriteCount();
            skips  += se.getSkipCount();
        }

        return ResponseEntity.ok(new JobRunResponse(
                execution.getId(),
                execution.getStatus().name(),
                execution.getExitStatus().getExitCode(),
                reads, writes, skips
        ));
    }

    record JobRunResponse(Long executionId, String status, String exitCode,
                          long readCount, long writeCount, long skipCount) {}
}
