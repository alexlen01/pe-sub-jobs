package com.ubs.pesubjobs.controller;

import com.ubs.pesubjobs.config.ImportFileResolver;
import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.config.IngestTallyListener;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.job.parameters.JobParametersBuilder;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.batch.core.step.StepExecution;
import org.springframework.batch.infrastructure.item.ExecutionContext;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.server.ResponseStatusException;

import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * On-demand feed trigger. Every route requires the SERVICE role (see the security configuration)
 * because each job replaces reference data pe-sub-api owns.
 *
 * <p>Callers name a job, not a path. The job's configured feed file is used by default; an
 * optional {@code file} parameter selects a different bare file name inside the configured import
 * directory. A caller can therefore choose <em>which</em> file in a known directory to load, and
 * cannot make the service read anywhere else on the host.
 */
@RestController
@RequestMapping("/jobs")
public class JobController {

    private static final Logger log = LoggerFactory.getLogger(JobController.class);

    /**
     * One run of a given feed at a time. Jobs execute on the calling request thread, and the
     * LP Master feed clears the table before repopulating it — so two overlapping runs of the
     * same job can have the second clear away what the first has already written, leaving the
     * platform holding part of one feed. Different feeds do not collide and are not serialised
     * against each other.
     */
    private final Set<String> running = ConcurrentHashMap.newKeySet();

    private final JobOperator jobOperator;
    private final IngestProperties ingestProperties;
    private final ImportFileResolver importFileResolver;
    private final Job umbrellaIngestJob;
    private final Job facilityIngestJob;
    private final Job lpMasterIngestJob;
    private final Job lpRecordsSeedJob;
    private final Job clsConcLimitIngestJob;

    public JobController(JobOperator jobOperator,
                         IngestProperties ingestProperties,
                         ImportFileResolver importFileResolver,
                         @Qualifier("umbrellaIngestJob")      Job umbrellaIngestJob,
                         @Qualifier("facilityIngestJob")      Job facilityIngestJob,
                         @Qualifier("lpMasterIngestJob")      Job lpMasterIngestJob,
                         @Qualifier("lpRecordsSeedJob")       Job lpRecordsSeedJob,
                         @Qualifier("clsConcLimitIngestJob")  Job clsConcLimitIngestJob) {
        this.jobOperator           = jobOperator;
        this.ingestProperties      = ingestProperties;
        this.importFileResolver    = importFileResolver;
        this.umbrellaIngestJob     = umbrellaIngestJob;
        this.facilityIngestJob     = facilityIngestJob;
        this.lpMasterIngestJob     = lpMasterIngestJob;
        this.lpRecordsSeedJob      = lpRecordsSeedJob;
        this.clsConcLimitIngestJob = clsConcLimitIngestJob;
    }

    @PostMapping("/{jobName}")
    public ResponseEntity<JobRunResponse> trigger(
            @PathVariable String jobName,
            @RequestParam(required = false) String file,
            Authentication authentication) throws Exception {

        Job job = switch (jobName) {
            case "umbrella-ingest"        -> umbrellaIngestJob;
            case "facility-ingest"        -> facilityIngestJob;
            case "lp-master-ingest"       -> lpMasterIngestJob;
            case "lp-records-seed"        -> lpRecordsSeedJob;
            case "cls-conc-limits-ingest" -> clsConcLimitIngestJob;
            default -> throw new ResponseStatusException(HttpStatus.NOT_FOUND, "Unknown job: " + jobName);
        };

        String filePath = resolveFeedPath(jobName, file);
        String caller = authentication != null ? authentication.getName() : "unknown";

        // Claimed before the run and released in the finally below, so a job that throws does not
        // leave its name latched and the feed unrunnable until restart.
        if (!running.add(jobName)) {
            log.warn("AUDIT job-trigger refused by={} job={} cause=already-running", caller, jobName);
            throw new ResponseStatusException(HttpStatus.CONFLICT,
                    jobName + " is already running; wait for it to finish before triggering it again.");
        }
        try {
            // Audit the intent before the run: a job that dies mid-execution must still leave a
            // record of who asked for it and against which feed.
            log.info("AUDIT job-trigger requested by={} job={} file={}", caller, jobName, filePath);

            JobParameters params = new JobParametersBuilder()
                    .addString("filePath", filePath)
                    .addLong("runId", System.currentTimeMillis())
                    .toJobParameters();

            JobExecution execution = jobOperator.start(job, params);

            long reads = 0, writes = 0;
            for (StepExecution se : execution.getStepExecutions()) {
                reads  += se.getReadCount();
                writes += se.getWriteCount();
            }

            // What pe-sub-api reported back, published onto the job by IngestTallyListener. A write
            // count only says a chunk was posted; these say what the platform actually holds now.
            ExecutionContext ctx = execution.getExecutionContext();
            long created = ctx.containsKey(IngestTallyListener.CREATED_KEY)
                    ? ctx.getLong(IngestTallyListener.CREATED_KEY) : 0;
            long updated = ctx.containsKey(IngestTallyListener.UPDATED_KEY)
                    ? ctx.getLong(IngestTallyListener.UPDATED_KEY) : 0;
            long notLanded = ctx.containsKey(IngestTallyListener.NOT_LANDED_KEY)
                    ? ctx.getLong(IngestTallyListener.NOT_LANDED_KEY) : 0;

            log.info("AUDIT job-trigger completed by={} job={} executionId={} status={} reads={} writes={} created={} updated={} rowsNotLanded={}",
                    caller, jobName, execution.getId(), execution.getStatus(), reads, writes,
                    created, updated, notLanded);

            return ResponseEntity.ok(new JobRunResponse(
                    execution.getId(),
                    execution.getStatus().name(),
                    execution.getExitStatus().getExitCode(),
                    reads, writes, created, updated, notLanded
            ));
        } finally {
            running.remove(jobName);
        }
    }

    /**
     * The configured feed for the job, or the named file inside the import root when the caller
     * supplied one. The configured paths come from deployment configuration, not the request, so
     * they are used as given and keep pointing at whichever feed set the environment selected.
     */
    private String resolveFeedPath(String jobName, String requestedFile) {
        if (requestedFile != null && !requestedFile.isBlank()) {
            return importFileResolver.resolve(requestedFile).toString();
        }
        String configured = switch (jobName) {
            case "umbrella-ingest"        -> ingestProperties.umbrellaFile();
            case "facility-ingest"        -> ingestProperties.facilityFile();
            case "lp-master-ingest"       -> ingestProperties.lpMasterFile();
            case "lp-records-seed"        -> ingestProperties.lpFacilitySeedsFile();
            case "cls-conc-limits-ingest" -> ingestProperties.clsConcLimitsFile();
            default -> null;
        };
        if (configured == null || configured.isBlank()) {
            throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                    "No feed file is configured for " + jobName + "; name one with the 'file' parameter.");
        }
        return configured;
    }

    /**
     * @param rowsNotLanded rows the feed carried that the platform does not hold afterwards —
     *                      refused by pe-sub-api, or dropped before being sent. Past the
     *                      configured limit the run is reported {@code FAILED}.
     */
    record JobRunResponse(Long executionId, String status, String exitCode,
                          long readCount, long writeCount,
                          long createdCount, long updatedCount, long rowsNotLanded) {}
}
