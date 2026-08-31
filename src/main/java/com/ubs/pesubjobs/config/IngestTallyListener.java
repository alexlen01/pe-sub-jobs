package com.ubs.pesubjobs.config;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.core.BatchStatus;
import org.springframework.batch.core.ExitStatus;
import org.springframework.batch.core.listener.StepExecutionListener;
import org.springframework.batch.core.step.StepExecution;
import org.springframework.batch.infrastructure.item.ExecutionContext;

/**
 * Closes a feed step on the numbers that matter: what pe-sub-api created and updated, and how many
 * rows the feed carried that are not in the platform afterwards.
 *
 * <p>Two jobs are done here. The counts are published onto the job execution so the trigger
 * response and the startup log can state them — a run that silently drops rows is worse than one
 * that fails, because nobody goes looking. And the documented threshold is applied: past it, the
 * run is a failure rather than a success with a footnote, because a feed that lost more than a
 * handful of rows has produced reference data no one should lend against.
 */
public class IngestTallyListener implements StepExecutionListener {

    private static final Logger log = LoggerFactory.getLogger(IngestTallyListener.class);

    public static final String CREATED_KEY = "ingest.created";
    public static final String UPDATED_KEY = "ingest.updated";
    public static final String NOT_LANDED_KEY = "ingest.rowsNotLanded";

    private final IngestTally tally;
    private final long maxRowsNotLanded;

    public IngestTallyListener(IngestTally tally, long maxRowsNotLanded) {
        this.tally = tally;
        this.maxRowsNotLanded = maxRowsNotLanded;
    }

    @Override
    public ExitStatus afterStep(StepExecution stepExecution) {
        // Rows this app dropped on the way out. A processor that returns null filters an item it
        // could not parse, which Batch counts separately from a skip — both are a row the operator
        // handed over and will not find later, so both count here.
        tally.recordDroppedBeforeSend(stepExecution.getSkipCount() + stepExecution.getFilterCount());

        long created = tally.created();
        long updated = tally.updated();
        long notLanded = tally.rowsNotLanded();

        ExecutionContext jobContext = stepExecution.getJobExecution().getExecutionContext();
        jobContext.putLong(CREATED_KEY, created);
        jobContext.putLong(UPDATED_KEY, updated);
        jobContext.putLong(NOT_LANDED_KEY, notLanded);

        log.info("[{}] step finished - reads={} writes={} created={} updated={} rowsNotLanded={}",
                stepExecution.getStepName(), stepExecution.getReadCount(),
                stepExecution.getWriteCount(), created, updated, notLanded);

        if (notLanded > maxRowsNotLanded) {
            String reason = "%d of the feed's rows did not land (limit %d): pe-sub-api refused them or they were dropped before being sent"
                    .formatted(notLanded, maxRowsNotLanded);
            log.error("[{}] {}", stepExecution.getStepName(), reason);
            // Failing here rather than by throwing: Spring Batch runs afterStep inside its finally
            // block and only logs an exception raised there, so the status has to be set outright
            // for the job to report the failure to whoever triggered it.
            stepExecution.setStatus(BatchStatus.FAILED);
            return ExitStatus.FAILED.addExitDescription(reason);
        }
        return stepExecution.getExitStatus();
    }
}
