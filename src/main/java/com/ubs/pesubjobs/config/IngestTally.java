package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.client.PeSubApiClient.ApiIngestSummary;

/**
 * What one feed run actually did to the platform, accumulated across the run's chunks.
 *
 * <p>Spring Batch counts what this app handled: rows read, rows handed to the writer, rows its own
 * skip policy dropped. It cannot count what pe-sub-api then declined to write — the API owns the
 * constraints, so a row can be read, processed, posted and still not exist afterwards. Discarding
 * the API's summary leaves those rows invisible: the run reports a clean write count while the
 * table is short, and a threshold on skipped rows has nothing to measure. This holds the API's
 * answer so the run can be judged on what landed rather than on what was sent.
 *
 * <p>Job-scoped: one tally per run, so a count can neither outlive the run that produced it nor
 * leak into the next one.
 */
public class IngestTally {

    private long created;
    private long updated;
    private long skippedByApi;
    private long droppedBeforeSend;

    /**
     * Folds one chunk's API response into the run.
     *
     * <p>A null summary means the API answered without a body. Nothing is counted for it: the rows
     * were sent, and inventing a created count for them would report a write nobody confirmed.
     */
    public synchronized void record(ApiIngestSummary summary) {
        if (summary == null) return;
        created += summary.created();
        updated += summary.updated();
        skippedByApi += summary.skipped();
    }

    /** Rows this app dropped before they ever reached the API — batch skips and processor filters. */
    public synchronized void recordDroppedBeforeSend(long rows) {
        droppedBeforeSend += rows;
    }

    public synchronized long created() {
        return created;
    }

    public synchronized long updated() {
        return updated;
    }

    /**
     * Every row the feed carried that is not in the platform: dropped on the way out, or refused on
     * arrival. The two are one number here because they have one consequence — a row the operator
     * believes they loaded is missing.
     */
    public synchronized long rowsNotLanded() {
        return skippedByApi + droppedBeforeSend;
    }
}
