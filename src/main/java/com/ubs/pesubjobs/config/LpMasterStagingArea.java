package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.model.ProcessedLpMaster;

import java.util.ArrayList;
import java.util.List;

/**
 * Holds the fully read and validated LP Master replacement between the staging step and the load
 * step of one run.
 *
 * <p>LP Master is repopulated wholesale, and the remote clear that precedes the load cannot be
 * rolled back by the batch transaction manager. The replacement is therefore assembled in full
 * before anything is deleted: if the feed cannot be read or a row cannot be processed, the run
 * fails here and the existing table is still intact.
 *
 * <p>Job-scoped, so one run's replacement is never visible to another.
 */
public class LpMasterStagingArea {

    private final List<ProcessedLpMaster> rows = new ArrayList<>();

    public void stage(ProcessedLpMaster row) {
        rows.add(row);
    }

    /** The staged replacement, in feed order. */
    public List<ProcessedLpMaster> rows() {
        return List.copyOf(rows);
    }

    public boolean isEmpty() {
        return rows.isEmpty();
    }

    public int size() {
        return rows.size();
    }
}
