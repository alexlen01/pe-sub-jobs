package com.ubs.pesubjobs.storage;

import org.springframework.core.io.Resource;

/**
 * Where the umbrellas/facilities/lp_master/lp_facility_seeds/cls_conc_limits feed CSVs live, and
 * how to reach one — a local directory in the {@code local} profile (docker-compose, developer
 * workstation), an Azure Blob Storage container everywhere else (AKS).
 *
 * <p>Job readers and {@link com.ubs.pesubjobs.JobStartupRunner} deal only in the {@code String}
 * identifier this interface hands back: an absolute filesystem path locally, a blob name on Azure.
 * Neither carries filesystem traversal or symlink semantics, but a caller-supplied name is still
 * validated in {@link #resolve(String)} against the configured location before it is trusted.
 */
public interface FeedFileStore {

    /**
     * @param identifier a previously resolved or configured identifier (never caller input)
     * @return true if a feed exists at the identifier and can be opened
     */
    boolean exists(String identifier);

    /**
     * @param identifier a previously resolved or configured identifier (never caller input)
     * @return a {@link Resource} the batch readers can open for the feed's content
     */
    Resource open(String identifier);

    /**
     * Resolves a caller-supplied bare file name (e.g. {@code facilities.csv}) — never a path — into
     * an identifier {@link #open(String)} can read. Rejects anything that could reach outside the
     * configured feed location, and anything that is not an existing, readable CSV there.
     *
     * @throws FeedFileNotAllowedException if the name is not an allowed, existing feed
     */
    String resolve(String fileName);

    /** True if the configured feed location itself is currently reachable (for health checks). */
    boolean isReachable();

    /** A human-readable description of the configured feed location, safe to expose in health details. */
    String describeLocation();
}
