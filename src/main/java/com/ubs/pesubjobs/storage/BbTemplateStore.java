package com.ubs.pesubjobs.storage;

import org.springframework.core.io.Resource;

import java.io.IOException;
import java.util.List;

/**
 * Where finished BB template workbooks are dropped for {@link com.ubs.pesubjobs.BbTemplateDirectoryImporter}
 * to pick up — a local directory in the {@code local} profile, an Azure Blob Storage container
 * everywhere else. The importer holds no parser and does no filesystem/blob calls of its own: it
 * asks this store what is currently there, and reads one workbook's bytes back through it to post
 * to pe-sub-api.
 */
public interface BbTemplateStore {

    /**
     * Every regular object currently in the configured drop location, in a stable name order so a
     * bad file earlier in the sweep does not affect this run's outcome for the ones after it.
     *
     * @throws IOException if the location could not be listed (missing directory, blob call failed)
     */
    List<BbTemplateObject> list() throws IOException;

    /** A {@link Resource} for the object's bytes, to attach as the multipart upload body. */
    Resource open(String identifier);

    /** True if the configured drop location itself is currently reachable (for health checks). */
    boolean isReachable();

    /** A human-readable description of the configured drop location, safe to expose in health details. */
    String describeLocation();
}
