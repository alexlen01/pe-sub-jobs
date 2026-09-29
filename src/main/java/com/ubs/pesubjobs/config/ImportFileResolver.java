package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.storage.FeedFileNotAllowedException;
import com.ubs.pesubjobs.storage.FeedFileStore;
import org.springframework.stereotype.Component;

/**
 * Resolves a caller-supplied feed file name to a reader identifier via the active {@link
 * FeedFileStore} — an absolute local path in the {@code local} profile, an Azure blob name
 * everywhere else. Kept as a thin, named seam so {@code JobController}'s dependency and intent —
 * "turn a caller-supplied bare name into something the readers can open" — do not change with the
 * storage backend.
 */
@Component
public class ImportFileResolver {

    private final FeedFileStore feedFileStore;

    public ImportFileResolver(FeedFileStore feedFileStore) {
        this.feedFileStore = feedFileStore;
    }

    /**
     * @param fileName a bare file name such as {@code facilities.csv}
     * @return the identifier to read (an absolute local path, or a blob name)
     * @throws FeedFileNotAllowedException if the name is not an allowed, existing feed
     */
    public String resolve(String fileName) {
        return feedFileStore.resolve(fileName);
    }
}
