package com.ubs.pesubjobs.health;

import com.ubs.pesubjobs.storage.BbTemplateStore;
import com.ubs.pesubjobs.storage.FeedFileStore;
import org.springframework.boot.health.contributor.Health;
import org.springframework.boot.health.contributor.HealthIndicator;
import org.springframework.stereotype.Component;

/**
 * Reports whether the locations the startup feeds and the BB-template watcher read from are
 * reachable — a local directory in the {@code local} profile ({@code ingest.import-root} and
 * {@code bb-template-import.directory}), an Azure Blob Storage container/prefix everywhere else. A
 * missing or unreachable one means a feed silently finds nothing rather than failing loudly, which
 * is exactly the case an operator needs surfaced.
 */
@Component("ingestDirectories")
public class IngestDirectoriesHealthIndicator implements HealthIndicator {

    private final FeedFileStore feedFileStore;
    private final BbTemplateStore bbTemplateStore;

    public IngestDirectoriesHealthIndicator(FeedFileStore feedFileStore, BbTemplateStore bbTemplateStore) {
        this.feedFileStore = feedFileStore;
        this.bbTemplateStore = bbTemplateStore;
    }

    @Override
    public Health health() {
        boolean importRootReadable = feedFileStore.isReachable();
        boolean bbTemplateDirReadable = bbTemplateStore.isReachable();

        Health.Builder builder = (importRootReadable && bbTemplateDirReadable ? Health.up() : Health.down())
            .withDetail("importRoot", feedFileStore.describeLocation())
            .withDetail("importRootReadable", importRootReadable)
            .withDetail("bbTemplateDirectory", bbTemplateStore.describeLocation())
            .withDetail("bbTemplateDirectoryReadable", bbTemplateDirReadable);
        return builder.build();
    }
}
