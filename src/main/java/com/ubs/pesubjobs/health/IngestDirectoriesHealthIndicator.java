package com.ubs.pesubjobs.health;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import com.ubs.pesubjobs.config.IngestProperties;
import org.springframework.boot.health.contributor.Health;
import org.springframework.boot.health.contributor.HealthIndicator;
import org.springframework.stereotype.Component;

import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Reports whether the directories the startup feeds and the BB-template watcher read from are
 * present and readable. Both are configured paths on the pod's own disk ({@code ingest.import-root}
 * and {@code bb-template-import.directory}) — a missing or unreadable one means a feed silently
 * finds nothing rather than failing loudly, which is exactly the case an operator needs surfaced.
 */
@Component("ingestDirectories")
public class IngestDirectoriesHealthIndicator implements HealthIndicator {

    private final Path importRoot;
    private final Path bbTemplateDirectory;

    public IngestDirectoriesHealthIndicator(IngestProperties ingestProperties,
                                             BbTemplateImportProperties bbTemplateImportProperties) {
        this.importRoot = Path.of(ingestProperties.importRoot()).toAbsolutePath().normalize();
        this.bbTemplateDirectory = Path.of(bbTemplateImportProperties.directory()).toAbsolutePath().normalize();
    }

    @Override
    public Health health() {
        boolean importRootReadable = isReadableDirectory(importRoot);
        boolean bbTemplateDirReadable = isReadableDirectory(bbTemplateDirectory);

        Health.Builder builder = (importRootReadable && bbTemplateDirReadable ? Health.up() : Health.down())
            .withDetail("importRoot", importRoot.toString())
            .withDetail("importRootReadable", importRootReadable)
            .withDetail("bbTemplateDirectory", bbTemplateDirectory.toString())
            .withDetail("bbTemplateDirectoryReadable", bbTemplateDirReadable);
        return builder.build();
    }

    private boolean isReadableDirectory(Path path) {
        return Files.isDirectory(path) && Files.isReadable(path);
    }
}
