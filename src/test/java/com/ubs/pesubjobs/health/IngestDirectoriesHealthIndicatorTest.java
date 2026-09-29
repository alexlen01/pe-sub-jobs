package com.ubs.pesubjobs.health;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.storage.LocalBbTemplateStore;
import com.ubs.pesubjobs.storage.LocalFeedFileStore;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;

import static org.assertj.core.api.Assertions.assertThat;

class IngestDirectoriesHealthIndicatorTest {

    @TempDir Path tempDir;

    @Test
    void bothDirectoriesReadable_reportsUp() throws IOException {
        Path importRoot = Files.createDirectory(tempDir.resolve("out"));
        Path bbTemplateDir = Files.createDirectory(tempDir.resolve("bb-templates"));

        IngestDirectoriesHealthIndicator indicator = new IngestDirectoriesHealthIndicator(
                new LocalFeedFileStore(ingestProperties(importRoot)),
                new LocalBbTemplateStore(bbTemplateProperties(bbTemplateDir)));

        var health = indicator.health();
        assertThat(health.getStatus().getCode()).isEqualTo("UP");
        assertThat(health.getDetails()).containsEntry("importRootReadable", true);
        assertThat(health.getDetails()).containsEntry("bbTemplateDirectoryReadable", true);
    }

    @Test
    void missingImportRoot_reportsDown() throws IOException {
        Path importRoot = tempDir.resolve("does-not-exist");
        Path bbTemplateDir = Files.createDirectory(tempDir.resolve("bb-templates"));

        IngestDirectoriesHealthIndicator indicator = new IngestDirectoriesHealthIndicator(
                new LocalFeedFileStore(ingestProperties(importRoot)),
                new LocalBbTemplateStore(bbTemplateProperties(bbTemplateDir)));

        var health = indicator.health();
        assertThat(health.getStatus().getCode()).isEqualTo("DOWN");
        assertThat(health.getDetails()).containsEntry("importRootReadable", false);
        assertThat(health.getDetails()).containsEntry("bbTemplateDirectoryReadable", true);
    }

    private IngestProperties ingestProperties(Path importRoot) {
        return new IngestProperties("facilities.csv", "umbrellas.csv", "lp_master.csv",
                "lp_facility_seeds.csv", "", "http://pe-sub-api", importRoot.toString(), 10,
                Duration.ofSeconds(30), Duration.ofSeconds(2), true);
    }

    private BbTemplateImportProperties bbTemplateProperties(Path directory) {
        return new BbTemplateImportProperties(true, directory.toString(), "http://pe-sub-api",
                Duration.ofSeconds(30), Duration.ofSeconds(2));
    }
}
