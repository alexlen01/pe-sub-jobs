package com.ubs.pesubjobs;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.MediaType;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.ResourceAccessException;
import org.springframework.web.client.RestClient;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.Comparator;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Stream;

@Component
public class BbTemplateDirectoryImporter implements ApplicationRunner {

    private static final Logger log = LoggerFactory.getLogger(BbTemplateDirectoryImporter.class);

    private final BbTemplateImportProperties props;
    private final RestClient restClient;
    private final Map<Path, Fingerprint> imported = new ConcurrentHashMap<>();
    private final AtomicBoolean scanRunning = new AtomicBoolean(false);

    public BbTemplateDirectoryImporter(BbTemplateImportProperties props) {
        this.props = props;
        this.restClient = RestClient.builder().baseUrl(trimTrailingSlash(props.apiBaseUrl())).build();
    }

    @Override
    public void run(ApplicationArguments args) {
        scan("startup");
    }

    @Scheduled(fixedDelayString = "${bb-template-import.scan-interval:30s}")
    public void scheduledScan() {
        scan("scheduled");
    }

    private void scan(String reason) {
        if (!props.enabled()) return;
        if (!scanRunning.compareAndSet(false, true)) {
            log.debug("BB template import scan skipped reason={} cause=scan-already-running", reason);
            return;
        }
        Path dir = Path.of(props.directory()).toAbsolutePath().normalize();
        try {
            Files.createDirectories(dir);
        } catch (IOException e) {
            log.warn("BB template import scan skipped reason={} directory={} error={}", reason, dir, e.getMessage(), e);
            scanRunning.set(false);
            return;
        }
        if (!apiAvailable(reason)) {
            scanRunning.set(false);
            return;
        }

        try (Stream<Path> files = Files.list(dir)) {
            files
                .filter(Files::isRegularFile)
                .filter(this::isImportWorkbook)
                .sorted(Comparator.comparing(path -> path.getFileName().toString()))
                .takeWhile(path -> importIfChanged(path, reason))
                .forEach(path -> {});
        } catch (IOException e) {
            log.warn("BB template import scan failed reason={} directory={} error={}", reason, dir, e.getMessage(), e);
        } finally {
            scanRunning.set(false);
        }
    }

    private boolean importIfChanged(Path path, String reason) {
        try {
            Fingerprint fp = fingerprint(path);
            if (!isStable(fp)) return true;
            if (fp.equals(imported.get(path))) return true;

            MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
            body.add("file", new FileSystemResource(path));
            restClient.post()
                .uri("/api/bb-templates/import?mode=upsert")
                .contentType(MediaType.MULTIPART_FORM_DATA)
                .body(body)
                .retrieve()
                .toBodilessEntity();

            imported.put(path, fp);
            log.info("BB template imported reason={} file={}", reason, path);
            return true;
        } catch (ResourceAccessException e) {
            log.warn("BB template import scan paused reason={} file={} apiBaseUrl={} error={}",
                reason, path, trimTrailingSlash(props.apiBaseUrl()), e.getMessage());
            return false;
        } catch (Exception e) {
            log.warn("BB template import failed reason={} file={} error={}", reason, path, e.getMessage(), e);
            return true;
        }
    }

    private boolean apiAvailable(String reason) {
        try {
            restClient.get()
                .uri("/api/ping")
                .retrieve()
                .toBodilessEntity();
            return true;
        } catch (ResourceAccessException e) {
            log.warn("BB template import scan skipped reason={} apiBaseUrl={} error={}",
                reason, trimTrailingSlash(props.apiBaseUrl()), e.getMessage());
            return false;
        }
    }

    private boolean isImportWorkbook(Path path) {
        String name = path.getFileName().toString().toLowerCase();
        return name.endsWith(".xlsx")
            && !name.startsWith("~$")
            && !name.endsWith(".tmp.xlsx")
            && !name.endsWith(".partial.xlsx");
    }

    private boolean isStable(Fingerprint fp) {
        return Instant.now().minus(props.stableAge()).toEpochMilli() >= fp.lastModifiedMillis();
    }

    private Fingerprint fingerprint(Path path) throws IOException {
        return new Fingerprint(Files.size(path), Files.getLastModifiedTime(path).toMillis());
    }

    private static String trimTrailingSlash(String raw) {
        String value = java.util.Objects.requireNonNull(raw, "bb-template-import.api-base-url must be set").trim();
        while (value.endsWith("/")) value = value.substring(0, value.length() - 1);
        return value;
    }

    private record Fingerprint(long size, long lastModifiedMillis) {}
}
