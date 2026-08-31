package com.ubs.pesubjobs;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import com.ubs.pesubjobs.security.JobsSecurityProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
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
import java.util.Locale;
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

    // Two constructors, so the injection point has to be named explicitly.
    @Autowired
    public BbTemplateDirectoryImporter(BbTemplateImportProperties props, JobsSecurityProperties security) {
        this(props, security, RestClient.builder());
    }

    /**
     * Takes the builder rather than making one, so the scan's decisions — which files it posts, and
     * what it does when the API refuses one — can be exercised against a stubbed transport instead
     * of only against a live API. Not the injection point: the constructor above is.
     */
    BbTemplateDirectoryImporter(BbTemplateImportProperties props, JobsSecurityProperties security,
                                RestClient.Builder restClientBuilder) {
        this.props = props;
        // Template import is SERVICE-gated on pe-sub-api alongside the ANALYST screen path, so
        // this caller asserts only what it is. In gateway mode a header-less post is 401.
        this.restClient = restClientBuilder
                .baseUrl(trimTrailingSlash(props.apiBaseUrl()))
                .defaultHeader(security.getUserHeader(), security.getServiceUser())
                .defaultHeader(security.getRolesHeader(), security.getRequiredRole())
                .build();
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

    /**
     * Case folds against ROOT, not the JVM default: a Turkish default locale folds the "I" in
     * ".PARTIAL.XLSX" to a dotless "ı", the guard below stops matching, and a workbook still
     * being written gets posted to the API half-formed.
     */
    private boolean isImportWorkbook(Path path) {
        String name = path.getFileName().toString().toLowerCase(Locale.ROOT);
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
