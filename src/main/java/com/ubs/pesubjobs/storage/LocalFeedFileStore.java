package com.ubs.pesubjobs.storage;

import com.ubs.pesubjobs.config.IngestProperties;
import org.springframework.context.annotation.Profile;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.Locale;

/**
 * Reads feed CSVs from the local filesystem — the {@code local} profile's docker-compose / developer
 * workflow. Identifiers are absolute, real (symlink-resolved) filesystem paths as strings.
 *
 * <p>{@link #resolve(String)} carries the traversal/symlink defenses the pre-Azure-migration
 * {@code ImportFileResolver} used to own directly: separators, drive letters, {@code ..} segments
 * and symbolic links are all rejected before a file is opened, and the resolved real path is
 * re-checked against the import root so a link planted inside the directory cannot redirect the
 * read either.
 */
@Component
@Profile("local")
public class LocalFeedFileStore implements FeedFileStore {

    private static final String REQUIRED_EXTENSION = ".csv";

    private final Path importRoot;

    public LocalFeedFileStore(IngestProperties props) {
        this.importRoot = Path.of(props.importRoot()).toAbsolutePath().normalize();
    }

    @Override
    public boolean exists(String identifier) {
        if (identifier == null || identifier.isBlank()) {
            return false;
        }
        Path path = Path.of(identifier);
        return Files.isRegularFile(path) && Files.isReadable(path);
    }

    @Override
    public Resource open(String identifier) {
        return new FileSystemResource(identifier);
    }

    @Override
    public String resolve(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            throw new FeedFileNotAllowedException("A feed file name is required.");
        }
        String name = fileName.trim();

        if (name.contains("/") || name.contains("\\") || name.contains("..") || name.contains(":")) {
            throw new FeedFileNotAllowedException(
                    "Feed file must be a bare file name inside the configured import directory: " + name);
        }
        if (!name.toLowerCase(Locale.ROOT).endsWith(REQUIRED_EXTENSION)) {
            throw new FeedFileNotAllowedException("Feed file must be a " + REQUIRED_EXTENSION + " file: " + name);
        }

        Path candidate;
        try {
            candidate = importRoot.resolve(name).normalize();
        } catch (InvalidPathException e) {
            throw new FeedFileNotAllowedException("Feed file name is not a valid file name: " + name);
        }
        // Belt-and-braces: normalize() alone would already have removed any traversal the checks
        // above missed, but an escape must fail closed rather than silently read elsewhere.
        if (!candidate.startsWith(importRoot)) {
            throw new FeedFileNotAllowedException(
                    "Feed file must be inside the configured import directory: " + name);
        }
        if (!Files.isRegularFile(candidate) || !Files.isReadable(candidate)) {
            throw new FeedFileNotAllowedException("No readable feed file named " + name + " in the import directory.");
        }
        // Resolve links last: a symlink placed inside the root must not widen the reachable set.
        try {
            Path real = candidate.toRealPath();
            if (!real.startsWith(importRoot)) {
                throw new FeedFileNotAllowedException(
                        "Feed file must be inside the configured import directory: " + name);
            }
            return real.toString();
        } catch (IOException e) {
            throw new FeedFileNotAllowedException("Feed file could not be opened: " + name);
        }
    }

    @Override
    public boolean isReachable() {
        return Files.isDirectory(importRoot) && Files.isReadable(importRoot);
    }

    @Override
    public String describeLocation() {
        return importRoot.toString();
    }
}
