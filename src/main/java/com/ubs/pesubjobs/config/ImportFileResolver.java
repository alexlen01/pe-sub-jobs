package com.ubs.pesubjobs.config;

import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;

/**
 * Resolves a caller-supplied feed file name to a path inside the configured import root.
 *
 * <p>The job trigger accepts a bare file name, never a path. Everything a caller could use to
 * leave the import root — separators, drive letters, {@code ..} segments, symbolic links — is
 * rejected before the file is opened, and the resolved real path is re-checked against the root
 * so a link planted inside the directory cannot redirect the read either.
 *
 * <p>Rejection messages name only what the caller supplied. The resolved absolute path is server
 * detail and is never returned.
 */
@Component
public class ImportFileResolver {

    /** Thrown for any name that is not an existing, readable CSV inside the import root. */
    public static class ImportFileNotAllowedException extends RuntimeException {
        public ImportFileNotAllowedException(String message) {
            super(message);
        }
    }

    private static final String REQUIRED_EXTENSION = ".csv";

    private final Path importRoot;

    public ImportFileResolver(IngestProperties props) {
        this.importRoot = Path.of(props.importRoot()).toAbsolutePath().normalize();
    }

    /**
     * @param fileName a bare file name such as {@code facilities.csv}
     * @return the absolute path to read
     * @throws ImportFileNotAllowedException if the name is not a readable CSV inside the root
     */
    public Path resolve(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            throw new ImportFileNotAllowedException("A feed file name is required.");
        }
        String name = fileName.trim();

        if (name.contains("/") || name.contains("\\") || name.contains("..") || name.contains(":")) {
            throw new ImportFileNotAllowedException(
                    "Feed file must be a bare file name inside the configured import directory: " + name);
        }
        if (!name.toLowerCase(java.util.Locale.ROOT).endsWith(REQUIRED_EXTENSION)) {
            throw new ImportFileNotAllowedException("Feed file must be a " + REQUIRED_EXTENSION + " file: " + name);
        }

        Path candidate;
        try {
            candidate = importRoot.resolve(name).normalize();
        } catch (InvalidPathException e) {
            throw new ImportFileNotAllowedException("Feed file name is not a valid file name: " + name);
        }
        // Belt-and-braces: normalize() alone would already have removed any traversal the checks
        // above missed, but an escape must fail closed rather than silently read elsewhere.
        if (!candidate.startsWith(importRoot)) {
            throw new ImportFileNotAllowedException(
                    "Feed file must be inside the configured import directory: " + name);
        }
        if (!Files.isRegularFile(candidate) || !Files.isReadable(candidate)) {
            throw new ImportFileNotAllowedException("No readable feed file named " + name + " in the import directory.");
        }
        // Resolve links last: a symlink placed inside the root must not widen the reachable set.
        try {
            Path real = candidate.toRealPath();
            if (!real.startsWith(importRoot)) {
                throw new ImportFileNotAllowedException(
                        "Feed file must be inside the configured import directory: " + name);
            }
            return real;
        } catch (IOException e) {
            throw new ImportFileNotAllowedException("Feed file could not be opened: " + name);
        }
    }

    /** The directory bare file names are resolved against. */
    public Path importRoot() {
        return importRoot;
    }
}
