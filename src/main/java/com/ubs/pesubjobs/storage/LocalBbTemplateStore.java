package com.ubs.pesubjobs.storage;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import org.springframework.context.annotation.Profile;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;

/**
 * Watches a local directory for finished BB template workbooks — the {@code local} profile's
 * docker-compose / developer workflow. Identifiers are absolute filesystem paths as strings.
 */
@Component
@Profile("local")
public class LocalBbTemplateStore implements BbTemplateStore {

    private final Path directory;

    public LocalBbTemplateStore(BbTemplateImportProperties props) {
        this.directory = Path.of(props.directory()).toAbsolutePath().normalize();
    }

    @Override
    public List<BbTemplateObject> list() throws IOException {
        Files.createDirectories(directory);
        List<Path> paths;
        try (Stream<Path> files = Files.list(directory)) {
            paths = files
                    .filter(Files::isRegularFile)
                    .sorted(Comparator.comparing(path -> path.getFileName().toString()))
                    .toList();
        }
        List<BbTemplateObject> objects = new ArrayList<>(paths.size());
        for (Path path : paths) {
            objects.add(new BbTemplateObject(
                    path.toString(),
                    path.getFileName().toString(),
                    Files.size(path),
                    Files.getLastModifiedTime(path).toInstant()));
        }
        return objects;
    }

    @Override
    public Resource open(String identifier) {
        return new FileSystemResource(identifier);
    }

    @Override
    public boolean isReachable() {
        return Files.isDirectory(directory) && Files.isReadable(directory);
    }

    @Override
    public String describeLocation() {
        return directory.toString();
    }
}
