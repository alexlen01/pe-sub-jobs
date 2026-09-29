package com.ubs.pesubjobs.storage;

import com.azure.identity.DefaultAzureCredentialBuilder;
import com.azure.storage.blob.BlobClient;
import com.azure.storage.blob.BlobContainerClient;
import com.azure.storage.blob.BlobServiceClient;
import com.azure.storage.blob.BlobServiceClientBuilder;
import com.ubs.pesubjobs.config.AzureStorageProperties;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Profile;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.util.Locale;

/**
 * Reads feed CSVs from Azure Blob Storage — every profile except {@code local} (AKS). Auth is
 * Azure AD Workload Identity: {@link DefaultAzureCredentialBuilder} auto-detects the AKS
 * workload-identity webhook's injected environment (AZURE_CLIENT_ID, AZURE_TENANT_ID,
 * AZURE_FEDERATED_TOKEN_FILE). No connection string or account key is ever configured. Identifiers
 * are blob names (including the configured prefix).
 *
 * <p>Blob names have no filesystem traversal or symlink semantics, but {@link #resolve(String)}
 * still rejects anything that is not a bare file name, so a caller cannot address a blob outside
 * the configured container/prefix by naming one with separators of its own.
 *
 * <p>The Blob client is built lazily, on first actual use, rather than in the constructor: this
 * bean is created unconditionally during Spring context startup, and {@link
 * com.ubs.pesubjobs.JobStartupRunner} calls {@link #exists(String)} directly from its {@code
 * ApplicationRunner}, outside of any try/catch. A misconfigured account URL or a transient
 * credential/DNS problem must fail the one feed run or health check that hit it, not the whole
 * pod's startup.
 */
@Component
@Profile("!local")
public class AzureFeedFileStore implements FeedFileStore {

    private static final String REQUIRED_EXTENSION = ".csv";

    private final String accountUrl;
    private final String containerName;
    private final String prefix;
    private volatile BlobContainerClient containerClient;

    @Autowired
    public AzureFeedFileStore(AzureStorageProperties props) {
        this.accountUrl = props.accountUrl();
        this.containerName = props.feedContainer();
        this.prefix = AzureBlobNames.normalizePrefix(props.feedPrefix());
    }

    /** Test-only seam: exercise the resolve/exists/open/health logic against a mocked client. */
    AzureFeedFileStore(BlobContainerClient containerClient, String prefix) {
        this.accountUrl = null;
        this.containerName = null;
        this.prefix = prefix;
        this.containerClient = containerClient;
    }

    private BlobContainerClient client() {
        BlobContainerClient client = containerClient;
        if (client == null) {
            synchronized (this) {
                client = containerClient;
                if (client == null) {
                    BlobServiceClient serviceClient = new BlobServiceClientBuilder()
                            .endpoint(AzureBlobNames.requireAccountUrl(accountUrl))
                            .credential(new DefaultAzureCredentialBuilder().build())
                            .buildClient();
                    client = serviceClient.getBlobContainerClient(containerName);
                    containerClient = client;
                }
            }
        }
        return client;
    }

    @Override
    public boolean exists(String identifier) {
        if (identifier == null || identifier.isBlank()) {
            return false;
        }
        try {
            return client().getBlobClient(identifier).exists();
        } catch (Exception e) {
            return false;
        }
    }

    @Override
    public Resource open(String identifier) {
        BlobClient blob = client().getBlobClient(identifier);
        String fileName = AzureBlobNames.bareName(identifier);
        return new InputStreamResource(blob.openInputStream()) {
            @Override
            public String getFilename() {
                return fileName;
            }
        };
    }

    @Override
    public String resolve(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            throw new FeedFileNotAllowedException("A feed file name is required.");
        }
        String name = fileName.trim();

        if (name.contains("/") || name.contains("\\") || name.contains("..") || name.contains(":")) {
            throw new FeedFileNotAllowedException(
                    "Feed file must be a bare file name inside the configured import location: " + name);
        }
        if (!name.toLowerCase(Locale.ROOT).endsWith(REQUIRED_EXTENSION)) {
            throw new FeedFileNotAllowedException("Feed file must be a " + REQUIRED_EXTENSION + " file: " + name);
        }

        String blobName = prefix + name;
        if (!client().getBlobClient(blobName).exists()) {
            throw new FeedFileNotAllowedException("No readable feed file named " + name + " in the import location.");
        }
        return blobName;
    }

    @Override
    public boolean isReachable() {
        try {
            return client().exists();
        } catch (Exception e) {
            return false;
        }
    }

    @Override
    public String describeLocation() {
        try {
            return AzureBlobNames.describe(client().getBlobContainerUrl(), prefix);
        } catch (Exception e) {
            return "unreachable (" + e.getMessage() + ")";
        }
    }
}
