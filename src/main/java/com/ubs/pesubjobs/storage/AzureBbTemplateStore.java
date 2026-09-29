package com.ubs.pesubjobs.storage;

import com.azure.identity.DefaultAzureCredentialBuilder;
import com.azure.storage.blob.BlobClient;
import com.azure.storage.blob.BlobContainerClient;
import com.azure.storage.blob.BlobServiceClient;
import com.azure.storage.blob.BlobServiceClientBuilder;
import com.azure.storage.blob.models.BlobItem;
import com.azure.storage.blob.models.BlobItemProperties;
import com.azure.storage.blob.models.ListBlobsOptions;
import com.ubs.pesubjobs.config.AzureStorageProperties;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Profile;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * Watches an Azure Blob Storage container for finished BB template workbooks — every profile
 * except {@code local} (AKS). Single replica, no cross-pod locking: {@link
 * com.ubs.pesubjobs.BbTemplateDirectoryImporter} still dedupes with its own in-memory map, now
 * keyed by blob name instead of {@code Path}.
 *
 * <p>The Blob client is built lazily, on first actual use, rather than in the constructor: this
 * bean is created unconditionally during Spring context startup, and {@link
 * com.ubs.pesubjobs.BbTemplateDirectoryImporter} calls {@link #list()} directly from its {@code
 * ApplicationRunner}. A misconfigured account URL or a transient credential/DNS problem must fail
 * the one scan that hit it (caught by the importer as an {@link IOException}), not the whole pod's
 * startup.
 */
@Component
@Profile("!local")
public class AzureBbTemplateStore implements BbTemplateStore {

    private final String accountUrl;
    private final String containerName;
    private final String prefix;
    private volatile BlobContainerClient containerClient;

    @Autowired
    public AzureBbTemplateStore(AzureStorageProperties props) {
        this.accountUrl = props.accountUrl();
        this.containerName = props.bbTemplateContainer();
        this.prefix = AzureBlobNames.normalizePrefix(props.bbTemplatePrefix());
    }

    /** Test-only seam: exercise the list/open/health logic against a mocked client. */
    AzureBbTemplateStore(BlobContainerClient containerClient, String prefix) {
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
    public List<BbTemplateObject> list() throws IOException {
        try {
            ListBlobsOptions options = new ListBlobsOptions().setPrefix(prefix.isEmpty() ? null : prefix);
            List<BbTemplateObject> objects = new ArrayList<>();
            for (BlobItem item : client().listBlobs(options, null)) {
                String blobName = item.getName();
                BlobItemProperties properties = item.getProperties();
                long size = properties != null && properties.getContentLength() != null
                        ? properties.getContentLength() : 0L;
                Instant lastModified = properties != null && properties.getLastModified() != null
                        ? properties.getLastModified().toInstant() : Instant.EPOCH;
                objects.add(new BbTemplateObject(blobName, AzureBlobNames.bareName(blobName), size, lastModified));
            }
            objects.sort(Comparator.comparing(BbTemplateObject::name));
            return objects;
        } catch (Exception e) {
            // Broad on purpose: a bad account-url or credential-chain failure surfaces here (lazy
            // client construction happens on the listBlobs() call above), not just a BlobStorageException
            // from a reachable-but-erroring service. Either way the importer's ApplicationRunner must
            // see an IOException it already knows how to log and recover from, not a bare RuntimeException
            // that would crash Spring Boot's startup sequence.
            throw new IOException("Listing bb_templates blobs failed: " + e.getMessage(), e);
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
