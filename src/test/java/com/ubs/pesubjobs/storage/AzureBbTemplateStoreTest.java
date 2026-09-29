package com.ubs.pesubjobs.storage;

import com.azure.core.http.rest.PagedIterable;
import com.azure.storage.blob.BlobClient;
import com.azure.storage.blob.BlobContainerClient;
import com.azure.storage.blob.models.BlobItem;
import com.azure.storage.blob.models.BlobItemProperties;
import com.azure.storage.blob.models.BlobStorageException;
import com.azure.storage.blob.models.ListBlobsOptions;
import com.azure.storage.blob.specialized.BlobInputStream;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.Resource;

import java.io.IOException;
import java.time.Instant;
import java.time.OffsetDateTime;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.isNull;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

/**
 * Exercises {@link AzureBbTemplateStore}'s own logic — turning a page of {@link BlobItem}s into
 * the poller's {@link BbTemplateObject} contract, error translation, and the streaming Resource
 * used for the multipart upload to pe-sub-api — against a mocked {@link BlobContainerClient}. No
 * live Azure account is touched; the real client-construction path in the public constructor
 * (Workload Identity credential + SDK builder) is exercised only via the {@code !local} profile in
 * production/AKS.
 */
class AzureBbTemplateStoreTest {

    private static final String PREFIX = "AgentBBs/bb_templates/";

    private final BlobContainerClient containerClient = mock(BlobContainerClient.class);
    private final AzureBbTemplateStore store = new AzureBbTemplateStore(containerClient, PREFIX);

    @Test
    void list_returnsObjectsSortedByBareNameWithSizeAndLastModified() throws IOException {
        BlobItem zeta = blobItem(PREFIX + "zeta.xlsx", 200L, OffsetDateTime.parse("2024-02-01T00:00:00Z"));
        BlobItem alpha = blobItem(PREFIX + "alpha.xlsx", 100L, OffsetDateTime.parse("2024-01-01T00:00:00Z"));
        stubListBlobs(zeta, alpha);

        List<BbTemplateObject> result = store.list();

        assertThat(result).extracting(object -> object.name()).containsExactly("alpha.xlsx", "zeta.xlsx");
        BbTemplateObject first = result.get(0);
        assertThat(first.identifier()).isEqualTo(PREFIX + "alpha.xlsx");
        assertThat(first.size()).isEqualTo(100L);
        assertThat(first.lastModified()).isEqualTo(OffsetDateTime.parse("2024-01-01T00:00:00Z").toInstant());
    }

    @Test
    void list_defaultsMissingSizeAndLastModifiedRatherThanFailing() throws IOException {
        // A properties block with no size/last-modified set — e.g. an SDK model that only
        // partially populated the fields the REST listing normally guarantees.
        BlobItem sparseProperties = new BlobItem()
                .setName(PREFIX + "bare.xlsx")
                .setProperties(new BlobItemProperties());
        stubListBlobs(sparseProperties);

        List<BbTemplateObject> result = store.list();

        assertThat(result).hasSize(1);
        assertThat(result.get(0).size()).isZero();
        assertThat(result.get(0).lastModified()).isEqualTo(Instant.EPOCH);
    }

    @Test
    void list_wrapsAFailedListingAsAnIOExceptionRatherThanTheSdkException() {
        when(containerClient.listBlobs(any(ListBlobsOptions.class), isNull()))
                .thenThrow(mock(BlobStorageException.class));

        assertThatThrownBy(() -> store.list()).isInstanceOf(IOException.class);
    }

    @Test
    void open_streamsFromTheBlobAndReportsTheBareFileNameForTheUploadPart() throws IOException {
        BlobClient blobClient = mock(BlobClient.class);
        BlobInputStream blobInputStream = mock(BlobInputStream.class);
        when(containerClient.getBlobClient(PREFIX + "citibank.xlsx")).thenReturn(blobClient);
        when(blobClient.openInputStream()).thenReturn(blobInputStream);

        Resource resource = store.open(PREFIX + "citibank.xlsx");

        // BbTemplateDirectoryImporter posts this Resource as a multipart file part; the part needs
        // a filename, which InputStreamResource returns as null unless overridden.
        assertThat(resource.getFilename()).isEqualTo("citibank.xlsx");
        assertThat(resource.getInputStream()).isSameAs(blobInputStream);
    }

    @Test
    void isReachable_reflectsContainerExistence() {
        when(containerClient.exists()).thenReturn(true);
        assertThat(store.isReachable()).isTrue();

        when(containerClient.exists()).thenReturn(false);
        assertThat(store.isReachable()).isFalse();
    }

    @Test
    void isReachable_isFalseWhenTheContainerCheckThrows() {
        when(containerClient.exists()).thenThrow(new RuntimeException("network blip"));

        assertThat(store.isReachable()).isFalse();
    }

    @Test
    void describeLocation_combinesTheContainerUrlAndThePrefix() {
        when(containerClient.getBlobContainerUrl())
                .thenReturn("https://acct.blob.core.windows.net/pe-sub-bb-templates");

        assertThat(store.describeLocation())
                .isEqualTo("https://acct.blob.core.windows.net/pe-sub-bb-templates/" + PREFIX);
    }

    @SuppressWarnings("unchecked")
    private void stubListBlobs(BlobItem... items) {
        PagedIterable<BlobItem> pagedIterable = mock(PagedIterable.class);
        when(pagedIterable.iterator()).thenReturn(List.of(items).iterator());
        when(containerClient.listBlobs(any(ListBlobsOptions.class), isNull())).thenReturn(pagedIterable);
    }

    private static BlobItem blobItem(String name, long size, OffsetDateTime lastModified) {
        return new BlobItem()
                .setName(name)
                .setProperties(new BlobItemProperties().setContentLength(size).setLastModified(lastModified));
    }
}
