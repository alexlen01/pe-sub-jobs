package com.ubs.pesubjobs.storage;

import com.azure.storage.blob.BlobClient;
import com.azure.storage.blob.BlobContainerClient;
import com.azure.storage.blob.specialized.BlobInputStream;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.Resource;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

/**
 * Exercises {@link AzureFeedFileStore}'s own logic — bare-name/extension validation, the blob
 * existence gate, and the streaming Resource it hands back to the batch readers — against a
 * mocked {@link BlobContainerClient}/{@link BlobClient}. No live Azure account is touched; the
 * real client-construction path in the public constructor (Workload Identity credential + SDK
 * builder) is exercised only via the {@code !local} profile in production/AKS.
 */
class AzureFeedFileStoreTest {

    private static final String PREFIX = "out/";

    private final BlobContainerClient containerClient = mock(BlobContainerClient.class);
    private final AzureFeedFileStore store = new AzureFeedFileStore(containerClient, PREFIX);

    @Test
    void exists_isDeferredToTheBlobClient() {
        BlobClient blobClient = mock(BlobClient.class);
        when(containerClient.getBlobClient("out/umbrellas.csv")).thenReturn(blobClient);
        when(blobClient.exists()).thenReturn(true);

        assertThat(store.exists("out/umbrellas.csv")).isTrue();
    }

    @Test
    void exists_isFalseForABlankOrMissingIdentifier_withoutCallingTheClient() {
        assertThat(store.exists("")).isFalse();
        assertThat(store.exists("   ")).isFalse();
        assertThat(store.exists(null)).isFalse();

        verify(containerClient, never()).getBlobClient(any());
    }

    @Test
    void resolve_rejectsNamesWithPathSeparatorsOrTraversal() {
        assertThatThrownBy(() -> store.resolve("../secrets.csv"))
                .isInstanceOf(FeedFileNotAllowedException.class);
        assertThatThrownBy(() -> store.resolve("sub/umbrellas.csv"))
                .isInstanceOf(FeedFileNotAllowedException.class);
        assertThatThrownBy(() -> store.resolve("C:\\umbrellas.csv"))
                .isInstanceOf(FeedFileNotAllowedException.class);

        verify(containerClient, never()).getBlobClient(any());
    }

    @Test
    void resolve_rejectsNonCsvNames() {
        assertThatThrownBy(() -> store.resolve("umbrellas.xlsx"))
                .isInstanceOf(FeedFileNotAllowedException.class);
    }

    @Test
    void resolve_rejectsBlankOrMissingNames() {
        assertThatThrownBy(() -> store.resolve(" "))
                .isInstanceOf(FeedFileNotAllowedException.class);
        assertThatThrownBy(() -> store.resolve(null))
                .isInstanceOf(FeedFileNotAllowedException.class);
    }

    @Test
    void resolve_rejectsAFileNameThatIsNotAnExistingBlob() {
        BlobClient blobClient = mock(BlobClient.class);
        when(containerClient.getBlobClient("out/umbrellas.csv")).thenReturn(blobClient);
        when(blobClient.exists()).thenReturn(false);

        assertThatThrownBy(() -> store.resolve("umbrellas.csv"))
                .isInstanceOf(FeedFileNotAllowedException.class);
    }

    @Test
    void resolve_returnsThePrefixedBlobNameForAnExistingCsv() {
        BlobClient blobClient = mock(BlobClient.class);
        when(containerClient.getBlobClient("out/umbrellas.csv")).thenReturn(blobClient);
        when(blobClient.exists()).thenReturn(true);

        assertThat(store.resolve("umbrellas.csv")).isEqualTo("out/umbrellas.csv");
    }

    @Test
    void open_streamsFromTheBlobAndReportsTheBareFileName() throws IOException {
        BlobClient blobClient = mock(BlobClient.class);
        BlobInputStream blobInputStream = mock(BlobInputStream.class);
        when(containerClient.getBlobClient("out/umbrellas.csv")).thenReturn(blobClient);
        when(blobClient.openInputStream()).thenReturn(blobInputStream);

        Resource resource = store.open("out/umbrellas.csv");

        // The FlatFileItemReader only needs a filename for logging/diagnostics; InputStreamResource
        // returns null unless overridden, so this pins the override AzureFeedFileStore relies on.
        assertThat(resource.getFilename()).isEqualTo("umbrellas.csv");
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
        when(containerClient.getBlobContainerUrl()).thenReturn("https://acct.blob.core.windows.net/pe-sub-feeds");

        assertThat(store.describeLocation()).isEqualTo("https://acct.blob.core.windows.net/pe-sub-feeds/out/");
    }

    @Test
    void describeLocation_isJustTheContainerUrlWhenThereIsNoPrefix() {
        AzureFeedFileStore noPrefixStore = new AzureFeedFileStore(containerClient, "");
        when(containerClient.getBlobContainerUrl()).thenReturn("https://acct.blob.core.windows.net/pe-sub-feeds");

        assertThat(noPrefixStore.describeLocation()).isEqualTo("https://acct.blob.core.windows.net/pe-sub-feeds");
    }
}
