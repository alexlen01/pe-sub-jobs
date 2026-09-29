package com.ubs.pesubjobs.storage;

/** Small shared helpers for the two Azure-backed stores — not part of the public storage contract. */
final class AzureBlobNames {

    private AzureBlobNames() {}

    /** Normalizes a configured prefix to either "" or a value ending in exactly one "/". */
    static String normalizePrefix(String prefix) {
        if (prefix == null || prefix.isBlank()) {
            return "";
        }
        String value = prefix.trim();
        while (value.startsWith("/")) {
            value = value.substring(1);
        }
        if (!value.isEmpty() && !value.endsWith("/")) {
            value = value + "/";
        }
        return value;
    }

    /** The bare file name at the end of a blob name, for filtering and multipart upload metadata. */
    static String bareName(String blobName) {
        int idx = blobName.lastIndexOf('/');
        return idx >= 0 ? blobName.substring(idx + 1) : blobName;
    }

    static String describe(String containerUrl, String prefix) {
        return prefix.isEmpty() ? containerUrl : containerUrl + "/" + prefix;
    }

    /**
     * Fails fast with an actionable message rather than letting the Azure SDK reject a blank or
     * malformed endpoint deep inside client construction — a misconfigured
     * {@code azure.storage.account-url} outside the {@code local} profile must stop the pod at
     * startup, not surface as an opaque connectivity error on the first feed run.
     */
    static String requireAccountUrl(String accountUrl) {
        if (accountUrl == null || accountUrl.isBlank()) {
            throw new IllegalStateException(
                    "azure.storage.account-url must be set outside the 'local' profile "
                    + "(e.g. https://<account>.blob.core.windows.net)");
        }
        return accountUrl;
    }
}
