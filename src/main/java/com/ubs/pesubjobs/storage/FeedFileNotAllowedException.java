package com.ubs.pesubjobs.storage;

/**
 * Thrown for a caller-supplied feed file name that is not an allowed, existing feed inside the
 * configured feed location (a local import root, or an Azure blob container/prefix).
 *
 * <p>Rejection messages name only what the caller supplied. The resolved identifier — a real
 * filesystem path, or a blob name — is server detail and is never returned to the caller.
 */
public class FeedFileNotAllowedException extends RuntimeException {
    public FeedFileNotAllowedException(String message) {
        super(message);
    }
}
