package com.ubs.pesubjobs.storage;

import java.time.Instant;

/**
 * One workbook the BB-template watcher can currently see in the configured drop location.
 *
 * @param identifier what {@link BbTemplateStore#open(String)} needs to read it back — an absolute
 *                    local path, or a blob name
 * @param name        the bare file name, used for the workbook/lock-file/partial-file filtering
 *                     {@link com.ubs.pesubjobs.BbTemplateDirectoryImporter} applies
 * @param size         byte size, part of the change-fingerprint that decides whether the file has
 *                     finished being written and whether it was already imported
 * @param lastModified last-modified instant, the other half of that fingerprint
 */
public record BbTemplateObject(String identifier, String name, long size, Instant lastModified) {}
