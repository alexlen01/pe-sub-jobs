package com.ubs.pesubjobs.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.DefaultValue;

/**
 * Where feed CSVs and bb_templates workbooks live in Azure Blob Storage, for every profile except
 * {@code local} (see {@code com.ubs.pesubjobs.storage}). Auth is Azure AD Workload Identity via
 * {@code DefaultAzureCredentialBuilder} — deliberately no connection string or account key
 * property here, only the account URL the client connects to.
 */
@ConfigurationProperties(prefix = "azure.storage")
public record AzureStorageProperties(
        // e.g. https://<account>.blob.core.windows.net — required outside the local profile.
        String accountUrl,
        @DefaultValue("pe-sub-feeds") String feedContainer,
        // Optional path prefix inside the feed container, e.g. "out/". Blank means the container root.
        @DefaultValue("") String feedPrefix,
        @DefaultValue("pe-sub-bb-templates") String bbTemplateContainer,
        // Optional path prefix inside the bb_templates container, e.g. "AgentBBs/bb_templates/".
        @DefaultValue("") String bbTemplatePrefix) {}
