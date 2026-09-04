package com.ubs.pesubjobs.model;

import java.math.BigDecimal;

/**
 * One umbrella feed row as the API takes it. Mirrors pe-sub-api's umbrella ingest payload: the
 * agreement's line is the group's {@code facilitySize}, since it is stated once over every member
 * fund and no member borrows it alone.
 */
public record ProcessedUmbrella(
        String key,
        String name,
        String obligorName,
        String agentBank,
        String accountNumber,
        BigDecimal facilitySize,
        // Null where the feed cannot know whether the members stand on one borrowing base, which
        // never turns an existing group's shared base off on the API side.
        Boolean crossCollateralized,
        // The agreement's own reference and its recorded terms. Null where the feed states nothing,
        // which the API reads as silence rather than as a correction.
        String agreementRef,
        String borrowerEntity,
        BigDecimal subLimit,
        String liabilityType
) {}
