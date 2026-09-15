package com.ubs.pesubjobs.model;

import java.math.BigDecimal;
import java.time.LocalDate;

/**
 * One umbrella feed row as the API takes it. Mirrors pe-sub-api's umbrella ingest payload: the whole
 * syndicated line is the group's {@code facilitySize} and UBS's slice of it the group's
 * {@code ubsParticipation}, since both are stated once over every member fund and no member borrows
 * either alone.
 */
public record ProcessedUmbrella(
        String key,
        String name,
        String agentBank,
        String accountNumber,
        BigDecimal facilitySize,
        // Null where the feed cannot know whether the members stand on one borrowing base, which
        // never turns an existing group's shared base off on the API side.
        Boolean crossCollateralized,
        // The agreement's own reference. Null where the feed states nothing, which the API reads as
        // silence rather than as a correction.
        String creditAgreementRef,
        // The agreement's own loan terms, which the group hands down to every member fund while it
        // is Active. Null where the feed states none, which never blanks a term an analyst recorded.
        LocalDate maturityDate,
        LocalDate collateralDate,
        // One of Pending, Active or Inactive as the feed states it — never "Not Stated", which is
        // what a group nothing has been said about already reads as. The API takes this on create
        // and to fill a blank, and never over a standing an analyst set.
        String facilityStatus,
        // UBS's slice of the agreement's line, refreshed each run alongside the size. Last in the
        // record to match the feed, which appends columns rather than inserting them.
        BigDecimal ubsParticipation
) {}
