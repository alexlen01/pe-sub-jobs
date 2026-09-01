package com.ubs.pesubjobs.model;

import java.math.BigDecimal;
import java.time.LocalDate;

public record ProcessedFacility(
        String agentBank,
        String name,
        String accountNumber,
        BigDecimal loanAmount,
        LocalDate maturityDate,
        String bankStatus,
        LocalDate bankStatusDate,
        BigDecimal ubsParticipation,
        LocalDate collateralDate,
        String umbrellaName,
        String umbrellaKey,
        // Null where the feed cannot know whether the group's members stand on one borrowing base,
        // which never turns an existing group's shared base off on the API side.
        Boolean umbrellaCrossCollateralized
) {}
