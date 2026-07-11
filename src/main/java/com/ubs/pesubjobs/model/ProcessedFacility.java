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
        LocalDate collateralDate
) {}
