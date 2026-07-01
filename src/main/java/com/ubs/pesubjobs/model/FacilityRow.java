package com.ubs.pesubjobs.model;

public record FacilityRow(
        String agentBank,
        String name,
        String accountNumber,
        String loanAmount,
        String maturityDate,
        String bankStatus,
        String bankStatusDate
) {}