package com.ubs.pesubjobs.model;

public record FacilityRow(
        String agentBank,
        String name,
        String accountNumber,
        String loanAmount,
        String maturityDate,
        String bankStatus,
        String bankStatusDate,
        String ubsParticipation,
        String collateralDate,
        // The umbrella subscription facility this fund borrows under, blank when it borrows alone.
        // Stated by the feed rather than worked out here: an umbrella is an account carrying more
        // than one facility, and this job reads the feed in chunks, so it never sees enough of the
        // file at once to notice.
        String umbrellaName
) {}
