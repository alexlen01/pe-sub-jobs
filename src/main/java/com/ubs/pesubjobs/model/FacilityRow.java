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
        // Stated by the feed rather than worked out here: a group spans several facility rows, and
        // this job reads the feed in chunks, so it never sees enough of the file at once to notice.
        String umbrellaName,
        // What the group is resolved by — an account number where its members share one, the credit
        // agreement where they hold one each. Blank falls back to the account number on the API side.
        String umbrellaKey,
        // "true" where the group's members stand on ONE borrowing base, so the base each of them
        // carries is the same base. Blank where the feed cannot know, which is every group but the
        // sleeves of a multi-tranche facility.
        String umbrellaCrossCollateralized,
        // Which sleeve of a multi-tranche facility this row is, and the facility it is a sleeve OF.
        // Both blank on the ordinary facility. The feed states the relationship rather than the
        // platform recovering it from a "(Committed)"/"(Uncommitted)" suffix, which finds only the
        // sleeve names someone wrote a pattern for and misses a currency, term or accordion sleeve.
        String trancheType,
        String trancheOf,
        // The credit agreement this row's group is held under, where the agent prints a reference.
        // Last in the row on purpose: it is the newest column, and a feed written before it existed
        // still loads and still groups by its key exactly as it did.
        String agreementRef
) {}
