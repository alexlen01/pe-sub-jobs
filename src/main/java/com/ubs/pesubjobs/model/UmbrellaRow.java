package com.ubs.pesubjobs.model;

/**
 * One raw row of the umbrella feed: the group layer above the facilities that share one credit
 * agreement.
 *
 * <p>A group is not a facility — it has no LP roster and no borrowing base of its own — so it is fed
 * on its own file rather than repeated on each member's row, where one agreement's terms would be
 * stated once per fund and whichever row loaded last would hold them.
 */
public record UmbrellaRow(
        // What the group is resolved by, and the only thing that survives an analyst renaming it:
        // the account number where the members share one, the credit agreement where they hold one
        // each.
        String key,
        // The name to onboard the group under, used on creation only.
        String name,
        String agentBank,
        // The number the group is administered under. Assigned by the PE Sub team, not issued by the
        // agent bank — the extract writes it through from the LP database, not from the agent's own
        // report.
        String accountNumber,
        // The whole agreement's line, as printed over the group.
        String loanAmount,
        // "true" where the members stand on ONE borrowing base. Blank where the feed cannot know,
        // which is every group but the sleeves of a multi-tranche facility.
        String crossCollateralized,
        // The credit agreement's own reference, where the agent prints one. What the API groups on
        // in preference to the account number, because an account is how a bank administers an
        // agreement and can be re-papered, whereas the reference is the agreement.
        String agreementRef
) {}
