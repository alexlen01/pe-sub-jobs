package com.ubs.pesubjobs.model;

/**
 * One lp_facility_seeds.csv row, posted verbatim to pe-sub-api's seed endpoint. Carries the
 * full per-LP column set of the LP DB Export (facility-level columns excluded) so the seed
 * matches the complete lp_records insert; component names mirror pe-sub-api's LpRecordSeedRow.
 * The legacy 7 fields come first — an old 7-column file still parses (missing columns blank).
 */
public record LpFacilitySeedRow(
        String facilityName,
        String investorName,
        String capCommit,
        String uncalled,
        String agentCls,
        String agentRate,
        String agentConc,
        String parent,
        String spv,
        String highQty,
        String investorType,
        String instVsHnw,
        String regionLocation,
        String investmentGrade,
        String ubsCls,
        String sp,
        String mdy,
        String fitch,
        String aum,
        String nav,
        String pension,
        String pensionFunded,
        String pctCapCommit,
        String calledCap,
        String pctUncalled,
        String pctCalled,
        String ubsConc,
        String ubsRate,
        String agentBb,
        String ubsBb,
        String notes) {}
