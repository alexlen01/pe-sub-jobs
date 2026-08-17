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
        String capitalCommitment,
        String uncalledCapital,
        String agentLpCategory,
        String agentAdvanceRate,
        String agentConcentrationLimit,
        String parent,
        String spv,
        String highQuality,
        String investorType,
        String institutionalOrHnw,
        String regionLocation,
        String investmentGrade,
        String ubsLpCategory,
        String spRating,
        String moodysRating,
        String fitchRating,
        String aum,
        String nav,
        String pensionAssets,
        String fundingRatio,
        String pctOfFundCommitments,
        String calledCapital,
        String pctOfFundUncalled,
        String pctLpCalled,
        String ubsConcentrationLimit,
        String ubsAdvanceRate,
        String agentBorrowingBase,
        String ubsBorrowingBase,
        String notes) {}
