package com.ubs.pesubjobs.model;

public record LpMasterRow(
        String investorName,
        String parent,
        String spv,
        String highQuality,
        String investorType,
        String institutionalOrHnw,
        String regionLocation,
        String investmentGrade,
        String spRating,
        String moodysRating,
        String fitchRating,
        String aum,
        String nav,
        String pensionAssets,
        String fundingRatio,
        String ubsLpCategory,
        String ubsDefaultAdvanceRate,
        String ubsDefaultConcentrationLimit,
        String notes
) {}