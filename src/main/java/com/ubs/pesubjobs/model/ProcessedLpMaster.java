package com.ubs.pesubjobs.model;

public record ProcessedLpMaster(
        String investorName,
        String parent,
        boolean spv,
        boolean highQuality,
        String investorType,
        String institutionalOrHnw,
        String regionLocation,
        boolean investmentGrade,
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