package com.ubs.pesubjobs.model;

public record ProcessedLpMaster(
        String investorName,
        String parent,
        boolean spv,
        boolean highQty,
        String investorType,
        String instVsHnw,
        String regionLocation,
        boolean investmentGrade,
        String sp,
        String mdy,
        String fitch,
        String aum,
        String nav,
        String pension,
        String pensionFunded,
        String ubsClassification,
        String ubsDefaultAdvRate,
        String ubsDefaultConcLimit,
        String notes
) {}