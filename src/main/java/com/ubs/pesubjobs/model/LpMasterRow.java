package com.ubs.pesubjobs.model;

public record LpMasterRow(
        String investorName,
        String parent,
        String spv,
        String highQty,
        String investorType,
        String instVsHnw,
        String regionLocation,
        String investmentGrade,
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