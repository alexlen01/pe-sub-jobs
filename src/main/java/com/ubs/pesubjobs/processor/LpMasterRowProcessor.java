package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.LpMasterRow;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import org.springframework.batch.infrastructure.item.ItemProcessor;

public class LpMasterRowProcessor implements ItemProcessor<LpMasterRow, ProcessedLpMaster> {

    @Override
    public ProcessedLpMaster process(LpMasterRow item) {
        if (item.investorName() == null || item.investorName().isBlank()) return null;
        return new ProcessedLpMaster(
                item.investorName().trim(),
                blankToNull(item.parent()),
                parseBool(item.spv()),
                parseBool(item.highQuality()),
                blankToNull(item.investorType()),
                blankToNull(item.institutionalOrHnw()),
                blankToNull(item.regionLocation()),
                parseBool(item.investmentGrade()),
                defaultEmpty(item.spRating()),
                defaultEmpty(item.moodysRating()),
                defaultEmpty(item.fitchRating()),
                blankToNull(item.aum()),
                blankToNull(item.nav()),
                blankToNull(item.pensionAssets()),
                blankToNull(item.fundingRatio()),
                blankToNull(item.ubsLpCategory()),
                blankToNull(item.ubsDefaultAdvanceRate()),
                blankToNull(item.ubsDefaultConcentrationLimit()),
                blankToNull(item.notes())
        );
    }

    private String blankToNull(String s) {
        return (s == null || s.isBlank()) ? null : s.trim();
    }

    private String defaultEmpty(String s) {
        return (s == null || s.isBlank()) ? "" : s.trim();
    }

    private boolean parseBool(String s) {
        return "true".equalsIgnoreCase(s == null ? "" : s.trim());
    }
}
