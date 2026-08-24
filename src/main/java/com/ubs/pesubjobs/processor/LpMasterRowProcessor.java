package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.LpMasterRow;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import org.springframework.batch.infrastructure.item.ItemProcessor;

public class LpMasterRowProcessor implements ItemProcessor<LpMasterRow, ProcessedLpMaster> {

    @Override
    public ProcessedLpMaster process(LpMasterRow item) {
        // A nameless row is not filtered out here: the API rejects it and the job fails, rather
        // than the row leaving the load unnoticed.
        return new ProcessedLpMaster(
                blankToNull(item.investorName()),
                blankToNull(item.parent()),
                parseBool(item.spv()),
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
