package com.ubs.pesubjobs.config;

import org.springframework.batch.infrastructure.item.file.transform.DelimitedLineTokenizer;

final class CsvLineTokenizers {

    private CsvLineTokenizers() {
    }

    static DelimitedLineTokenizer lenientQuotedCsvTokenizer(String... names) {
        DelimitedLineTokenizer tokenizer = new DelimitedLineTokenizer();
        tokenizer.setQuoteCharacter('"');
        tokenizer.setNames(names);
        tokenizer.setStrict(false);
        return tokenizer;
    }
}
