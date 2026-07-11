package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.model.FacilityRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import com.ubs.pesubjobs.processor.FacilityRowProcessor;
import org.springframework.batch.core.configuration.annotation.StepScope;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.Step;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.infrastructure.item.database.JdbcBatchItemWriter;
import org.springframework.batch.infrastructure.item.database.builder.JdbcBatchItemWriterBuilder;
import org.springframework.batch.infrastructure.item.file.FlatFileItemReader;
import org.springframework.batch.infrastructure.item.file.builder.FlatFileItemReaderBuilder;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.FileSystemResource;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.transaction.PlatformTransactionManager;

import javax.sql.DataSource;

@Configuration
public class FacilityIngestJobConfig {

    @Bean
    public Job facilityIngestJob(JobRepository jobRepository,
                                 @Qualifier("facilityIngestStep") Step facilityIngestStep) {
        return new JobBuilder("facilityIngestJob", jobRepository)
                .start(facilityIngestStep)
                .build();
    }

    @Bean("facilityIngestStep")
    public Step facilityIngestStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("facilityReader") FlatFileItemReader<FacilityRow> facilityReader,
                                   FacilityRowProcessor facilityProcessor,
                                   @Qualifier("facilityWriter") JdbcBatchItemWriter<ProcessedFacility> facilityWriter) {
        return new StepBuilder("facilityIngestStep", jobRepository)
                .<FacilityRow, ProcessedFacility>chunk(50)
                .transactionManager(txManager)
                .reader(facilityReader)
                .processor(facilityProcessor)
                .writer(facilityWriter)
                .faultTolerant()
                .skip(Exception.class)
                .skipLimit(10)
                .build();
    }

    @Bean("facilityReader")
    @StepScope
    public FlatFileItemReader<FacilityRow> facilityReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<FacilityRow>()
                .name("facilityReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .lineTokenizer(CsvLineTokenizers.lenientQuotedCsvTokenizer(
                        "agentBank", "name", "accountNumber", "loanAmount",
                        "maturityDate", "bankStatus", "bankStatusDate",
                        "ubsParticipation", "collateralDate"))
                .fieldSetMapper(fs -> new FacilityRow(
                        fs.readString("agentBank"),
                        fs.readString("name"),
                        fs.readString("accountNumber"),
                        fs.readString("loanAmount"),
                        fs.readString("maturityDate"),
                        fs.readString("bankStatus"),
                        fs.readString("bankStatusDate"),
                        fs.readString("ubsParticipation"),
                        fs.readString("collateralDate")
                ))
                .build();
    }

    @Bean
    public FacilityRowProcessor facilityProcessor() {
        return new FacilityRowProcessor();
    }

    @Bean("facilityWriter")
    public JdbcBatchItemWriter<ProcessedFacility> facilityWriter(DataSource dataSource) {
        String sql = """
                INSERT INTO facilities (
                    name, agent_bank, account_number, loan_amount, maturity_date,
                    bank_status, bank_status_date, ubs_participation, collateral_date,
                    status, conc_limit_m, created_at, updated_at
                )
                VALUES (
                    :name, :agentBank, :accountNumber, :loanAmount, :maturityDate,
                    :bankStatus, :bankStatusDate, :ubsParticipation, :collateralDate,
                    'Not Started', 25.00, NOW(), NOW()
                )
                ON CONFLICT (name) DO UPDATE SET
                    agent_bank        = EXCLUDED.agent_bank,
                    account_number    = EXCLUDED.account_number,
                    loan_amount       = EXCLUDED.loan_amount,
                    maturity_date     = EXCLUDED.maturity_date,
                    bank_status       = EXCLUDED.bank_status,
                    bank_status_date  = EXCLUDED.bank_status_date,
                    ubs_participation = EXCLUDED.ubs_participation,
                    collateral_date   = EXCLUDED.collateral_date,
                    updated_at        = NOW()
                """;

        return new JdbcBatchItemWriterBuilder<ProcessedFacility>()
                .dataSource(dataSource)
                .sql(sql)
                .itemSqlParameterSourceProvider(item -> {
                    MapSqlParameterSource params = new MapSqlParameterSource();
                    params.addValue("name",          item.name());
                    params.addValue("agentBank",      item.agentBank());
                    params.addValue("accountNumber",  item.accountNumber());
                    params.addValue("loanAmount",     item.loanAmount());
                    params.addValue("maturityDate",   item.maturityDate());
                    params.addValue("bankStatus",       item.bankStatus());
                    params.addValue("bankStatusDate",   item.bankStatusDate());
                    params.addValue("ubsParticipation", item.ubsParticipation());
                    params.addValue("collateralDate",   item.collateralDate());
                    return params;
                })
                .build();
    }
}
