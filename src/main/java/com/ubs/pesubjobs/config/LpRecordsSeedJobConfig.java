package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import com.ubs.pesubjobs.model.ProcessedLpFacilitySeed;
import com.ubs.pesubjobs.processor.LpFacilitySeedRowProcessor;
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
public class LpRecordsSeedJobConfig {

    @Bean
    public Job lpRecordsSeedJob(JobRepository jobRepository,
                                 @Qualifier("lpRecordsSeedStep") Step lpRecordsSeedStep) {
        return new JobBuilder("lpRecordsSeedJob", jobRepository)
                .start(lpRecordsSeedStep)
                .build();
    }

    @Bean("lpRecordsSeedStep")
    public Step lpRecordsSeedStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("lpFacilitySeedReader") FlatFileItemReader<LpFacilitySeedRow> reader,
                                   LpFacilitySeedRowProcessor processor,
                                   @Qualifier("lpRecordsSeedWriter") JdbcBatchItemWriter<ProcessedLpFacilitySeed> writer) {
        return new StepBuilder("lpRecordsSeedStep", jobRepository)
                .<LpFacilitySeedRow, ProcessedLpFacilitySeed>chunk(50, txManager)
                .reader(reader)
                .processor(processor)
                .writer(writer)
                .faultTolerant().skip(Exception.class).skipLimit(10)
                .build();
    }

    @Bean("lpFacilitySeedReader")
    @org.springframework.batch.core.configuration.annotation.StepScope
    public FlatFileItemReader<LpFacilitySeedRow> lpFacilitySeedReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<LpFacilitySeedRow>()
                .name("lpFacilitySeedReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .delimited().quoteCharacter('"')
                .names("facilityName", "investorName", "capCommit", "uncalled",
                       "agentCls", "agentRate", "agentConc")
                .fieldSetMapper(fs -> new LpFacilitySeedRow(
                        fs.readString("facilityName"),
                        fs.readString("investorName"),
                        fs.readString("capCommit"),
                        fs.readString("uncalled"),
                        fs.readString("agentCls"),
                        fs.readString("agentRate"),
                        fs.readString("agentConc")))
                .build();
    }

    @Bean("lpRecordsSeedWriter")
    public JdbcBatchItemWriter<ProcessedLpFacilitySeed> lpRecordsSeedWriter(DataSource dataSource) {
        String sql = """
                INSERT INTO lp_records (
                    facility_id, lp_master_id, source_seq, investor_name, parent,
                    spv, high_qty, investor_type, inst_vs_hnw, region_location,
                    investment_grade, classification, agent_cls,
                    sp, mdy, fitch, aum, nav, pension, pension_funded,
                    cap_commit, uncalled_capital, agent_rate, agent_conc,
                    included, rcl, transferee, created_at, updated_at
                ) VALUES (
                    :facilityId, :lpMasterId, :sourceSeq, :investorName, :parent,
                    :spv, :highQty, :investorType, :instVsHnw, :regionLocation,
                    :investmentGrade, :classification, :agentCls,
                    :sp, :mdy, :fitch, :aum, :nav, :pension, :pensionFunded,
                    :capCommit, :uncalled, :agentRate, :agentConc,
                    true, false, false, NOW(), NOW()
                )
                ON CONFLICT (facility_id, investor_name) DO NOTHING
                """;
        return new JdbcBatchItemWriterBuilder<ProcessedLpFacilitySeed>()
                .dataSource(dataSource)
                .sql(sql)
                .itemSqlParameterSourceProvider(r -> new MapSqlParameterSource()
                        .addValue("facilityId",      r.facilityId())
                        .addValue("lpMasterId",      r.lpMasterId())
                        .addValue("sourceSeq",       r.sourceSeq())
                        .addValue("investorName",    r.investorName())
                        .addValue("parent",          r.parent())
                        .addValue("spv",             r.spv())
                        .addValue("highQty",         r.highQty())
                        .addValue("investorType",    r.investorType())
                        .addValue("instVsHnw",       r.instVsHnw())
                        .addValue("regionLocation",  r.regionLocation())
                        .addValue("investmentGrade", r.investmentGrade())
                        .addValue("classification",  r.classification())
                        .addValue("agentCls",        r.agentCls())
                        .addValue("sp",              r.sp())
                        .addValue("mdy",             r.mdy())
                        .addValue("fitch",           r.fitch())
                        .addValue("aum",             r.aum())
                        .addValue("nav",             r.nav())
                        .addValue("pension",         r.pension())
                        .addValue("pensionFunded",   r.pensionFunded())
                        .addValue("capCommit",       r.capCommit())
                        .addValue("uncalled",        r.uncalled())
                        .addValue("agentRate",       r.agentRate())
                        .addValue("agentConc",       r.agentConc()))
                .build();
    }
}
