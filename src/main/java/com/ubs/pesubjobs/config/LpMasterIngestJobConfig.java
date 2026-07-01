package com.ubs.pesubjobs.config;

import com.ubs.pesubjobs.model.LpMasterRow;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import com.ubs.pesubjobs.processor.LpMasterRowProcessor;
import org.springframework.batch.core.Job;
import org.springframework.batch.core.Step;
import org.springframework.batch.core.configuration.annotation.StepScope;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.item.database.JdbcBatchItemWriter;
import org.springframework.batch.item.database.builder.JdbcBatchItemWriterBuilder;
import org.springframework.batch.item.file.FlatFileItemReader;
import org.springframework.batch.item.file.builder.FlatFileItemReaderBuilder;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.FileSystemResource;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.transaction.PlatformTransactionManager;

import javax.sql.DataSource;

@Configuration
public class LpMasterIngestJobConfig {

    @Bean
    public Job lpMasterIngestJob(JobRepository jobRepository,
                                 @Qualifier("lpMasterIngestStep") Step lpMasterIngestStep) {
        return new JobBuilder("lpMasterIngestJob", jobRepository)
                .start(lpMasterIngestStep)
                .build();
    }

    @Bean("lpMasterIngestStep")
    public Step lpMasterIngestStep(JobRepository jobRepository,
                                   PlatformTransactionManager txManager,
                                   @Qualifier("lpMasterReader") FlatFileItemReader<LpMasterRow> lpMasterReader,
                                   LpMasterRowProcessor lpMasterProcessor,
                                   @Qualifier("lpMasterWriter") JdbcBatchItemWriter<ProcessedLpMaster> lpMasterWriter) {
        return new StepBuilder("lpMasterIngestStep", jobRepository)
                .<LpMasterRow, ProcessedLpMaster>chunk(50, txManager)
                .reader(lpMasterReader)
                .processor(lpMasterProcessor)
                .writer(lpMasterWriter)
                .faultTolerant()
                .skip(Exception.class)
                .skipLimit(10)
                .build();
    }

    @Bean("lpMasterReader")
    @StepScope
    public FlatFileItemReader<LpMasterRow> lpMasterReader(
            @Value("#{jobParameters['filePath']}") String filePath) {
        return new FlatFileItemReaderBuilder<LpMasterRow>()
                .name("lpMasterReader")
                .resource(new FileSystemResource(filePath))
                .linesToSkip(1)
                .delimited()
                .quoteCharacter('"')
                .names("investorName", "parent", "spv", "highQty", "investorType",
                       "instVsHnw", "regionLocation", "investmentGrade", "sp", "mdy",
                       "fitch", "aum", "nav", "pension", "pensionFunded",
                       "ubsClassification", "ubsDefaultAdvRate", "ubsDefaultConcLimit", "notes")
                .fieldSetMapper(fs -> new LpMasterRow(
                        fs.readString("investorName"),
                        fs.readString("parent"),
                        fs.readString("spv"),
                        fs.readString("highQty"),
                        fs.readString("investorType"),
                        fs.readString("instVsHnw"),
                        fs.readString("regionLocation"),
                        fs.readString("investmentGrade"),
                        fs.readString("sp"),
                        fs.readString("mdy"),
                        fs.readString("fitch"),
                        fs.readString("aum"),
                        fs.readString("nav"),
                        fs.readString("pension"),
                        fs.readString("pensionFunded"),
                        fs.readString("ubsClassification"),
                        fs.readString("ubsDefaultAdvRate"),
                        fs.readString("ubsDefaultConcLimit"),
                        fs.readString("notes")
                ))
                .build();
    }

    @Bean
    public LpMasterRowProcessor lpMasterProcessor() {
        return new LpMasterRowProcessor();
    }

    @Bean("lpMasterWriter")
    public JdbcBatchItemWriter<ProcessedLpMaster> lpMasterWriter(DataSource dataSource) {
        String sql = """
                INSERT INTO lp_master (
                    investor_name, parent, spv, high_qty, investor_type, inst_vs_hnw,
                    region_location, investment_grade, sp, mdy, fitch, aum, nav, pension,
                    pension_funded, ubs_classification, ubs_default_adv_rate,
                    ubs_default_conc_limit, notes, created_at, updated_at
                )
                VALUES (
                    :investorName, :parent, :spv, :highQty, :investorType, :instVsHnw,
                    :regionLocation, :investmentGrade, :sp, :mdy, :fitch, :aum, :nav, :pension,
                    :pensionFunded, :ubsClassification, :ubsDefaultAdvRate,
                    :ubsDefaultConcLimit, :notes, NOW(), NOW()
                )
                ON CONFLICT (investor_name) DO UPDATE SET
                    parent                 = EXCLUDED.parent,
                    spv                    = EXCLUDED.spv,
                    high_qty               = EXCLUDED.high_qty,
                    investor_type          = EXCLUDED.investor_type,
                    inst_vs_hnw            = EXCLUDED.inst_vs_hnw,
                    region_location        = EXCLUDED.region_location,
                    investment_grade       = EXCLUDED.investment_grade,
                    sp                     = EXCLUDED.sp,
                    mdy                    = EXCLUDED.mdy,
                    fitch                  = EXCLUDED.fitch,
                    aum                    = EXCLUDED.aum,
                    nav                    = EXCLUDED.nav,
                    pension                = EXCLUDED.pension,
                    pension_funded         = EXCLUDED.pension_funded,
                    ubs_classification     = EXCLUDED.ubs_classification,
                    ubs_default_adv_rate   = EXCLUDED.ubs_default_adv_rate,
                    ubs_default_conc_limit = EXCLUDED.ubs_default_conc_limit,
                    notes                  = EXCLUDED.notes,
                    updated_at             = NOW()
                """;

        return new JdbcBatchItemWriterBuilder<ProcessedLpMaster>()
                .dataSource(dataSource)
                .sql(sql)
                .itemSqlParameterSourceProvider(item -> {
                    MapSqlParameterSource params = new MapSqlParameterSource();
                    params.addValue("investorName",        item.investorName());
                    params.addValue("parent",              item.parent());
                    params.addValue("spv",                 item.spv());
                    params.addValue("highQty",             item.highQty());
                    params.addValue("investorType",        item.investorType());
                    params.addValue("instVsHnw",           item.instVsHnw());
                    params.addValue("regionLocation",      item.regionLocation());
                    params.addValue("investmentGrade",     item.investmentGrade());
                    params.addValue("sp",                  item.sp());
                    params.addValue("mdy",                 item.mdy());
                    params.addValue("fitch",               item.fitch());
                    params.addValue("aum",                 item.aum());
                    params.addValue("nav",                 item.nav());
                    params.addValue("pension",             item.pension());
                    params.addValue("pensionFunded",       item.pensionFunded());
                    params.addValue("ubsClassification",   item.ubsClassification());
                    params.addValue("ubsDefaultAdvRate",   item.ubsDefaultAdvRate());
                    params.addValue("ubsDefaultConcLimit", item.ubsDefaultConcLimit());
                    params.addValue("notes",               item.notes());
                    return params;
                })
                .build();
    }
}