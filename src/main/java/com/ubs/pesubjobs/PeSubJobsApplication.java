package com.ubs.pesubjobs;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.ConfigurationPropertiesScan;

@SpringBootApplication
@ConfigurationPropertiesScan
public class PeSubJobsApplication {

    public static void main(String[] args) {
        SpringApplication.run(PeSubJobsApplication.class, args);
    }
}