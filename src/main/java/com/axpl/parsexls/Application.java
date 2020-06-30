package com.axpl.parsexls;

import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
@EnableConfigurationProperties
@Slf4j
public class Application {

    private final ParserManager crawlerWorkersManager;

    @Autowired
    public Application(ParserManager crawlerWorkersManager) {
        this.crawlerWorkersManager = crawlerWorkersManager;
    }

    public static void main(String[] args) {
        log.info("ApplicationVersion: {}", "2.8.14");
//        SpringApplication.run(Application.class, args);
        new SpringApplicationBuilder(com.axpl.parsexls.Application.class)
                .web(false)
                .run(args);
    }

    @Bean
    CommandLineRunner init() {
        return args -> crawlerWorkersManager.run();
    }

}
