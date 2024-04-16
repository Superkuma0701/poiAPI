package com.git.poitest;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PoiTestApplication {
//    private  static  final Logger logger = LoggerFactory.getLogger(PoiTestApplication.class);
    public static void main(String[] args) {
        SpringApplication.run(PoiTestApplication.class, args);
    }


//    @Bean
//    public CommandLineRunner commandLineRunner(){
//        return run -> {
//            logger.debug("this is debug");
//            logger.info("this is info");
//            logger.warn("this is warn");
//            logger.error("this is error");
//        };
//    }
}
