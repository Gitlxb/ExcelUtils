package com.example.project;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ProjectApplication {

    public static void main(String[] args) {
        // 全局设置Zip压缩比限制，避免Apache POI的"Zip bomb detected"错误
        ZipSecureFile.setMinInflateRatio(0.001);
        SpringApplication.run(ProjectApplication.class, args);
    }

}
