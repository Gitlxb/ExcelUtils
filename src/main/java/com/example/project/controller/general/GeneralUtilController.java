package com.example.project.controller.general;

import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

/**
 * General 模块公用端点 Controller
 */
@RestController
@RequestMapping("/api/general")
public class GeneralUtilController extends GeneralBaseController {

    @GetMapping("/download/{fileName}")
    public ResponseEntity<Resource> downloadFile(@PathVariable String fileName) {
        return doDownloadFile(fileName);
    }
}
