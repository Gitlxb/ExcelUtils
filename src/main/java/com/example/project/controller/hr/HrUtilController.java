package com.example.project.controller.hr;

import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

/**
 * HR 模块公用端点 Controller
 */
@RestController
@RequestMapping("/api/hr")
public class HrUtilController extends HrBaseController {

    @GetMapping("/download/{fileName}")
    public ResponseEntity<Resource> downloadFile(@PathVariable String fileName) {
        return doDownloadFile(fileName);
    }
}
