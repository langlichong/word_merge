package com.huhu.wordmerge.appose;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class WordMergeController {


    @Autowired
    private MergeWordService mergeWordService;

    @PostMapping("word/merge")
    public ResponseEntity<String> mergeDocs(){


        try {
            mergeWordService.mergeDocx();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return ResponseEntity.ok("merged successfully !");
    }
}
