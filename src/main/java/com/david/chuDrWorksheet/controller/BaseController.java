package com.david.chuDrWorksheet.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.david.chuDrWorksheet.service.WorksheetTransformService;

@RestController
public class BaseController {
	
	@Autowired
	private WorksheetTransformService transformService;

	@GetMapping("/api/process-excel")
	public String processExcel(@RequestParam(defaultValue = "C:/path/to/來源表.xlsx") String sourcePath,
			@RequestParam(defaultValue = "C:/path/to/sample.xlsx") String templatePath,
			@RequestParam(defaultValue = "C:/path/to/輸出結果.xlsx") String outputPath) {
		try {
			transformService.transformWorksheet(sourcePath, templatePath, outputPath);
			return "Excel 轉換成功！檔案已輸出至: " + outputPath;
		} catch (Exception e) {
			e.printStackTrace();
			return "轉換失敗: " + e.getMessage();
		}
	}
}
