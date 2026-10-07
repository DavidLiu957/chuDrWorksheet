package com.david.chuDrWorksheet;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import com.david.chuDrWorksheet.service.WorksheetTransformService;

@SpringBootTest
class ChuDrWorksheetApplicationTests {

	@Autowired
	private WorksheetTransformService transformService;

	@Test
	void testProcessExcel() {
		String sourceFile = "C:\\Users\\q966q\\eclipse-workspace-2025-12\\chuWorkingTableWorkspace\\chuData\\小闖的秘密寶典26年9月.xlsx";
		String templateFile = "C:\\Users\\q966q\\eclipse-workspace-2025-12\\chuWorkingTableWorkspace\\chuData\\醫生操作sample.xlsx";
		String outputFile = "C:\\Users\\q966q\\eclipse-workspace-2025-12\\chuWorkingTableWorkspace\\chuData\\醫生操作9月.xlsx";

		try {
			transformService.transformWorksheet(sourceFile, templateFile, outputFile);
			System.out.println("Excel 轉換完成！已輸出至：" + outputFile);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}