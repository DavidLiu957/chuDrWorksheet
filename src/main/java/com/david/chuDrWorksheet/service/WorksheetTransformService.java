package com.david.chuDrWorksheet.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@Service
public class WorksheetTransformService {

	public void transformWorksheet(String sourceFilePath, String templateFilePath, String outputFilePath)
			throws IOException {
		try (Workbook srcWorkbook = new XSSFWorkbook(new FileInputStream(sourceFilePath));
				Workbook targetWorkbook = new XSSFWorkbook(new FileInputStream(templateFilePath))) {

			Sheet srcSheet = srcWorkbook.getSheetAt(0);
			Sheet targetSheet = targetWorkbook.getSheetAt(0);

			// 定義各醫師起始欄位索引 (0-based)
			Map<String, Integer> doctorColOffset = new HashMap<>();
			doctorColOffset.put("剡醫師", 0); // A ~ D
			doctorColOffset.put("江醫師", 4); // E ~ H
			doctorColOffset.put("陳醫師", 8); // I ~ L

			// 追蹤各醫師目前寫入到的 row index（Row 1 為醫師名稱、Row 2 為標題，資料從 Row 3 也就是 index 2 開始）
			Map<String, Integer> doctorNextRow = new HashMap<>();
			doctorNextRow.put("剡醫師", 2);
			doctorNextRow.put("江醫師", 2);
			doctorNextRow.put("陳醫師", 2);

			int lastRowNum = srcSheet.getLastRowNum();

			// 從第 4 列開始讀取 (Row index 3)
			for (int r = 3; r <= lastRowNum; r++) {
				Row srcRow = srcSheet.getRow(r);
				if (srcRow == null) {
					continue;
				}

				// 讀取醫師欄位 (Col 9)
				String doctor = String.format("%s醫師",getCellValueAsString(srcRow.getCell(19)).trim());
				if (doctor.isEmpty() || !doctorColOffset.containsKey(doctor)) {
					continue; // 醫師為空或不在指定名單內則略過
				}

				// 提取來源欄位資料
				String dateStr = getCellValueAsString(srcRow.getCell(10));
				String customerName = getCellValueAsString(srcRow.getCell(11));
				String amount = getCellValueAsString(srcRow.getCell(12));
				String totalCount = getCellValueAsString(srcRow.getCell(13));
				String usedCount = getCellValueAsString(srcRow.getCell(14));
				String treatment = getCellValueAsString(srcRow.getCell(16));
				String position = getCellValueAsString(srcRow.getCell(17));

				// 組裝客戶資訊: "客戶姓名 療程總次數-療程次數 購買金額"
				String customerInfo = String.format("%s %s-%s %s", customerName, totalCount, usedCount, amount).trim();
				System.out.println("customerInfo:" + customerInfo);
				// 取得當前醫師的寫入座標
				int baseCol = doctorColOffset.get(doctor);
				int targetRowIdx = doctorNextRow.get(doctor);

				Row targetRow = targetSheet.getRow(targetRowIdx);
				if (targetRow == null) {
					targetRow = targetSheet.createRow(targetRowIdx);
				}

				// 填入 4 個欄位
				targetRow.createCell(baseCol).setCellValue(dateStr); // 日期
				targetRow.createCell(baseCol + 1).setCellValue(treatment); // 項目
				targetRow.createCell(baseCol + 2).setCellValue(position); // 位置
				targetRow.createCell(baseCol + 3).setCellValue(customerInfo); // 客戶資訊

				// 推進該醫師的列號
				doctorNextRow.put(doctor, targetRowIdx + 1);
			}

			// 輸出結果檔案
			try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
				targetWorkbook.write(fos);
			}
		}
	}

	/**
	 * 安全提取儲存格文字內容，相容日期、數字及公式格式
	 */
	private String getCellValueAsString(Cell cell) {
		if (cell == null) {
			return "";
		}
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().trim();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				SimpleDateFormat sdf = new SimpleDateFormat("M月d日");
				return sdf.format(date);
			}
			double num = cell.getNumericCellValue();
			if (num == Math.floor(num)) {
				return String.valueOf((long) num);
			}
			return String.valueOf(num);
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case FORMULA:
			try {
				return cell.getStringCellValue();
			} catch (IllegalStateException e) {
				double evalNum = cell.getNumericCellValue();
				if (evalNum == Math.floor(evalNum)) {
					return String.valueOf((long) evalNum);
				}
				return String.valueOf(evalNum);
			}
		default:
			return "";
		}
	}
}