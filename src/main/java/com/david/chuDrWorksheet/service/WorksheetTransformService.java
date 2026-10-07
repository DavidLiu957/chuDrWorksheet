package com.david.chuDrWorksheet.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class WorksheetTransformService {

	private static final int COL_DATE = 10; // 日期
	private static final int COL_CUSTOMER_NAME = 11; // 客戶姓名
	private static final int COL_AMOUNT = 12; // 購買金額
	private static final int COL_TOTAL_COUNT = 13; // 療程總次數
	private static final int COL_USED_COUNT = 14; // 療程使用次數
	private static final int COL_REMAINING_COUNT = 15; // 療程剩餘次數
	private static final int COL_TREATMENT = 16; // 治療項目
	private static final int COL_POSITION = 17; // 位置
	private static final int COL_DOCTOR = 19; // 醫師

	// 微整項目
	private static final String[] MINIMALLY_INVASIVE_COSMETIC_PROCEDURES = { "魔法", "肉毒", "玻尿酸", "奇蹟", "音波", "volbella",
			"voluma", "鳳凰", "埋線", "onda", "瑞斯朗" };

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
				String doctor = String.format("%s醫師", getCellValueAsString(srcRow.getCell(COL_DOCTOR)).trim());
				if (doctor.isEmpty() || !doctorColOffset.containsKey(doctor)) {
					continue; // 醫師為空或不在指定名單內則略過
				}

				// 提取來源欄位資料
				String dateStr = getCellValueAsString(srcRow.getCell(COL_DATE));
				String customerName = getCellValueAsString(srcRow.getCell(COL_CUSTOMER_NAME));
				String amountStr = getCellValueAsString(srcRow.getCell(COL_AMOUNT));
				String totalCountStr = getCellValueAsString(srcRow.getCell(COL_TOTAL_COUNT));
				String usedCountStr = parseMaxUsedCount(getCellValueAsString(srcRow.getCell(COL_USED_COUNT)));
				String remaining = getCellValueAsString(srcRow.getCell(COL_REMAINING_COUNT));
				String treatment = getCellValueAsString(srcRow.getCell(COL_TREATMENT));
				String position = getCellValueAsString(srcRow.getCell(COL_POSITION));

				// 檢查 是否是微整清單中的任何一項（忽略大小寫）
				boolean isMatched = false;
				if (treatment != null && !treatment.trim().isEmpty()) {
					String lowerTreatment = treatment.toLowerCase();
					for (String item : MINIMALLY_INVASIVE_COSMETIC_PROCEDURES) {
						if (lowerTreatment.contains(item.toLowerCase())) {
							isMatched = true;
							break;
						}
					}
				}
				// 是微整形 才執行 單次金額計算
				int positionUnits = 0;
				if (isMatched) {
					// 判斷位置欄位是否有數字並加總
					positionUnits = extractTotalUnitsFromPosition(position);
				}
				String finalAmount = amountStr;
				// 1. 優先檢查金額儲存格是否與下一列合併 則是專案
				if (isMergedWithNextRow(srcSheet, r, COL_AMOUNT)) {
					// 跨列合併時，直接標記為 "專案 金額"
					finalAmount = String.format("專案%s!!", amountStr);
				} else if (isMatched && positionUnits > 0) {
					// 是 微整 且非單次 則算單次價格
					try {
						double amount = Double.parseDouble(amountStr.replaceAll("[,\\s]", ""));
						double totalCount = Double.parseDouble(totalCountStr.replaceAll("[,\\s]", ""));

						if (totalCount > 0) {
							double calculatedAmount = (amount / totalCount) * positionUnits;
							// 四捨五入為整數
							finalAmount = String.valueOf(Math.round(calculatedAmount));
						}
					} catch (NumberFormatException e) {
						// 金額或次數非有效數值時，保留原始字串不中斷流程
						finalAmount = amountStr;
					}
				} else {
					// 否則直接寫金額
					finalAmount = amountStr;
				}
				// 2. 組裝客戶資訊: "客戶姓名 療程總次數-療程次數 購買金額"，若 totalCount 或 usedCount 為空，則省略中間的 次數-次數
				String customerInfo;
				boolean hasCounts = !totalCountStr.trim().isEmpty() && !usedCountStr.trim().isEmpty();

				if (isMatched && hasCounts) {
					customerInfo = String.format("%s %s-%s 使用%s 剩餘%s %s", customerName, totalCountStr.trim(),
							usedCountStr.trim(), positionUnits, remaining, finalAmount).trim();
				} else if (hasCounts) {
					// 兩者皆有值: "客戶姓名 總次數-使用次數 金額"
					customerInfo = String
							.format("%s %s-%s %s", customerName, totalCountStr.trim(), usedCountStr.trim(), finalAmount)
							.trim();
				} else {
					// 缺少次數: "客戶姓名 金額"
					customerInfo = String.format("%s %s", customerName, finalAmount).trim();
				}
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

	/**
	 * 從文字中提取所有數字並計算總和。 例如："抬頭12 魚尾12 皺眉12" -> 12 + 12 + 12 = 36 "voluma1 vobella1"
	 * -> 1 + 1 = 2 若無任何數字則回傳 0
	 */
	private int extractTotalUnitsFromPosition(String position) {
		if (position == null || position.trim().isEmpty()) {
			return 1;
		}
		Pattern pattern = Pattern.compile("\\d+");
		Matcher matcher = pattern.matcher(position);
		int totalUnits = 0;
		while (matcher.find()) {
			totalUnits += Integer.parseInt(matcher.group());
		}
		return totalUnits > 0 ? totalUnits : 1;
	}

	/**
	 * 檢查指定座標 (rowIdx, colIdx) 是否與下一列 (rowIdx + 1) 合併。
	 */
	private boolean isMergedWithNextRow(Sheet sheet, int rowIdx, int colIdx) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for (int i = 0; i < numMergedRegions; i++) {
			CellRangeAddress region = sheet.getMergedRegion(i);
			// 該儲存格落在合併範圍內，且該範圍跨到了下方列
			if (region.isInRange(rowIdx, colIdx)) {
				return region.getLastRow() > rowIdx;
			}
		}
		return false;
	}

	/**
	 * 處理療程使用次數： 若包含小數點或多個數字（如 "5.6"、"3.4"、"11.12"），分割並取出較大的數值回傳。 若無小數點或純整數，則原樣回傳。
	 */
	private String parseMaxUsedCount(String rawValue) {
		if (rawValue == null || rawValue.trim().isEmpty()) {
			return "";
		}

		String cleanVal = rawValue.trim();

		// 判斷是否含有小數點
		if (cleanVal.contains(".")) {
			// 以小數點為分隔符（正則需轉義 \\.）
			String[] parts = cleanVal.split("\\.");
			try {
				long maxVal = Long.MIN_VALUE;
				boolean hasValidNumber = false;

				for (String part : parts) {
					String trimmedPart = part.trim();
					if (!trimmedPart.isEmpty()) {
						long current = Long.parseLong(trimmedPart);
						if (current > maxVal) {
							maxVal = current;
							hasValidNumber = true;
						}
					}
				}

				if (hasValidNumber) {
					return String.valueOf(maxVal);
				}
			} catch (NumberFormatException e) {
				// 若包含非純數字字元（如 5.6次），保留原始字串
				return cleanVal;
			}
		}

		return cleanVal;
	}
}