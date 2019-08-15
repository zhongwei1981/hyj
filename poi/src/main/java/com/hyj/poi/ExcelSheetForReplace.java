package com.hyj.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheetForReplace {
	private static Logger log = Logger.getLogger(ExcelSheetForReplace.class.getName());

	private static int COL_INDEX_KEY = 0;
	private static int COL_INDEX_REPLACE_SHEET = 1;
	private static int COL_INDEX_REPLACE_ROW_COL = 2;

	FileInputStream fis;
	Workbook wb;
	Sheet sheet;

	public ExcelSheetForReplace(String path, String sheetName) throws IOException {
		fis = new FileInputStream(path);
		if (path.endsWith(Common.EXCEL_SUFFIX)) {
			wb = new HSSFWorkbook(fis);
		} else if (path.endsWith(Common.EXCEL_X_SUFFIX)) {
			wb = new XSSFWorkbook(fis);
		} else {
			throw new RuntimeException("Unsupported path = " + path);
		}
		sheet = wb.getSheet(sheetName);
	}

	public void close() throws IOException {
		wb.close();
		fis.close();
	}

	public HashMap<String, ReplaceData> getReplaceDataMap(int rowNum) {
		HashMap<String, ReplaceData> replaceDataMap = new HashMap<>();

		for(int row = 0; row < rowNum; row++) {
			String key = getSheetCellAsString(sheet, row, COL_INDEX_KEY);
			String replaceSheet = getSheetCellAsString(sheet, row, COL_INDEX_REPLACE_SHEET);
			String replaceRowCol = getSheetCellAsString(sheet, row, COL_INDEX_REPLACE_ROW_COL);

			// verify key is wrapped by specified string
			if (key.startsWith(Common.START_KEY) && key.endsWith(Common.END_KEY)) {
				key = key.replaceAll(Common.START_KEY, "");
				key = key.replaceAll(Common.END_KEY, "");
			} else {
				throw new RuntimeException(String.format(
						"Unexpected key = %s, which is expected to be wrapped by (%s & %s)",
						key, Common.START_KEY, Common.END_KEY));
			}
			ReplaceData data = new ReplaceData(replaceSheet, replaceRowCol);

			replaceDataMap.put(key, data);
			log.info(String.format("(%s <- %s)", key, data.getContent()));
		}

		return replaceDataMap;
	}

	private static String getSheetCellAsString(Sheet sheet, int row, int col) {
		//log.info(String.format("#### (%s, %d, %d)", "sheet", row, col));
		String ret = "";

		Cell cell = sheet.getRow(row).getCell(col);
		CellType type = cell.getCellTypeEnum();
		switch (type) {
		case NUMERIC:
			ret = Double.toString(cell.getNumericCellValue());
			break;
		case STRING:
			ret = cell.getStringCellValue();
			break;
		default:
			log.error("#### not supported type=" + type);
		}

		return ret;
	}
}
