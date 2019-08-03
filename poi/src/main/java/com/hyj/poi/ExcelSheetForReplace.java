package com.hyj.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class ExcelSheetForReplace {
	private static Logger log = Logger.getLogger(ExcelSheetForReplace.class.getName());

	private static int COL_INDEX_KEY = 0;
	private static int COL_INDEX_VALUE = 1;

	FileInputStream fis;
	HSSFWorkbook wb;
	HSSFSheet sheet;

	public ExcelSheetForReplace(String path, String sheetName) throws IOException {
		fis = new FileInputStream(path);
		wb = new HSSFWorkbook(fis);
		sheet = wb.getSheet(sheetName);
	}

	public void close() throws IOException {
		wb.close();
		fis.close();
	}

	public HashMap<String, String> getReplaceMap(int rowNum) {
		HashMap<String, String> replaceMap = new HashMap<>();

		for(int row = 0; row < rowNum; row++) {
			String key = getSheetCellAsString(sheet, row, COL_INDEX_KEY);
			String val = getSheetCellAsString(sheet, row, COL_INDEX_VALUE);
			//log.info(String.format("#### %d: (%s, %s)", row, key, val));
			replaceMap.put(key, val);
		}

		return replaceMap;
	}

	private static String getSheetCellAsString(HSSFSheet sheet, int row, int col) {
		String ret = "";

		HSSFCell cell = sheet.getRow(row).getCell(col);
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
