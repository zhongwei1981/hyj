package com.hyj.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class ExcelOfData {
	private static Logger log = Logger.getLogger(ExcelOfData.class.getName());

	FileInputStream fis;
	HSSFWorkbook wb;

	public ExcelOfData(String path) throws IOException {
		fis = new FileInputStream(path);
		wb = new HSSFWorkbook(fis);
	}

	public void close() throws IOException {
		wb.close();
		fis.close();
	}

	public HashMap<String, String> getReplaceMap(HashMap<String, ReplaceData> replaceDataMap) {
		HashMap<String, String> replaceMap = new HashMap<>();

		replaceDataMap.forEach((k, v) -> {
			log.info(String.format("#### (%s, %s, %d, %d)", k, v.sheetName, v.row, v.col));

			String val = getSheetCellAsString(v.sheetName, v.row, v.col);
			replaceMap.put(k, val);
			log.info(String.format("#### (%s, %s) <- %s, %d, %d", k, val, v.sheetName, v.row, v.col));
		});

		return replaceMap;
	}

	private String getSheetCellAsString(String sheetName, int row, int col) {
		log.info(String.format("#### (%s, %d, %d)", sheetName, row, col));
		String ret = "";

		HSSFSheet sheet = wb.getSheet(sheetName);
		HSSFCell cell = sheet.getRow(row).getCell(col);
		CellType type = cell.getCellTypeEnum();
		switch (type) {
		case NUMERIC:
			//FIXME: value = double -> String -> double
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
