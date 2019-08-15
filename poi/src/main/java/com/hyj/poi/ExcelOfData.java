package com.hyj.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOfData {
	private static Logger log = Logger.getLogger(ExcelOfData.class.getName());

	private static NumberFormat NF6 = NumberFormat.getNumberInstance();

	static {
		NF6.setMaximumFractionDigits(6);
		NF6.setRoundingMode(RoundingMode.UP);
	}

	private static final String EXCEL_SUFFIX = ".xls";
	private static final String EXCEL_X_SUFFIX = ".xlsx";

	FileInputStream fis;
	Workbook wb;

	public ExcelOfData(String path) throws IOException {
		fis = new FileInputStream(path);
		if (path.endsWith(EXCEL_SUFFIX)) {
			wb = new HSSFWorkbook(fis);
		} else if (path.endsWith(EXCEL_X_SUFFIX)) {
			wb = new XSSFWorkbook(fis);
		} else {
			throw new RuntimeException("Unsupported path = " + path);
		}
	}

	public void close() throws IOException {
		wb.close();
		fis.close();
	}

	public HashMap<String, String> getReplaceMap(HashMap<String, ReplaceData> replaceDataMap) {
		HashMap<String, String> replaceMap = new HashMap<>();

		replaceDataMap.forEach((k, v) -> {
			String val = getSheetCellAsString(v.sheetName, v.row, v.col);
			String formatVal = NF6.format(Double.parseDouble(val));
			replaceMap.put(k, formatVal);
			log.info(String.format("(%s, %s) <- (%s, %d, %d: %s)", k, formatVal, v.sheetName, v.row, v.col, val));
		});

		return replaceMap;
	}

	private String getSheetCellAsString(String sheetName, int row, int col) {
		//log.info(String.format("#### (%s, %d, %d)", sheetName, row, col));
		String ret = "";

		Sheet sheet = wb.getSheet(sheetName);
		Cell cell = sheet.getRow(row).getCell(col);
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
