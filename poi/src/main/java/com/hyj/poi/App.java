package com.hyj.poi;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

import org.apache.log4j.Logger;

public class App {
	private static Logger log = Logger.getLogger(App.class.getName());

	private static final String WORD_SUFFIX = ".doc";
	private static final String WORD_X_SUFFIX = ".docx";

	private static final int ARG_WORK_DIR_INDEX = 0;
	private static final int ARG_WORD_TEMPLATE_INDEX = 1;
	private static final int ARG_EXCEL_OF_DATA_INDEX = 2;
	private static final int ARG_EXCEL_FOR_REPLACE_INDEX = 3;
	private static final int ARG_EXCEL_FOR_REPLACE_SHEET_INDEX = 4;
	private static final int ARG_EXCEL_FOR_REPLACE_ROW_NUM_INDEX = 5;

	private static String argWorkDir = ".\\";	// By default, use current dir
	private static String argWordTemplateName;

	private static String argExcelOfDataName;

	private static String argExcelForReplaceName;
	private static String argExcelForReplaceSheetName;
	private static int argExcelForReplaceRowNum;

	/**
	 * Replace word by excel.
	 * @param args, E:\hyj\ WordTemplate.doc ExcelOfData.xls ExcelForReplace.xls Sheet2 3
	 * @throws IOException
	 */
	public static void  main(String[] args) throws IOException {
		log.info("start");

		// init
		boolean isOK = parseArgs(args);
		if (!isOK) {
			log.error("Failed to parse args, see example");
			return;
		}

		// Excel for replace
		String excelForReplaceFilePath = getFilePath(argExcelForReplaceName);
		log.info(String.format("#### (%s, %s)", excelForReplaceFilePath, argExcelForReplaceSheetName));
		ExcelSheetForReplace excelSheetForReplace = new ExcelSheetForReplace(excelForReplaceFilePath, argExcelForReplaceSheetName);
		HashMap<String, ReplaceData> replaceDataMap = excelSheetForReplace.getReplaceDataMap(argExcelForReplaceRowNum);
		excelSheetForReplace.close();

		// Excel of Data
		String excelOfDataFilePath = getFilePath(argExcelOfDataName);
		log.info(String.format("#### (%s)", excelOfDataFilePath));
		ExcelOfData excelOfData = new ExcelOfData(excelOfDataFilePath);
		HashMap<String, String> replaceMap = excelOfData.getReplaceMap(replaceDataMap);
		excelOfData.close();

		// Word template
		int indexOfSuffix = argWordTemplateName.lastIndexOf(".");
		String strWordTemp = argWordTemplateName.substring(0, indexOfSuffix);
		String strWordTempSuffix = argWordTemplateName.substring(indexOfSuffix);

		String wordTemplatePath = getFilePath(argWordTemplateName);
		String newWordName = strWordTemp + "_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
		String newPath = getFilePath(newWordName + strWordTempSuffix);
		log.info(String.format("#### (%s)", wordTemplatePath));
		if (strWordTempSuffix.equalsIgnoreCase(WORD_SUFFIX)) {
			WordTemplateOfHWPF wt = new WordTemplateOfHWPF(wordTemplatePath);
			wt.replace(replaceMap);
			wt.saveAs(newPath);
			wt.close();
		} else if (strWordTempSuffix.equalsIgnoreCase(WORD_X_SUFFIX)) {
			WordTemplateOfXWPF wt = new WordTemplateOfXWPF(wordTemplatePath);
			wt.replace(replaceMap);
			wt.saveAs(newPath);
			wt.close();
		} else {
			throw new RuntimeException("Unsupported path = " + argWordTemplateName);
		}

		// Done
		log.info("done");
	}

	private static boolean parseArgs(String[] args) {
		log.info("args.length = " + args.length);
		for (int i = 0; i < args.length; i++) {
			log.info(String.format("%d: %s", i, args[i]));
		}

		// ARG_WORK_DIR_INDEX
		argWorkDir = args[ARG_WORK_DIR_INDEX];
		if (!argWorkDir.endsWith("\\")) {
			argWorkDir += "\\";
		}

		// ARG_WORD_TEMPLATE_INDEX
		argWordTemplateName = args[ARG_WORD_TEMPLATE_INDEX];

		// ARG_EXCEL_OF_DATA_INDEX
		argExcelOfDataName = args[ARG_EXCEL_OF_DATA_INDEX];

		// ARG_EXCEL_FOR_REPLACE_INDEX
		argExcelForReplaceName = args[ARG_EXCEL_FOR_REPLACE_INDEX];

		// ARG_EXCEL_FOR_REPLACE_INDEX
		argExcelForReplaceSheetName = args[ARG_EXCEL_FOR_REPLACE_SHEET_INDEX];

		// ARG_EXCEL_FOR_REPLACE_ROW_NUM_INDEX
		argExcelForReplaceRowNum = Integer.parseInt(args[ARG_EXCEL_FOR_REPLACE_ROW_NUM_INDEX]);
		if (argExcelForReplaceRowNum <= 0) {
			log.error(String.format("[%d] positive integer is allowed, but received %d.",
					ARG_EXCEL_FOR_REPLACE_ROW_NUM_INDEX, argExcelForReplaceRowNum));
			return false;
		}

		return true;
	}

	@Deprecated
	private static String getFilePath(String fileName, String suffix) {
		return argWorkDir + fileName + suffix;
	}

	private static String getFilePath(String fileName) {
		return argWorkDir + fileName;
	}
}
