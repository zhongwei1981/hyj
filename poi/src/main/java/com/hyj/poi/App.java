package com.hyj.poi;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

import org.apache.log4j.Logger;

public class App {
	private static Logger log = Logger.getLogger(App.class.getName());

	private static final String EXCEL_SUFFIX = ".xls";
	private static final String WORD_SUFFIX = ".doc";

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
		String excelForReplaceFilePath = getFilePath(argExcelForReplaceName, EXCEL_SUFFIX);
		log.info(String.format("#### (%s, %s)", excelForReplaceFilePath, argExcelForReplaceSheetName));
		ExcelSheetForReplace excelSheetForReplace = new ExcelSheetForReplace(excelForReplaceFilePath, argExcelForReplaceSheetName);
		HashMap<String, ReplaceData> replaceDataMap = excelSheetForReplace.getReplaceDataMap(argExcelForReplaceRowNum);
		excelSheetForReplace.close();

		// Excel of Data
		String excelOfDataFilePath = getFilePath(argExcelOfDataName, EXCEL_SUFFIX);
		log.info(String.format("#### (%s)", excelOfDataFilePath));
		ExcelOfData excelOfData = new ExcelOfData(excelOfDataFilePath);
		HashMap<String, String> replaceMap = excelOfData.getReplaceMap(replaceDataMap);
		excelOfData.close();

		// Word template
		String wordTemplatePath = getFilePath(argWordTemplateName, WORD_SUFFIX);
		String newWordName = argWordTemplateName + "_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
		String newPath = getFilePath(newWordName, WORD_SUFFIX);
		log.info(String.format("#### (%s)", wordTemplatePath));
		WordTemplate wordTemplate = new WordTemplate(wordTemplatePath);
		wordTemplate.replaceRange(replaceMap);
		wordTemplate.saveAs(newPath);
		wordTemplate.close();

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
		if (!argWordTemplateName.endsWith(WORD_SUFFIX)) {
			log.error(String.format("[%d] Only %s is allowed, but received %s.",
					ARG_WORD_TEMPLATE_INDEX, WORD_SUFFIX, argWordTemplateName));
			return false;
		}
		argWordTemplateName = argWordTemplateName.substring(0,
				argWordTemplateName.length() - WORD_SUFFIX.length());

		// ARG_EXCEL_OF_DATA_INDEX
		argExcelOfDataName = args[ARG_EXCEL_OF_DATA_INDEX];
		if (!argExcelOfDataName.endsWith(EXCEL_SUFFIX)) {
			log.error(String.format("[%d] Only %s is allowed, but received %s.",
					ARG_EXCEL_OF_DATA_INDEX, EXCEL_SUFFIX, argExcelOfDataName));
			return false;
		}
		argExcelOfDataName = argExcelOfDataName.substring(0,
				argExcelOfDataName.length() - EXCEL_SUFFIX.length());

		// ARG_EXCEL_FOR_REPLACE_INDEX
		argExcelForReplaceName = args[ARG_EXCEL_FOR_REPLACE_INDEX];
		if (!argExcelForReplaceName.endsWith(EXCEL_SUFFIX)) {
			log.error(String.format("[%d] Only %s is allowed, but received %s.",
					ARG_EXCEL_FOR_REPLACE_INDEX, EXCEL_SUFFIX, argExcelForReplaceName));
			return false;
		}
		argExcelForReplaceName = argExcelForReplaceName.substring(0,
				argExcelForReplaceName.length() - EXCEL_SUFFIX.length());

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

	private static String getFilePath(String fileName, String suffix) {
		return argWorkDir + fileName + suffix;
	}
}
