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

	private static String argWorkDir = ".\\";	// By default, use current dir
	private static String argWordTemplateName;

	private static String argExcelForReplaceName;
	private static String argExcelForReplaceSheetName;
	private static int argExcelForReplaceRowNum;

	/**
	 * Replace word by excel, ex [E:\hyj\, WordTemplate.doc, ExcelForReplace.xls, Sheet2, 3].
	 * @param args
	 * 		0 - working dir
	 * 		1 - word path
	 * 		2 - excel path
	 * 		3 - excel sheet name
	 * 		4 - excel sheet cell row num, ex 3
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
		//FIXME: v = double -> String -> double
		String excelFilePath = getFilePath(argExcelForReplaceName, EXCEL_SUFFIX);

		ExcelSheetForReplace excelSheetForReplace = new ExcelSheetForReplace(excelFilePath, argExcelForReplaceSheetName);
		HashMap<String, String> replaceMap = excelSheetForReplace.getReplaceMap(argExcelForReplaceRowNum);
		excelSheetForReplace.close();

		// Word template
		String wordTemplatePath = getFilePath(argWordTemplateName, WORD_SUFFIX);
		String newWordName = argWordTemplateName + "_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
		String newPath = getFilePath(newWordName, WORD_SUFFIX);

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

		// args[0]
		argWorkDir = args[0];
		if (!argWorkDir.endsWith("\\")) {
			argWorkDir += "\\";
		}

		// args[1]
		argWordTemplateName = args[1];
		if (!argWordTemplateName.endsWith(WORD_SUFFIX)) {
			log.error(String.format("[1] Only %s is allowed, but received %s.",
					WORD_SUFFIX, argWordTemplateName));
			return false;
		}
		argWordTemplateName = argWordTemplateName.substring(0,
				argWordTemplateName.length() - WORD_SUFFIX.length());

		// args[2]
		argExcelForReplaceName = args[2];
		if (!argExcelForReplaceName.endsWith(EXCEL_SUFFIX)) {
			log.error(String.format("[2] Only %s is allowed, but received %s.",
					EXCEL_SUFFIX, argExcelForReplaceName));
			return false;
		}
		argExcelForReplaceName = argExcelForReplaceName.substring(0,
				argExcelForReplaceName.length() - EXCEL_SUFFIX.length());

		// args[3]
		argExcelForReplaceSheetName = args[3];

		// args[4]
		argExcelForReplaceRowNum = Integer.parseInt(args[4]);
		if (argExcelForReplaceRowNum <= 0) {
			log.error(String.format("[4] positive integer is allowed, but received %d.",
					argExcelForReplaceRowNum));
			return false;
		}

		return true;
	}

	private static String getFilePath(String fileName, String suffix) {
		return argWorkDir + fileName + suffix;
	}
}
