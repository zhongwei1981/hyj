package com.hyj.ExcelToWord;

import org.apache.log4j.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class WordUtilTest extends TestCase {
	private static Logger log = Logger.getLogger(WordUtilTest.class.getName());

	/**
	 * Create the test case
	 *
	 * @param testName name of the test case
	 */
	public WordUtilTest(String testName) {
		super(testName);
	}

	/**
	 * @return the suite of tests being tested
	 */
	public static Test suite() {
		return new TestSuite(WordUtilTest.class);
	}

	/**
	 * Rigourous Test :-)
	 */
	public void testApp() {
		assertTrue(true);
	}
/*
	public static void testApp1() throws Exception {
		log.info("#### start");
		WordUtil word = new WordUtil();
		word.createNewDocument();
		word.createTable("", 5, 5);
		word.mergeCell(1, 1, 1, 1, 5);
		word.mergeCell(1, 2, 1, 2, 5);
		word.mergeCell(1, 3, 1, 3, 5);
		word.putTxtToCell(1, 1, 1, "主题");
		word.putTxtToCell(1, 2, 1, "时间");
		word.putTxtToCell(1, 3, 1, "人员");
		word.putTxtToCell(1, 4, 2, "说话了");
		word.saveAs(TEST_FILE_NAME);
		word.close();

		log.info("#### done");
	}

	public static void testApp2() throws Exception {
		log.info("#### start");
		WordUtil word = new WordUtil();
		word.openDocument(TEST_FILE_NAME);
		word.replaceAllText(REPLACE_STR_1, REPLACE_STR_1_NEW);
		word.saveAs(TEST_FILE_NAME_1);
		word.close();
		log.info("#### done");
	}
*/
/*	public static void testWordReplaceAll() {
		ActiveXComponent wordApp = new ActiveXComponent("Word.Application");
		wordApp.setProperty("Visible", new Variant(false)); // 不可见打开word
		wordApp.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏

		Dispatch docs = wordApp.getProperty("Documents").toDispatch();

		Dispatch doc = Dispatch.call(docs, "Open", TEST_WORD_FILE_NAME).toDispatch();
		Dispatch selection = Dispatch.get(wordApp, "Selection").toDispatch();

		while (find(wordApp, selection, REPLACE_STR_1)) {
			Dispatch.put(selection, "Text", REPLACE_STR_1_NEW);
			Dispatch.call(selection, "MoveRight");
		}

		Dispatch.call(Dispatch.call(wordApp, "WordBasic").getDispatch(), "FileSaveAs", TEST_WORD_FILE_NAME_1);

		//Dispatch.call(doc, "Save");
		//Dispatch.call(doc, "Close", new Variant(true));
		Dispatch.call(wordApp, "Quit");
	}

	private static String TEST_WORD_FILE_NAME = "E:\\hyj\\jacobTest.doc";
	private static String TEST_WORD_FILE_NAME_1 = "E:\\hyj\\jacobTest_1.doc";

	private static String REPLACE_STR_1 = "##主题##";
	private static String REPLACE_STR_1_NEW = "3.14";
*/
	private static boolean find(ActiveXComponent wordApp, Dispatch selection, String text) {
		boolean ret = false;
		if (text == null || text.equals("")) {
			return ret;
		}

		Dispatch find = wordApp.call(selection, "Find").toDispatch();	// 从selection所在位置开始查询
		Dispatch.put(find, "Text", text);			// 设置要查找的内容
		Dispatch.put(find, "Forward", "True");		// 向前查找
		Dispatch.put(find, "Format", "True");		// 设置格式
		Dispatch.put(find, "MatchCase", "True");	// 大小写匹配
		Dispatch.put(find, "MatchWholeWord", "True");	// 全字匹配
		ret = Dispatch.call(find, "Execute").getBoolean();	// 查找并选中
		return ret;
	}

	private static String TEST_EXCEL_FILE_NAME = "E:\\hyj\\jacobTest.xlsx";
/*	public static void testExcelReadWrite() {
		final boolean isReadOnly = false;

		ActiveXComponent ExcelApp = new ActiveXComponent("Excel.Application");
		ExcelApp.setProperty("Visible", new Variant(false));	// false为不显示打开Excel

		Dispatch wbs = ExcelApp.getProperty("Workbooks").toDispatch();

		Dispatch wb = Dispatch.invoke(wbs, "Open", Dispatch.Method,
				new Object[] { TEST_EXCEL_FILE_NAME, new Variant(false), new Variant(isReadOnly) }, // 是否以只读方式打开
				new int[1]).toDispatch();

		setExcelValue(wb, "sh1", "A1", "2");
		log.info(getExcelValue(wb, "基础设置","G10"));

		Dispatch.call(wb, "Save");
		Dispatch.call(wb, "Close", new Variant(false));
		ExcelApp.invoke("Quit", new Variant[] {});
	}
*/
	private static void setExcelValue(Dispatch wb, String sheetName ,String pos, String val) {
		//Dispatch sheet = Dispatch.get(workbook,"ActiveSheet").toDispatch();
		Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
		Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[] { sheetName }, new int[1]).toDispatch();
		Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] { pos }, new int[1]).toDispatch();
		Dispatch.put(cell, "Value", val);
	}

	private static String getExcelValue(Dispatch wb, String sheetName,String pos) {
		// sheet = Dispatch.get(workbook,"ActiveSheet").toDispatch();
		Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
		Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[] { new String(sheetName) }, new int[1]).toDispatch();
		Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] { pos }, new int[1]).toDispatch();
		String value = Dispatch.get(cell, "Value").toString();
		return value;
	}
}
