package com.hyj.ExcelToWord;

import org.apache.log4j.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

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
	public static void testReplaceAll() {
		ActiveXComponent wordApp = new ActiveXComponent("Word.Application");
		wordApp.setProperty("Visible", new Variant(false)); // 不可见打开word
		wordApp.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏

		Dispatch docs = wordApp.getProperty("Documents").toDispatch();

		Dispatch doc = Dispatch.call(docs, "Open", TEST_FILE_NAME).toDispatch();
		Dispatch selection = Dispatch.get(wordApp, "Selection").toDispatch();

		while (find(wordApp, selection, REPLACE_STR_1)) {
			Dispatch.put(selection, "Text", REPLACE_STR_1_NEW);
			Dispatch.call(selection, "MoveRight");
		}

		Dispatch.call(Dispatch.call(wordApp, "WordBasic").getDispatch(), "FileSaveAs", TEST_FILE_NAME_1);

		//Dispatch.call(doc, "Save");
		//Dispatch.call(doc, "Close", new Variant(true));
		Dispatch.call(wordApp, "Quit");
	}

	private static String TEST_FILE_NAME = "E:\\hyj\\jacobTest.doc";
	private static String TEST_FILE_NAME_1 = "E:\\hyj\\jacobTest_1.doc";

	private static String REPLACE_STR_1 = "##主题##";
	private static String REPLACE_STR_1_NEW = "3.14";

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
}
