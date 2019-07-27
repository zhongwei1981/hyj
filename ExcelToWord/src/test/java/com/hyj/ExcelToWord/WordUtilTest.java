package com.hyj.ExcelToWord;

import org.apache.log4j.Logger;

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
		log.info("### xxx");
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
*/
	public static void testApp2() throws Exception {
		log.info("#### start");
		WordUtil word = new WordUtil();
		word.openDocument(TEST_FILE_NAME);
		word.replaceAllText(REPLACE_STR_1, REPLACE_STR_1_NEW);
		word.saveAs(TEST_FILE_NAME_1);
		word.close();
		log.info("#### done");
	}

	private static String TEST_FILE_NAME = "E:\\hyj\\jacobTest.doc";

	private static String REPLACE_STR_1 = "##主题##";
	private static String REPLACE_STR_1_NEW = "3.14";

	private static String TEST_FILE_NAME_1 = "E:\\hyj\\jacobTest_1.doc";
}
