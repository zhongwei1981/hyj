package com.hyj.ExcelToWord;

import org.apache.log4j.Logger;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class HyjTest extends TestCase {
	private static Logger log = Logger.getLogger(HyjTest.class.getName());

	/**
	 * Create the test case
	 *
	 * @param testName name of the test case
	 */
	public HyjTest(String testName) {
		super(testName);
	}

	/**
	 * @return the suite of tests being tested
	 */
	public static Test suite() {
		return new TestSuite(HyjTest.class);
	}

	/**
	 * Rigourous Test :-)
	 */
	public void testApp() {
		assertTrue(true);
		log.info("### xxx");
	}

	public static void testApp1() throws Exception {
		log.info("#### start");
		Hyj word = new Hyj();
		word.createNewDocument();
		word.createTable("", 5, 5);
		word.mergeCell(1, 1, 1, 1, 5);
		word.mergeCell(1, 2, 1, 2, 5);
		word.mergeCell(1, 3, 1, 3, 5);
		word.putTxtToCell(1, 1, 1, "主题");
		word.putTxtToCell(1, 2, 1, "时间");
		word.putTxtToCell(1, 3, 1, "人员");
		word.putTxtToCell(1, 4, 2, "说话了");
		word.save("E:\\hyj\\jacobTest.doc");
		word.close();

		log.info("#### done");
	}
}
