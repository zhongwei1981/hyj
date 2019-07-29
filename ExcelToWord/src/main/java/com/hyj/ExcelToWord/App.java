package com.hyj.ExcelToWord;

import org.apache.log4j.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class App {
	private static Logger log = Logger.getLogger(App.class.getName());

	public static void main(String[] args) {
		log.info("#### Hello World!");

		testWordReplaceAll();
	}

	private static String TEST_WORD_FILE_NAME = "D:\\zhongwei\\hyj\\jacobTest.doc";
	private static String TEST_WORD_FILE_NAME_1 = "D:\\zhongwei\\hyj\\jacobTest_1.doc";

	private static String REPLACE_STR_1 = "##主题##";
	private static String REPLACE_STR_1_NEW = "3.14";
	private static void testWordReplaceAll() {
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

	private static boolean find(ActiveXComponent wordApp, Dispatch selection, String text) {
		boolean ret = false;
		if (text == null || text.equals("")) {
			return ret;
		}

		Dispatch find = ActiveXComponent.call(selection, "Find").toDispatch();	// 从selection所在位置开始查询
		Dispatch.put(find, "Text", text);			// 设置要查找的内容
		Dispatch.put(find, "Forward", "True");		// 向前查找
		Dispatch.put(find, "Format", "True");		// 设置格式
		Dispatch.put(find, "MatchCase", "True");	// 大小写匹配
		Dispatch.put(find, "MatchWholeWord", "True");	// 全字匹配
		ret = Dispatch.call(find, "Execute").getBoolean();	// 查找并选中
		return ret;
	}
}
