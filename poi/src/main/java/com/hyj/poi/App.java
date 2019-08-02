package com.hyj.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Bookmark;
import org.apache.poi.hwpf.usermodel.Bookmarks;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import com.hyj.poi.word.OOXML.MSWordTool;

/**
 * HSSFWorkbook			excel文档对象
 * HSSFSheet			excel的sheet
 * HSSFRow				excel的行
 * HSSFCell				excel的单元格
 * HSSFFont				excel字体
 * HSSFName				名称
 * HSSFDataFormat		日期格式
 * HSSFHeader			sheet头
 * HSSFFooter			sheet尾
 * HSSFCellStyle		cell样式
 * HSSFDateUtil			日期
 * HSSFPrintSetup		打印
 * HSSFErrorConstants	错误信息表
 *
 * @author ETIX-2018-2
 */
public class App {
	private static Logger log = Logger.getLogger(App.class.getName());

	private static String TEST_EXCEL = "D:\\zhongwei\\hyj\\test.xls";
	private static String sheetName = "Sheet1";

	private static String TEST_WORD = "D:\\zhongwei\\hyj\\jacobTest.doc";
	private static String TEST_WORD_1 = "D:\\zhongwei\\hyj\\jacobTest_1.doc";
	private static String key = "##主题##";

	public static void main(String[] args) throws FileNotFoundException, IOException {
		log.info("## Hello World!");

		// Excel
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(TEST_EXCEL));
		HSSFSheet sheetAt = workbook.getSheet(sheetName);
		HSSFCell cell = sheetAt.getRow(1).getCell(1);
		CellType type = cell.getCellTypeEnum();
		switch (type) {
		case NUMERIC:
			log.info("#### (1, 1) NUMERIC = " + cell.getNumericCellValue());
			break;
		case STRING:
			log.info("#### (1, 1) STRING = " + cell.getStringCellValue());
			break;
		default:
			log.error("#### not supported type=" + type);
		}
		workbook.close();

		// Word -- OLE2
		readWordFile(TEST_WORD);

		// Word -- OOXML
		if (false) {
			MSWordTool tool = new MSWordTool();
			tool.setTemplate(TEST_WORD);
			Map<String, String> mapToReplace = new HashMap<>();
			mapToReplace.put(key, "3.14");
			tool.replaceBookMark(mapToReplace);
			tool.saveAs(TEST_WORD_1);
		}
	}

	public static <T> List<String> readWordFile(String path) throws FileNotFoundException, IOException {
		List<String> contextList = new ArrayList<>();
		InputStream stream = null;
		stream = new FileInputStream(new File(path));
		if (path.endsWith(".doc")) {
			HWPFDocument document = new HWPFDocument(stream);
			log.info("#### 0. " + document.getDocumentText());
			Bookmarks bms = document.getBookmarks();
			log.info("#### 1. " + bms.getBookmarksCount());
			for (int i = 0; i < bms.getBookmarksCount(); i++) {
				Bookmark bm = bms.getBookmark(i);
				log.info(String.format("#### 2. %s, %d, %d", bm.getName(), bm.getStart(), bm.getEnd()));
			}

			Range range = document.getRange();
			//读取word文本内容
			log.info(range.text());
			HashMap<String, String> map = new HashMap<>();
			map.put("key", "这是标题");
			//替换文本内容
			map.forEach((k, v) -> {
				range.replaceText(k, v);
			});

			WordExtractor extractor = new WordExtractor(document);
			String[] contextArray = extractor.getParagraphText();
			//			Arrays.asList(contextArray).forEach(context -> contextList.add(CharMatcher.whitespace().removeFrom(context)));
			extractor.close();

			FileOutputStream fos = new FileOutputStream(TEST_WORD_1);
			document.write(fos);
			fos.close();

			document.close();
		} else if (path.endsWith(".docx")) {
			XWPFDocument document = new XWPFDocument(stream).getXWPFDocument();
			List<XWPFParagraph> paragraphList = document.getParagraphs();
			//			paragraphList.forEach(paragraph -> contextList.add(CharMatcher.whitespace().removeFrom(paragraph.getParagraphText())));
			document.close();
		} else {
			log.info("not word doc:" + path);
		}

		stream.close();
		return contextList;
	}
}
