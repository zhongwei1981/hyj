package com.hyj.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordTemplateOfXWPF {
	private static Logger log = Logger.getLogger(WordTemplateOfXWPF.class.getName());

	FileInputStream fis;
	XWPFDocument document;

	public WordTemplateOfXWPF(String path) throws IOException {
		fis = new FileInputStream(path);
		document = new XWPFDocument(fis);
	}

	public void close() throws IOException {
		document.close();
		fis.close();
	}

	public void replace(HashMap<String, String> replaceMap) {
		replaceParagraphs(document.getParagraphs(), replaceMap);
		replaceTables(document.getTables(), replaceMap);
	}

	private static void replaceParagraphs(List<XWPFParagraph> paras, HashMap<String, String> replaceMap) {
		// 替换段落中的指定文字
		for (XWPFParagraph paragraph : paras) {
			//Check if need replace or not
			String textOfPara = paragraph.getText();
			if (textOfPara.indexOf(Common.START_KEY) < 0) {
				continue;
			}
			log.info("To replace: " + textOfPara);

			// Find & Replace
			Iterator<XWPFRun> it = paragraph.getRuns().iterator();
			while (it.hasNext()) {
				XWPFRun runStart = it.next();
				String start = runStart.toString();
				if (!start.equalsIgnoreCase(Common.START_KEY)) {
					continue;
				}

				// Replace START_KEY
				runStart.setText("", 0);

				// Replace key
				XWPFRun runKey = it.next();
				String key = runKey.toString();
				// Replace END_KEY & Done
				while (true) {
					XWPFRun runEnd = it.next();

					int indexOfEndKey = runEnd.toString().indexOf(Common.END_KEY);
					if (indexOfEndKey == 0) {
						replaceRun(runKey, replaceMap, key);
						runEnd.setText("", 0);
						break;
					} else if (indexOfEndKey > 0) {
						key += runEnd.toString().substring(0, indexOfEndKey);
						replaceRun(runKey, replaceMap, key);
						runEnd.setText("", 0);
						break;
					}

					key += runEnd.toString();
					runEnd.setText("", 0);
				}
			}
		}
	}

	private static void replaceRun(XWPFRun run, HashMap<String, String> replaceMap, String key) {
		if (!replaceMap.containsKey(key)) {
			throw new RuntimeException("Unexpected key: " + key);
		}

		log.info(String.format("replaced (%s <- %s)", key, replaceMap.get(key)));
		run.setText(replaceMap.get(key), 0);
	}

	private static void replaceTables(List<XWPFTable> tables, HashMap<String, String> replaceMap) {
		// 替换表格中的指定文字
		for (XWPFTable table : tables) {
			for (XWPFTableRow row : table.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
//					log.info("10 - " + cell.getParagraphs().size());
					String str = cell.getText();
//					log.info("11 - " + str);
					for (Map.Entry<String, String> e : replaceMap.entrySet()) {
//						log.info(String.format("#### [docx Table] repalce (%s, %s)", k, v));
						str = str.replace(e.getKey(), e.getValue());
					}
//					log.info("12 - " + str);

					// FIXME: here we consider the cell.getParagraphs().size() is always 1
					cell.removeParagraph(0);
					cell.setText(str);
				}
			}
		}
	}

	public void saveAs(String newPath) throws IOException {
		FileOutputStream fos = new FileOutputStream(newPath);
		document.write(fos);
		fos.close();
	}
}
