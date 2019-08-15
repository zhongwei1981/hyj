package com.hyj.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
		//replaceInPara(document.getParagraphs(), replaceMap);
		//replaceXXX(document.getParagraphs(), replaceMap);
		replaceTables(document.getTables(), replaceMap);
	}

	private static void replaceXXX(List<XWPFParagraph> paras, Map<String, String> replaceMap) {
		for (XWPFParagraph para : paras) {
			String paraText = para.getParagraphText();
			log.info("100 paraText = " + paraText);
			for (Map.Entry<String, String> e : replaceMap.entrySet()) {
				String k = e.getKey();
				String v = e.getValue();
				log.info(String.format("#### [docx Table] repalce (%s, %s)", k, v));
				paraText = paraText.replace(e.getKey(), e.getValue());
			}
			log.info("101 paraText = " + paraText);
		}
	}

	private void replaceInPara(List<XWPFParagraph> paras, Map<String, String> replaceMap) {
		for (XWPFParagraph para : paras) {
			replaceInPara(para, replaceMap);
		}
	}
	private void replaceInPara(XWPFParagraph para, Map<String, String> replaceMap) {
		List<XWPFRun> runs;
		Matcher matcher;
		log.info("100 para.getParagraphText() = " + para.getParagraphText());
		if (this.matcher(para.getParagraphText()).find()) {
			runs = para.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				XWPFRun run = runs.get(i);
				String runText = run.toString();
				log.info("10 runText = " + runText);
				matcher = this.matcher(runText);
				if (matcher.find()) {
					while ((matcher = this.matcher(runText)).find()) {
						log.info("101 matcher.group(1) = " + matcher.group(1));
						runText = matcher.replaceFirst(String.valueOf(replaceMap.get(matcher.group(1))));
					}
					log.info("11 runText = " + runText);
					//直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
					//所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
					para.removeRun(i);
					para.insertNewRun(i).setText(runText);
				}
			}
		}
	}
	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}

	private static void replaceParagraphs(List<XWPFParagraph> paras, HashMap<String, String> replaceMap) {
		// 替换段落中的指定文字
//		StringBuilder text = new StringBuilder();
//        for (XWPFParagraph p : paragraphs) {
//            text.append(p.getText());
//        }

		for (XWPFParagraph paragraph : paras) {
			for (XWPFRun run : paragraph.getRuns()) {
				log.info("100-" + run.toString());
				String str = run.getText(run.getTextPosition());
				log.info("01 - " + str);
				for (Map.Entry<String, String> e : replaceMap.entrySet()) {
					String k = e.getKey();
					String v = e.getValue();
					log.info(String.format("#### [docx Paragraph] repalce (%s, %s)", k, v));
					str = str.replace(k, v);
				}
				log.info("02 - " + str);
				run.setText(str, 0);
			}
		}
	}

	private static void replaceTables(List<XWPFTable> tables, HashMap<String, String> replaceMap) {
		// 替换表格中的指定文字
		for (XWPFTable table : tables) {
			for (XWPFTableRow row : table.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					log.info("10 - " + cell.getParagraphs().size());
					String str = cell.getText();
					log.info("11 - " + str);
					for (Map.Entry<String, String> e : replaceMap.entrySet()) {
						String k = e.getKey();
						String v = e.getValue();
						log.info(String.format("#### [docx Table] repalce (%s, %s)", k, v));
						str = str.replace(e.getKey(), e.getValue());
					}
					log.info("12 - " + str);

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
