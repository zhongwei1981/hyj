package com.hyj.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class WordTemplateOfHWPF {
	private static Logger log = Logger.getLogger(WordTemplateOfHWPF.class.getName());

	FileInputStream fis;
	HWPFDocument document;

	public WordTemplateOfHWPF(String path) throws IOException {
		fis = new FileInputStream(path);
		document = new HWPFDocument(fis);
	}

	public void close() throws IOException {
		document.close();
		fis.close();
	}

	public void replace(HashMap<String, String> replaceMap) {
		Range range = document.getRange();
		replaceMap.forEach((k, v) -> {
			log.info(String.format("#### [doc] repalce (%s, %s)", k, v));
			range.replaceText(k, v);
		});
	}

	public void saveAs(String newPath) throws IOException {
		FileOutputStream fos = new FileOutputStream(newPath);
		document.write(fos);
		fos.close();
	}
}
