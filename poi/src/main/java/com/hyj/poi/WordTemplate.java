package com.hyj.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class WordTemplate {
	private static Logger log = Logger.getLogger(WordTemplate.class.getName());

	private static NumberFormat NF6 = NumberFormat.getNumberInstance();

	static {
		NF6.setMaximumFractionDigits(6);
		NF6.setRoundingMode(RoundingMode.UP);
	}

	FileInputStream fis;
	HWPFDocument document;

	public WordTemplate(String path) throws IOException {
		fis = new FileInputStream(path);
		document = new HWPFDocument(fis);
	}

	public void close() throws IOException {
		document.close();
		fis.close();
	}

	public void replaceRange(HashMap<String, String> replaceMap) {
		Range range = document.getRange();
		replaceMap.forEach((k, v) -> {
			String val = NF6.format(Double.parseDouble(v));
			log.info(String.format("#### repalce (%s, %s)", k, val));
			range.replaceText(k, val);
		});
	}

	public void saveAs(String newPath) throws IOException {
		FileOutputStream fos = new FileOutputStream(newPath);
		document.write(fos);
		fos.close();
	}
}
