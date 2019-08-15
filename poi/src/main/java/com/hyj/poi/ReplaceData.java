package com.hyj.poi;

import org.apache.log4j.Logger;

public class ReplaceData {
	private static Logger log = Logger.getLogger(ReplaceData.class.getName());

	public String sheetName;
	public int row;
	public int col;

	private String strRowCol;	//for backup only

	/**
	 *
	 * @param sheetName
	 * @param strRowCol, C5=(col:C-A=2, row:5-1=4)
	 */
	public ReplaceData(String sheetName, String strRowCol) {
		this.sheetName = sheetName;

		this.col = strRowCol.charAt(0) - 'A';
		this.row = Integer.parseInt(strRowCol.substring(1)) - 1;

		this.strRowCol = strRowCol;
	}

	public String getContent() {
		return String.format("(%s, %d, %d: %s)", sheetName, row, col, strRowCol);
	}
}
