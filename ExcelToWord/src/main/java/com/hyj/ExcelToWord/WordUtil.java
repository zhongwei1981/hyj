package com.hyj.ExcelToWord;

import org.apache.log4j.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

//注意java操作word关键是定位操作对象;
/**
 * jacob to R/W MS Word
 *
 * @author
 */
public class WordUtil {
	private static Logger log = Logger.getLogger(WordUtil.class.getName());

	private Dispatch docDispatch; // word doc
	private ActiveXComponent wordApp; // word运行程序对象
	private Dispatch docs; // 所有word文档集合
	private Dispatch selectionDispatch; // 选定的范围或插入点
	private boolean saveOnExit = true;

	public WordUtil() throws Exception {
		if (wordApp == null) {
			wordApp = new ActiveXComponent("Word.Application");
			wordApp.setProperty("Visible", new Variant(false)); // 不可见打开word
			wordApp.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
		}
		if (docs == null) {
			docs = wordApp.getProperty("Documents").toDispatch();
		}
	}

	/**
	 * 设置退出时参数
	 *
	 * @param saveOnExit: boolean true-退出时保存文件，false-退出时不保存文件
	 */
	public void setSaveOnExit(boolean saveOnExit) {
		this.saveOnExit = saveOnExit;
	}

	/**
	 * 创建一个新的word文档
	 */
	public void createNewDocument() {
		docDispatch = Dispatch.call(docs, "Add").toDispatch();
		selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
	}

	public void openDocument(String docPath) {
		closeDocument();
		docDispatch = Dispatch.call(docs, "Open", docPath).toDispatch();
		selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
	}

	/**
	 * 只读 打开一个保护文档,
	 *
	 * @param docPath-文件全名
	 * @param pwd-密码
	 */
	public void openDocumentOnlyRead(String docPath, String pwd) throws Exception {
		closeDocument();
		// doc = Dispatch.invoke(documents, "Open", Dispatch.Method,
		// new Object[]{docPath, new Variant(false), new Variant(true), new
		// Variant(true), pwd},
		// new int[1]).toDispatch();//打开word文件
		docDispatch = Dispatch.callN(docs, "Open", new Object[] { docPath, new Variant(false), new Variant(true),
				new Variant(true), pwd, "", new Variant(false) }).toDispatch();
		selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
	}

	public void openDocument(String docPath, String pwd) throws Exception {
		closeDocument();
		docDispatch = Dispatch
				.callN(docs, "Open",
						new Object[] { docPath, new Variant(false), new Variant(false), new Variant(true), pwd })
				.toDispatch();
		selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
	}

	/**
	 * 把选定的内容或插入点向上移动
	 *
	 * @param pos: 移动的距离
	 */
	public void moveUp(int pos) {
		if (selectionDispatch == null) {
			selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
		}
		for (int i = 0; i < pos; i++) {
			Dispatch.call(selectionDispatch, "MoveUp");
		}
	}

	/**
	 * 把选定的内容或者插入点向下移动
	 *
	 * @param pos 移动的距离
	 */
	public void moveDown(int pos) {
		if (selectionDispatch == null) {
			selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
		}
		for (int i = 0; i < pos; i++) {
			Dispatch.call(selectionDispatch, "MoveDown");
		}
	}

	/**
	 * 把选定的内容或者插入点向左移动
	 *
	 * @param pos 移动的距离
	 */
	public void moveLeft(int pos) {
		if (selectionDispatch == null) {
			selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
		}
		for (int i = 0; i < pos; i++) {
			Dispatch.call(selectionDispatch, "MoveLeft");
		}
	}

	/**
	 * 把选定的内容或者插入点向右移动
	 *
	 * @param pos 移动的距离
	 */
	public void moveRight(int pos) {
		if (selectionDispatch == null) {
			selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
		}
		for (int i = 0; i < pos; i++) {
			Dispatch.call(selectionDispatch, "MoveRight");
		}
	}

	/**
	 * 把插入点移动到文件首位置
	 *
	 */
	public void moveStart() {
		if (selectionDispatch == null) {
			selectionDispatch = Dispatch.get(wordApp, "Selection").toDispatch();
		}
		Dispatch.call(selectionDispatch, "HomeKey", new Variant(6));
	}

	/**
	 * 从选定内容或插入点开始查找文本
	 *
	 * @param toFindText 要查找的文本
	 * @return boolean true-查找到并选中该文本，false-未查找到文本
	 */
	@SuppressWarnings("static-access")
	public boolean find(String toFindText) {
		if (toFindText == null || toFindText.equals("")) {
			return false;
		}

		Dispatch find = wordApp.call(selectionDispatch, "Find").toDispatch();	// 从selection所在位置开始查询
		Dispatch.put(find, "Text", toFindText);		// 设置要查找的内容
		Dispatch.put(find, "Forward", "True");		// 向前查找
		Dispatch.put(find, "Format", "True");		// 设置格式
		Dispatch.put(find, "MatchCase", "True");	// 大小写匹配
		Dispatch.put(find, "MatchWholeWord", "True");	// 全字匹配
		return Dispatch.call(find, "Execute").getBoolean();	// 查找并选中
	}

	public boolean replaceText(String toFindText, String newText) {
		if (!find(toFindText)) {
			return false;
		}
		Dispatch.put(selectionDispatch, "Text", newText);
		return true;
	}

	public void replaceAllText(String toFindText, String newText) {
		while (find(toFindText)) {
			Dispatch.put(selectionDispatch, "Text", newText);
			Dispatch.call(selectionDispatch, "MoveRight");
		}
	}

	/**
	 * 在当前插入点插入字符串
	 *
	 * @param newText 要插入的新字符串
	 */
	public void insertText(String newText) {
		Dispatch.put(selectionDispatch, "Text", newText);
	}

	/**
	 *
	 * @param toFindText 要查找的字符串
	 * @param imagePath  图片路径
	 * @return
	 */
	public boolean replaceImage(String toFindText, String imagePath) {
		if (!find(toFindText)) {
			return false;
		}
		Dispatch.call(Dispatch.get(selectionDispatch, "InLineShapes").toDispatch(), "AddPicture", imagePath);
		return true;
	}

	/**
	 * 全局替换图片
	 *
	 * @param toFindText 查找字符串
	 * @param imagePath  图片路径
	 */
	public void replaceAllImage(String toFindText, String imagePath) {
		while (find(toFindText)) {
			Dispatch.call(Dispatch.get(selectionDispatch, "InLineShapes").toDispatch(), "AddPicture", imagePath);
			Dispatch.call(selectionDispatch, "MoveRight");
		}
	}

	/**
	 * 在当前插入点插入图片
	 *
	 * @param imagePath 图片路径
	 */
	public void insertImage(String imagePath) {
		Dispatch.call(Dispatch.get(selectionDispatch, "InLineShapes").toDispatch(), "AddPicture", imagePath);
	}

	/**
	 * 合并单元格
	 *
	 * @param tableIndex
	 * @param fstCellRowIdx
	 * @param fstCellColIdx
	 * @param secCellRowIdx
	 * @param secCellColIdx
	 */
	public void mergeCell(int tableIndex, int fstCellRowIdx, int fstCellColIdx, int secCellRowIdx, int secCellColIdx) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch fstCell = Dispatch.call(table, "Cell", new Variant(fstCellRowIdx), new Variant(fstCellColIdx))
				.toDispatch();
		Dispatch secCell = Dispatch.call(table, "Cell", new Variant(secCellRowIdx), new Variant(secCellColIdx))
				.toDispatch();
		Dispatch.call(fstCell, "Merge", secCell);
	}

	/**
	 * 在指定的单元格里填写数据
	 *
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 * @param txt
	 */
	public void putTxtToCell(int tableIndex, int cellRowIdx, int cellColIdx, String txt) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		Dispatch.put(selectionDispatch, "Text", txt);
	}

	/**
	 * 获得指定的单元格里数据
	 *
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 * @return
	 */
	public String getTxtFromCell(int tableIndex, int cellRowIdx, int cellColIdx) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		String ret = "";
		ret = Dispatch.get(selectionDispatch, "Text").toString();
		ret = ret.substring(0, ret.length() - 1); // 去掉最后的回车符;
		return ret;
	}

	/**
	 * 在当前文档拷贝剪贴板数据
	 *
	 * @param pos
	 */
	public void pasteExcelSheet(String pos) {
		moveStart();
		if (this.find(pos)) {
			Dispatch textRange = Dispatch.get(selectionDispatch, "Range").toDispatch();
			Dispatch.call(textRange, "Paste");
		}
	}

	/**
	 * 在当前文档指定的位置拷贝表格
	 *
	 * @param pos        当前文档指定的位置
	 * @param tableIndex 被拷贝的表格在word文档中所处的位置
	 */
	public void copyTable(String pos, int tableIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch range = Dispatch.get(table, "Range").toDispatch();
		Dispatch.call(range, "Copy");
		if (this.find(pos)) {
			Dispatch textRange = Dispatch.get(selectionDispatch, "Range").toDispatch();
			Dispatch.call(textRange, "Paste");
		}
	}

	/**
	 * 在当前文档指定的位置拷贝来自另一个文档中的表格
	 *
	 * @param anotherDocPath 另一个文档的磁盘路径
	 * @param tableIndex     被拷贝的表格在另一格文档中的位置
	 * @param pos            当前文档指定的位置
	 */
	public void copyTableFromAnotherDoc(String anotherDocPath, int tableIndex, String pos) {
		Dispatch doc2 = null;
		try {
			doc2 = Dispatch.call(docs, "Open", anotherDocPath).toDispatch();
			// 所有表格
			Dispatch tables = Dispatch.get(doc2, "Tables").toDispatch();
			// 要填充的表格
			Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
			Dispatch range = Dispatch.get(table, "Range").toDispatch();
			Dispatch.call(range, "Copy");
			if (this.find(pos)) {
				Dispatch textRange = Dispatch.get(selectionDispatch, "Range").toDispatch();
				Dispatch.call(textRange, "Paste");
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (doc2 != null) {
				Dispatch.call(doc2, "Close", new Variant(saveOnExit));
				doc2 = null;
			}
		}
	}

	/**
	 * 在当前文档指定的位置拷贝来自另一个文档中的图片
	 *
	 * @param anotherDocPath 另一个文档的磁盘路径
	 * @param shapeIndex     被拷贝的图片在另一格文档中的位置
	 * @param pos            当前文档指定的位置
	 */
	public void copyImageFromAnotherDoc(String anotherDocPath, int shapeIndex, String pos) {
		Dispatch doc2 = null;
		try {
			doc2 = Dispatch.call(docs, "Open", anotherDocPath).toDispatch();
			Dispatch shapes = Dispatch.get(doc2, "InLineShapes").toDispatch();
			Dispatch shape = Dispatch.call(shapes, "Item", new Variant(shapeIndex)).toDispatch();
			Dispatch imageRange = Dispatch.get(shape, "Range").toDispatch();
			Dispatch.call(imageRange, "Copy");
			if (this.find(pos)) {
				Dispatch textRange = Dispatch.get(selectionDispatch, "Range").toDispatch();
				Dispatch.call(textRange, "Paste");
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (doc2 != null) {
				Dispatch.call(doc2, "Close", new Variant(saveOnExit));
				doc2 = null;
			}
		}
	}

	/**
	 * 创建表格
	 *
	 * @param pos  位置
	 * @param cols 列数
	 * @param rows 行数
	 */
	public void createTable(String pos, int numCols, int numRows) {
		if (find(pos)) {
			Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
			Dispatch range = Dispatch.get(selectionDispatch, "Range").toDispatch();
			@SuppressWarnings("unused")
			Dispatch newTable = Dispatch.call(tables, "Add", range, new Variant(numRows), new Variant(numCols))
					.toDispatch();
			Dispatch.call(selectionDispatch, "MoveRight");
		} else {
			Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
			Dispatch range = Dispatch.get(selectionDispatch, "Range").toDispatch();
			@SuppressWarnings("unused")
			Dispatch newTable = Dispatch.call(tables, "Add", range, new Variant(numRows), new Variant(numCols))
					.toDispatch();
			Dispatch.call(selectionDispatch, "MoveRight");
		}
	}

	/**
	 * 在指定行前面增加行
	 *
	 * @param tableIndex word文件中的第N张表(从1开始)
	 * @param rowIndex   指定行的序号(从1开始)
	 */
	public void addTableRow(int tableIndex, int rowIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.call(rows, "Item", new Variant(rowIndex)).toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * 在第1行前增加一行
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 */
	public void addFirstTableRow(int tableIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.get(rows, "First").toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * 在最后1行前增加一行
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 */
	public void addLastTableRow(int tableIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.get(rows, "Last").toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * 增加一行
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 */
	public void addRow(int tableIndex) {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch.call(rows, "Add");
	}

	/**
	 * 增加一列
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 */
	public void addCol(int tableIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch.call(cols, "Add").toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * 在指定列前面增加表格的列
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 * @param colIndex   制定列的序号 (从1开始)
	 */
	public void addTableCol(int tableIndex, int colIndex) {
		// 所有表格
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		log.info(Dispatch.get(cols, "Count"));
		Dispatch col = Dispatch.call(cols, "Item", new Variant(colIndex)).toDispatch();
		// Dispatch col = Dispatch.get(cols, "First").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * 在第1列前增加一列
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 */
	public void addFirstTableCol(int tableIndex) {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch col = Dispatch.get(cols, "First").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * 在最后一列前增加一列
	 *
	 * @param tableIndex word文档中的第N张表(从1开始)
	 */
	public void addLastTableCol(int tableIndex) {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		// 要填充的表格
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// 表格的所有行
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch col = Dispatch.get(cols, "Last").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * 自动调整表格
	 *
	 */
	@SuppressWarnings("deprecation")
	public void autoFitTable() {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		int count = Dispatch.get(tables, "Count").toInt();
		for (int i = 0; i < count; i++) {
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i + 1)).toDispatch();
			Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
			Dispatch.call(cols, "AutoFit");
		}
	}

	/**
	 * 调用word里的宏以调整表格的宽度,其中宏保存在document下
	 *
	 */
	@SuppressWarnings("deprecation")
	public void callWordMacro() {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		int count = Dispatch.get(tables, "Count").toInt();
		Variant vMacroName = new Variant("Normal.NewMacros.tableFit");
		@SuppressWarnings("unused")
		Variant vParam = new Variant("param1");
		@SuppressWarnings("unused")
		Variant para[] = new Variant[] { vMacroName };
		for (int i = 0; i < count; i++) {
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i + 1)).toDispatch();
			Dispatch.call(table, "Select");
			Dispatch.call(wordApp, "Run", "tableFitContent");
		}
	}

	/**
	 * 设置当前选定内容的字体
	 *
	 * @param boldSize
	 * @param italicSize
	 * @param underLineSize 下划线
	 * @param colorSize     字体颜色
	 * @param size          字体大小
	 * @param name          字体名称
	 */
	public void setFont(boolean bold, boolean italic, boolean underLine, String colorSize, String size, String name) {
		Dispatch font = Dispatch.get(selectionDispatch, "Font").toDispatch();
		Dispatch.put(font, "Name", new Variant(name));
		Dispatch.put(font, "Bold", new Variant(bold));
		Dispatch.put(font, "Italic", new Variant(italic));
		Dispatch.put(font, "Underline", new Variant(underLine));
		Dispatch.put(font, "Color", colorSize);
		Dispatch.put(font, "Size", size);
	}

	/**
	 * 设置单元格被选中
	 *
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 */
	public void setTableCellSelected(int tableIndex, int cellRowIdx, int cellColIdx) {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
	}

	/**
	 * 设置选定单元格的垂直对起方式, 请使用setTableCellSelected选中一个单元格
	 *
	 * @param align 0-顶端, 1-居中, 3-底端
	 */
	public void setCellVerticalAlign(int verticalAlign) {
		Dispatch cells = Dispatch.get(selectionDispatch, "Cells").toDispatch();
		Dispatch.put(cells, "VerticalAlignment", new Variant(verticalAlign));
	}

	/**
	 * 设置当前文档中所有表格水平居中方式及其它一些格式,用在将word文件转化为html中,针对申报表
	 */
	@SuppressWarnings("deprecation")
	public void setApplyTableFormat() {
		Dispatch tables = Dispatch.get(docDispatch, "Tables").toDispatch();
		int tabCount = Integer.valueOf(Dispatch.get(tables, "Count").toString()); // log.info(tabCount);
		log.info("*******************************************************");
		for (int i = 1; i <= tabCount; i++) {
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i)).toDispatch();
			Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
			if (i == 1) {
				Dispatch.put(rows, "Alignment", new Variant(2)); // 1-居中,2-Right
				continue;
			}
			Dispatch.put(rows, "Alignment", new Variant(1)); // 1-居中
			Dispatch.call(table, "AutoFitBehavior", new Variant(1));// 设置自动调整表格方式,1-根据窗口自动调整
			Dispatch.put(table, "PreferredWidthType", new Variant(1));
			Dispatch.put(table, "PreferredWidth", new Variant(700));
			log.info(Dispatch.get(rows, "HeightRule").toString());
			Dispatch.put(rows, "HeightRule", new Variant(1)); // 0-自动wdRowHeightAuto,1-最小值wdRowHeightAtLeast,
																// 2-固定wdRowHeightExactly
			Dispatch.put(rows, "Height", new Variant(0.04 * 28.35));
			// int oldAlign = Integer.valueOf(Dispatch.get(rows, "Alignment").toString());
			// log.info("Algin:" + oldAlign);
		}
	}

	/**
	 * 设置段落格式
	 *
	 * @param alignment                    0-左对齐, 1-右对齐, 2-右对齐, 3-两端对齐, 4-分散对齐
	 * @param lineSpaceingRule
	 * @param lineUnitBefore
	 * @param lineUnitAfter
	 * @param characterUnitFirstLineIndent
	 */
	public void setParagraphsProperties(int alignment, int lineSpaceingRule, int lineUnitBefore, int lineUnitAfter,
			int characterUnitFirstLineIndent) {
		Dispatch paragraphs = Dispatch.get(selectionDispatch, "Paragraphs").toDispatch();
		Dispatch.put(paragraphs, "Alignment", new Variant(alignment)); // 对齐方式
		Dispatch.put(paragraphs, "LineSpacingRule", new Variant(lineSpaceingRule)); // 行距
		Dispatch.put(paragraphs, "LineUnitBefore", new Variant(lineUnitBefore)); // 段前
		Dispatch.put(paragraphs, "LineUnitAfter", new Variant(lineUnitAfter)); // 段后
		Dispatch.put(paragraphs, "CharacterUnitFirstLineIndent", new Variant(characterUnitFirstLineIndent)); // 首行缩进字符数
	}

	/**
	 * 打印当前段落格式, 使用前,请先选中段落
	 */
	public void getParagraphsProperties() {
		Dispatch paragraphs = Dispatch.get(selectionDispatch, "Paragraphs").toDispatch();
		String val = Dispatch.get(paragraphs, "LineSpacingRule").toString(); // 行距
		log.info("行距:" + val);
		val = Dispatch.get(paragraphs, "Alignment").toString(); // 对齐方式
		log.info("对齐方式:" + val); // 0-左对齐, 1-右对齐, 2-右对齐, 3-两端对齐, 4-分散对齐
		val = Dispatch.get(paragraphs, "LineUnitBefore").toString(); // 段前行数
		log.info("段前行数:" + val);
		val = Dispatch.get(paragraphs, "LineUnitAfter").toString(); // 段后行数
		log.info("段后行数:" + val);
		val = Dispatch.get(paragraphs, "FirstLineIndent").toString(); // 首行缩进
		log.info("首行缩进:" + val);
		val = Dispatch.get(paragraphs, "CharacterUnitFirstLineIndent").toString(); // 首行缩进字符数
		log.info("首行缩进字符数:" + val);
	}

	public void saveAs(String newPath) {
		Dispatch.call(Dispatch.call(wordApp, "WordBasic").getDispatch(), "FileSaveAs", newPath);
	}

	/**
	 * 文件保存为html格式
	 *
	 * @param savePath
	 * @param htmlPath
	 */
	public void saveAsHtml(String htmlPath) {
		Dispatch.invoke(docDispatch, "SaveAs", Dispatch.Method, new Object[] { htmlPath, new Variant(8) }, new int[1]);
	}

	/**
	 * 关闭文档
	 *
	 * @param val 0不保存修改 -1 保存修改 -2 提示是否保存修改
	 */
	public void closeDocument(int val) {
		Dispatch.call(docDispatch, "Close", new Variant(val));
		docDispatch = null;
	}

	/**
	 * 关闭当前word文档
	 *
	 */
	public void closeDocument() {
		if (docDispatch != null) {
			Dispatch.call(docDispatch, "Save");
			Dispatch.call(docDispatch, "Close", new Variant(saveOnExit));
			docDispatch = null;
		}
	}

	public void closeDocumentWithoutSave() {
		if (docDispatch != null) {
			Dispatch.call(docDispatch, "Close", new Variant(false));
			docDispatch = null;
		}
	}

	/**
	 * 关闭全部应用
	 *
	 */
	public void close() {
		// closeDocument();
		if (wordApp != null) {
			Dispatch.call(wordApp, "Quit");
			wordApp = null;
		}
		selectionDispatch = null;
		docs = null;
	}

	/**
	 * 打印当前word文档
	 *
	 */
	public void printFile() {
		if (docDispatch != null) {
			Dispatch.call(docDispatch, "PrintOut");
		}
	}

	/**
	 * 保护当前档,如果不存在, 使用expression.Protect(Type, NoReset, Password)
	 *
	 * @param pwd WdProtectionType 可以是下列 WdProtectionType 常量之一：
	 *            1-wdAllowOnlyComments, 2-wdAllowOnlyFormFields,
	 *            0-wdAllowOnlyRevisions, -1-wdNoProtection, 3-wdAllowOnlyReading
	 *
	 */
	public void protectedWord(String pwd) {
		String protectionType = Dispatch.get(docDispatch, "ProtectionType").toString();
		if (protectionType.equals("-1")) {
			Dispatch.call(docDispatch, "Protect", new Variant(3), new Variant(true), pwd);
		}
	}

	/**
	 * 解除文档保护,如果存在
	 *
	 * @param pwd WdProtectionType 常量之一(Long 类型，只读)：
	 *            1-wdAllowOnlyComments,2-wdAllowOnlyFormFields、
	 *            0-wdAllowOnlyRevisions,-1-wdNoProtection, 3-wdAllowOnlyReading
	 *
	 */
	public void unProtectedWord(String pwd) {
		String protectionType = Dispatch.get(docDispatch, "ProtectionType").toString();
		if (protectionType.equals("3")) {
			Dispatch.call(docDispatch, "Unprotect", pwd);
		}
	}

	/**
	 * 设置word文档安全级别
	 *
	 * @param value 1-msoAutomationSecurityByUI 使用“安全”对话框指定的安全设置。
	 *              2-msoAutomationSecurityForceDisable 在程序打开的所有文件中禁用所有宏，而不显示任何安全提醒。
	 *              3-msoAutomationSecurityLow 启用所有宏，这是启动应用程序时的默认值。
	 */
	public void setAutomationSecurity(int value) {
		wordApp.setProperty("AutomationSecurity", new Variant(value));
	}

	/**
	 * 读取文档中第paragraphsIndex段文字的内容;
	 *
	 * @param paragraphsIndex
	 * @return
	 */
	public String getParagraphs(int paragraphsIndex) {
		String ret = "";
		Dispatch paragraphs = Dispatch.get(docDispatch, "Paragraphs").toDispatch(); // 所有段落
		int paragraphCount = Dispatch.get(paragraphs, "Count").getInt(); // 一共的段落数
		Dispatch paragraph = null;
		Dispatch range = null;
		if (paragraphCount > paragraphsIndex && 0 < paragraphsIndex) {
			paragraph = Dispatch.call(paragraphs, "Item", new Variant(paragraphsIndex)).toDispatch();
			range = Dispatch.get(paragraph, "Range").toDispatch();
			ret = Dispatch.get(range, "Text").toString();
		}
		return ret;
	}
}
