##https://poi.apache.org/apidocs/index.html?org/apache/poi/openxml4j/opc/internal/package-summary.html
##https://poi.apache.org/apidocs/3.17/
##http://blog.sina.com.cn/s/blog_885585cb0101gnz7.html

##https://www.cnblogs.com/gdwkong/p/8669220.html
一、Apache POI概述, see http://poi.apache.org/components/spreadsheet/
Apache POI结构:
-- 2007, Microsoft OLE 2复合文档格式, Microsoft Word 97 - 2007
-- 2007+, OOXML (Office Open XML)

-- OLE 2 Documents
	POIFS - the oldest and most stable part of POI. It is our port of the OLE 2 Compound Document Format
		此组件是所有其他POI元素的基本因素。 它用于显式读取不同的文件。
-- OLE 2 Document Properties
	HPSF - OLE 2 property set format, 用于提取MS-Office文件的属性集。
		Property sets are mostly use to store a document's properties (title, author, date of last modification etc.),
			but they can be used for application-specific purposes as well.
-- Excel Documents
	HSSF - 提供读写Microsoft Excel格式档案(*.xls)的功能。			-- Microsoft Excel 97 (-2003) file format (BIFF8)
	XSSF - 提供读写Microsoft Excel OOXML格式档案(*.xlsx)的功能。	-- Microsoft Excel XML (2007+) file format (OOXML)
		SS is a package that provides common support for both formats with a common API. 
	SXSSF - (Since POI 3.8 beta3), an API-compatible streaming extension of XSSF to be used 
		when very large spreadsheets have to be produced, and heap space is limited.
	see http://poi.apache.org/components/spreadsheet/
-- Word Documents
	HWPF - 提供读写Microsoft Word格式档案(*.doc)的功能。			-- Microsoft Word 97 (-2003) file format
	XWPF - 提供读写Microsoft Word OOXML格式档案(*.docx)的功能。	-- WordprocessingML (2007+) format from the OOXML specification
-- PowerPoint Documents
	HSLF - Microsoft PowerPoint 97(-2003) file format
	XSLF - PresentationML (2007+) format from the OOXML specification
-- Visio Documents
	HDGF - Microsoft Visio 97(-2003) file format. It currently only supports reading at a very low level, and simple text extraction.
		提供读写Microsoft Visio格式档案的功能, 包含MS-Visio二进制文件的类和方法。 HDGF=Horrible DiaGram格式
	XDGF - Microsoft Visio XML (.vsdx) file format
-- Publisher Documents
	HPBF - Microsoft Publisher 98(-2007) file format
		用于读取和写入MS-Publisher文件。 HPBF=Horrible Publisher格式
-- Outlook Attachments - TNEF (winmail.dat) 
	HMEF - Microsoft TNEF (Transport Neutral Encoding Format) file format.
		TNEF is sometimes used by Outlook for encoding the message, and will typically come through as winmail.dat.
		HMEF currently only supports reading at a low level, but we hope to add text and attachment extraction.
-- Outlook Messages
	HSMF - Microsoft Outlook message file format
		It currently only some of the textual content of MSG files, and some attachments.
		Further support and documentation is coming in slowly.
		For now, users are advised to consult the unit tests for example use.
		
	
## HSSF XSSF SXSSF
    https://www.jianshu.com/p/db69d6901c38

POI 3.17版本是支持jdk6的最后版本

二、HSSF概况
	HSSF 是Horrible SpreadSheet Format的缩写，通过HSSF，你可以用纯Java代码来读取、写入、修改Excel文件。
	HSSF 为读取操作提供了两类API：
		user model: 用户模型
		event user model: 事件-用户模型

三、 POI EXCEL文档结构类
	HSSFWorkbook		excel文档对象
	HSSFSheet			excel的sheet
	HSSFRow				excel的行
	HSSFCell			excel的单元格
	HSSFFont			excel字体
	HSSFName			名称
	HSSFDataFormat		日期格式
	HSSFHeader			sheet头
	HSSFFooter			sheet尾
	HSSFCellStyle		cell样式
	HSSFDateUtil		日期
	HSSFPrintSetup		打印
	HSSFErrorConstants	错误信息表

##https://blog.csdn.net/u010770896/article/details/79863477		- OOXML Documents
Range：	它表示一个范围，这个范围可以是整个文档，也可以是里面的某一小节（Section），也可以是某一个段落（Paragraph），还可以是拥有共同属性的一段文本（CharacterRun）。
	Section：		word文档的一个小节，一个word文档可以由多个小节构成。
	Paragraph：		word文档的一个段落，一个小节可以由多个段落构成。
	CharacterRun：	具有相同属性的一段文本，一个段落可以由多个CharacterRun组成。
	Section、Paragraph、CharacterRun和Table都继承自Range
Table：	一个表格。
	TableRow：	表格对应的行。
	TableCell：	表格对应的单元格。
