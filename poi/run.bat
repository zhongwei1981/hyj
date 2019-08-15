rem mvn clean assembly:assembly

rem java -jar target\poi-0.0.1-SNAPSHOT-jar-with-dependencies.jar jacobTest test Sheet2 3
rem java -jar target\poi-0.0.1-SNAPSHOT-jar-with-dependencies.jar %1 %2 %3 %4 %5
rem java -jar target\poi-0.0.1-SNAPSHOT-jar-with-dependencies.jar E:\hyj\ WordTemplate.doc ExcelOfData.xls ExcelForReplace.xls Sheet3 3

rem java -jar target\poi-0.0.1-SNAPSHOT-jar-with-dependencies.jar D:\zhongwei\hyj TestWordTemplate.doc TestExcelOfData.xls TestExcelForReplace.xls Sheet3 3
rem java -jar target\poi-0.0.1-SNAPSHOT-jar-with-dependencies.jar D:\zhongwei\hyj TestWordTemplate.doc TestExcelOfData.xlsx TestExcelForReplace.xlsx Sheet3 3
java -jar target\poi-0.0.1-SNAPSHOT-jar-with-dependencies.jar D:\zhongwei\hyj TestWordTemplate.docx TestExcelOfData.xlsx TestExcelForReplace.xlsx Sheet3 5