mvn elipse:clean
mvn elipse:elipse

mvn dependency:sources
mvn dependency:resolve -Dclassifier=javadoc
mvn dependency:copy-dependencies -DoutputDirectory=lib

## See run.bat
mvn clean assembly:assembly
java -Xbootclasspath/a:./3rdPartyJars/jacob-1.19.jar; -jar -Djava.library.path=./3rdPartyLibs/jacob-1.19-x64.dll; target/ExcelToWord-0.0.1-SNAPSHOT-jar-with-dependencies.jar

mvn clean package

Microsoft Visual C++ 2005 SP1 Redistributable Package (x86)
	https://www.microsoft.com/en-us/download/details.aspx?id=5638
