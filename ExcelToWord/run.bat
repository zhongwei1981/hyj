rem mvn clean assembly:assembly

java -Xbootclasspath/a:./3rdPartyJars/jacob-1.19.jar; -jar -Djava.library.path=./3rdPartyLibs/jacob-1.19-x64.dll; target/ExcelToWord-0.0.1-SNAPSHOT-jar-with-dependencies.jar