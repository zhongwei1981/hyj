mvn elipse:clean
mvn elipse:elipse

mvn dependency:sources
mvn dependency:resolve -Dclassifier=javadoc
mvn dependency:copy-dependencies -DoutputDirectory=lib

mvn clean package