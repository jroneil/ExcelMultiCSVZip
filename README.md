# ExcelMultiCSVZip
Export excel to csv (single and multiple sheets) then zip the files
Application will convert multiple sheet and single sheet excel files to a single csv file.
The sheets will then be packaged in a zip file.
Excel files with extenstions xls and xlsx.

To run file open a command window
Type at the command prompt java –jar export <sourceDir> <targetDir> <zipDir>  debugOnOff

#Examples
Debug on
java -jar Excelcsv.jar "C:\\xlstocsv\\excel" "C:\\xlstocsv\\csv" "C:\\xlstocsv\\zip" true

Debug Off
java -jar Excelcsv.jar "C:\\xlstocsv\\excel" "C:\\xlstocsv\\csv" "C:\\xlstocsv\\zip"


Debug True will show comments

