## IISLogsToExcel.exe - Converts IIS Logs to Excel

### Tool Details
================
- Tool allow folder drop, only if folder contains log file
- Tool converts log files of selected folder to excel file/s
- Converted excel files are generated at source location
- Checking first checkbox will generate single workbook with files names with respective file data
- Without having first checkbox checed, tool will generate one excel file for each log file
- Checking second checkbox will generate pivot table with hour as filter, time as row label, cs-uri-stem as value with count and time-taken as value with average
- Tool throws error message box if destination excel is locked or encounters exception
- If desination excel is locked, tool resumes only after closing error message box

### Tool State
==============
- Ready - Application launched and ready for processing
- Processing data for file <filename>... - Application processing log data for <filename>
- Process Completed. - Application completed processing log data
