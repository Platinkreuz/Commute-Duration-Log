Commute-Duration-Log
====================

Queries the MapQuest API to get the travel time between home and work at fifteen-minute intervals throughout the day and creates a chart to display the results.

To install:

1. In InitializeWorkbook.vbs, provide the number of addresses to log commute duration for.
2. Run InitializeWorkbook.vbs to create the log workbook.
3. Trust access to the VBA project object model (In Excel 2010: File > Options > Trust Center > Trust Center Settings > Macro Settings).*
4. Run InitializeVBA.vbs to import the required modules into the log workbook: mdlRefreshCharts.bas, JSON.bas, cJSONScript.cls, and cStringBuilder.cls. Also creates a Workbook_Open function to automatically refresh charts with any new data. (May be asked to save file as a macro-enabled workbook, *.xlsm.)
5. In LogTravelTime.vbs, provide an API key and all from/to addresses.
6. Create a scheduled task to run LogTravelTime.vbs at fifteen minute intervals throughout the day.

* If you'd rather not do this, then you can manually open up the VBA project editor, import the four required modules, and add "RefreshCharts" to the Workbook_Open event.
