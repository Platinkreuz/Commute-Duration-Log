Dim strWorkbookPath, strModulePath

strWorkbookPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "Commute Duration.xlsx"		' Full path to log workbook
strModulePath = ""											' Path to folder containing downloaded modules

InitializeVBA strWorkbookPath, strModulePath

Sub InitializeVBA(strWorkbookPath, strModulePath)

' Import required VBA modules and add custom functions.
' Platinkreuz, August 2014

Dim objExcel, wbk, objVBIDE, vbaProj

Set objExcel = CreateObject("Excel.Application")
Set wbk = objExcel.Workbooks.Open(strWorkbookPath)
Set vbaProj = wbk.VBProject

vbaProj.VBComponents.Import strModulePath & "mdlRefreshCharts.bas"
vbaProj.VBComponents.Import strModulePath & "JSON.bas"
vbaProj.VBComponents.Import strModulePath & "cJSONScript.cls"
vbaProj.VBComponents.Import strModulePath & "cStringBuilder.cls"

vbaProj.VBComponents("ThisWorkbook").CodeModule.InsertLines 1, "Private Sub Workbook_Open()"
vbaProj.VBComponents("ThisWorkbook").CodeModule.InsertLines 2, ""
vbaProj.VBComponents("ThisWorkbook").CodeModule.InsertLines 3, "RefreshCharts"
vbaProj.VBComponents("ThisWorkbook").CodeModule.InsertLines 4, ""
vbaProj.VBComponents("ThisWorkbook").CodeModule.InsertLines 5, "End Sub"

wbk.Close True

Set objExcel = Nothing

End Sub