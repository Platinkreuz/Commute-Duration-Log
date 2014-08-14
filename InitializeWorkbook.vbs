Dim intAddressCount, strPath

intAddressCount = 3		' Number of addresses to log
strPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "Commute Duration.xlsx"		' Full path to log spreadsheet (assumed to be in the same directory as this script)

InitializeWorkbook intAddressCount, strPath

Sub InitializeWorkbook(intAddressCount, strPath)

' Create a workbook and prepare it for commute duration logging.
' Platinkreuz, August 2014

Dim objExcel, wbk, wst, cht

Set objExcel = CreateObject("Excel.Application")
Set wbk = objExcel.Workbooks.Add

' Add or remove additional worksheets.
Do Until wbk.Worksheets.Count = intAddressCount
    If wbk.Worksheets.Count < intAddressCount Then
        wbk.Worksheets.Add , wbk.Worksheets(wbk.Worksheets.Count)
    ElseIf wbk.Worksheets.Count > intAddressCount Then
        wbk.Worksheets(wbk.Worksheets.Count).Delete
    End If
Loop

For Each wst In wbk.Worksheets
    ' Create column headings.
    wst.Range("A1:E1") = Array("JSON", "Date", "Time", "Seconds", "Minutes")
    
    ' Create a blank scatter plot.
    Set cht = wbk.Charts.Add
    Set cht = cht.Location(2, wst.Name)
    
    cht.ChartType = 75
    cht.HasTitle = False
    cht.Axes(1).TickLabels.NumberFormat = "[$-409]h AM/PM;@"
    cht.Axes(1).MaximumScale = 1.0
    cht.Axes(1).MajorUnit = 1 / 12
    cht.Axes(2).MinimumScale = 0.0
    cht.Axes(2).MaximumScale = 45.0
    cht.Parent.Top = wst.Range("G2").Top
    cht.Parent.Left = wst.Range("G2").Left
    cht.Parent.Height = wbk.Windows(1).VisibleRange.Height - wst.rows(1).Height * 6
    cht.Parent.Width = wbk.Windows(1).VisibleRange.Width - wst.UsedRange.Width - wst.Columns(1).Width * 2
    
    wst.Range("A1").Activate
    
Next

wbk.Worksheets(1).Activate
wbk.Close True, strPath

Set objExcel = Nothing

End Sub
