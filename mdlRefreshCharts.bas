Attribute VB_Name = "mdlRefreshCharts"
Option Explicit

Public Sub RefreshCharts()

' Add new data to commute duration charts.
' Platinkreuz, August 2014

Dim wst As Worksheet
Dim cht As ChartObject
Dim srs As Series
Dim rngDate As Range
Dim rngTime As Range
Dim rngDuration As Range
Dim jj As Integer
Dim ii As Integer
Dim rng As Range
Dim rngSeries() As Range

For Each wst In ThisWorkbook.Worksheets
    Set cht = wst.ChartObjects(1)
    
    ' Remove all existing series.
    For Each srs In cht.Chart.SeriesCollection
        srs.Delete
    Next srs
    
    Set rngDate = Intersect(wst.UsedRange, wst.UsedRange.Find("Date", , , , , , True).EntireColumn).Offset(1)
    Set rngTime = Intersect(wst.UsedRange, wst.UsedRange.Find("Time", , , , , , True).EntireColumn).Offset(1)
    Set rngDuration = Intersect(wst.UsedRange, wst.UsedRange.Find("Minutes", , , , , , True).EntireColumn).Offset(1)
    
    jj = 0
    ReDim rngSeries(0)
    Set rng = rngDate.Cells(1)
    
    ' Iterate through all records and add a range to the array for each day.
    For ii = 1 To rngDate.Cells.Count
        If Day(rng) <> Day(rngDate.Cells(ii).Value) Then
            ReDim Preserve rngSeries(jj)
            
            Set rngSeries(jj) = wst.Range(rng, rngDate.Cells(ii - 1))
            
            Set rng = rngDate.Cells(ii)
            
            jj = jj + 1
            
        End If
    Next ii
    
    ' Add the day-ranges to the chart.
    If jj > 0 Then
        For ii = 0 To UBound(rngSeries)
            Set srs = cht.Chart.SeriesCollection.NewSeries
            
            srs.XValues = Intersect(rngSeries(ii).EntireRow, rngTime)
            srs.Values = Intersect(rngSeries(ii).EntireRow, rngDuration)
            srs.Name = Format(rngSeries(ii).Cells(1).Value, "ddd, m/d/yy")
            
        Next ii
    End If
Next wst

End Sub

Public Function GetSubElement(strJSON As String, strElement As String, strSubElement As String)

' Extract the desired JSON element.
' Platinkreuz, August 2014

Dim objJSON As Scripting.Dictionary

Set objJSON = JSON.parse(strJSON)
Set objJSON = objJSON(strElement)

GetSubElement = objJSON(strSubElement)

End Function
