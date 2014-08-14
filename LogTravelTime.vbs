Dim strAPIKey, strFromAddress, strToAddress1, strToAddress2, strToAddress3, strPath, dtSwitch

strAPIKey = ""			' API key provided by MapQuest: http://developer.mapquest.com/web/products/open, click on "Get MapQuest AppKey"
strFromAddress = ""		' Start address (typically home address)
strToAddress1 = ""		' Destination address
strToAddress2 = ""		' Additional destination address (leave as "" if an additional destination address is not needed)
strToAddress3 = ""		' Additional destination address (leave as "" if an additional destination address is not needed)
strPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "Commute Duration.xlsx"		' Full path to log spreadsheet (assumed to be in the same directory as this script)
dtSwitch = #12:30:00 PM#		' Switch start and finish addresses at this time of day (change to 11:59 PM to disable)

GetTravelTime strAPIKey, strFromAddress, strToAddress1, strToAddress2, strToAddress3, strPath, dtSwitch

Sub GetTravelTime(strAPIKey, strFromAddress, strToAddress1, strToAddress2, strToAddress3, strPath, dtSwitch)

' Use MapQuest API to get travel time between two locations.
' Platinkreuz, August 2014

Dim objXMLHTTP, varToAddress, objExcel, wbk, wst, ii, strURL, rng

Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

varToAddress = Array(strToAddress1, strToAddress2, strToAddress3)

Set objExcel = CreateObject("Excel.Application")
Set wbk = objExcel.Workbooks.Open(strPath)

' Record travel time for each supplied destination on a different sheet.
For ii = 0 To UBound(varToAddress)
	If varToAddress(ii) <> "" Then
		Set wst = wbk.Worksheets(ii + 1)
		
		' Use home address as destination if after the given switch time.
		If Time < dtSwitch Then
			strURL = "http://www.mapquestapi.com/directions/v2/route?key=" & strAPIKey & "&from=" & strFromAddress & "&to=" & varToAddress(ii) & "&narrativeType=none"
		ElseIf Time > dtSwitch Then
			strURL = "http://www.mapquestapi.com/directions/v2/route?key=" & strAPIKey & "&from=" & varToAddress(ii) & "&to=" & strFromAddress & "&narrativeType=none"
		End If
		
		objXMLHTTP.Open "GET", strURL, False
		objXMLHTTP.send
		
		' If request goes through, record request time and travel duration.
		If objXMLHTTP.Status = 200 Then
			Set rng = wst.UsedRange.Columns(1).Cells(wst.UsedRange.Columns(1).Cells.Count).Offset(1)
			
			rng.Value = objXMLHTTP.responseText
			rng.Offset(, 1).Value = Now
			rng.Offset(, 2).Formula = "=TIME(HOUR(" & rng.Offset(, 1).Address & "),MINUTE(" & rng.Offset(, 1).Address & "),SECOND(" & rng.Offset(, 1).Address & "))"
			rng.Offset(, 3).Formula = "=GetsubElement(" & rng.Address & ",""route"",""realTime"")"
			rng.Offset(, 4).Formula = "=" & rng.Offset(, 3).Address & "/60"
			
		End If
    End If
Next

wbk.Close True

Set objExcel = Nothing
Set objXMLHTTP = Nothing

End Sub
