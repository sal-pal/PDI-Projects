'Get data from clipboard for manipulation
Dim objMSIE
Set objMSIE = CreateObject("InternetExplorer.Application")
objMSIE.Navigate("about:blank")
FetchClipboardData = objMSIE.document.parentwindow.clipboardData.GetData("text")
objMSIE.Quit



'If line breaks found, remove new line characters and paste modified data to active cell; then select first cell of next row. 
lineAmount = 0
For i=1 To len(FetchClipboardData)
	If mid(FetchClipboardData,i,1) = vbCr Then 
		lineAmount = lineAmount + 1
		removedLineBreaks = Replace(FetchClipboardData, vbCrlf, " ")
		Set objExcel = GetObject(,"Excel.Application") 
		objExcel.ActiveCell.Value = removedLineBreaks
		row = objExcel.ActiveCell.Row
		objExcel.Cells(row+1,2).Select
		Exit For 
	End If
Next 


'If no line breaks found, enter ClipBrd data to active cell and select first cell of next row. 
If lineAmount = 0 Then 
	Set objExcel = GetObject(,"Excel.Application") 
	objExcel.ActiveCell.Value = FetchClipboardData
	row = objExcel.ActiveCell.Row
	objExcel.Cells(row+1,2).Select
End If 

























