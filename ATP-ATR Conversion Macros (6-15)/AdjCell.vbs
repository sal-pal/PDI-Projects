'Get data from clipboard for manipulation
Dim objMSIE
Set objMSIE = CreateObject("InternetExplorer.Application")
objMSIE.Navigate("about:blank")
FetchClipboardData = objMSIE.document.parentwindow.clipboardData.GetData("text")
objMSIE.Quit


'If line breaks found, remove new line characters and paste modified data to active cell; then select adjacent cell to the right. 
lineAmount = 0
For i=1 To len(FetchClipboardData)
	If mid(FetchClipboardData,i,1) = vbCr Then 
		lineAmount = lineAmount + 1
		removedLineBreaks = Replace(FetchClipboardData, vbCrlf, " ")
		Set objExcel = GetObject(,"Excel.Application") 
		objExcel.ActiveCell.Value = removedLineBreaks
		objExcel.ActiveCell.Offset(0,1).Select
		Exit For 
	End If
Next 


'If no line breaks found, enter ClipBrd data to active cell and select adjacent cell to the right. 
If lineAmount = 0 Then 
	Set objExcel = GetObject(,"Excel.Application") 
	objExcel.ActiveCell.Value = FetchClipboardData
	objExcel.ActiveCell.Offset(0,1).Select
End If 











