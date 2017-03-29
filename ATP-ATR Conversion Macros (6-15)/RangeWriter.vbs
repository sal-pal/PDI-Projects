	'MACRO FOR WRITING RANGE VALUES WITHIN EXCEL BASED ON THE PRESNECE OF KEYWORDS


'Get data from clipboard for manipulation
Dim objMSIE
Set objMSIE = CreateObject("InternetExplorer.Application")
objMSIE.Navigate("about:blank")
FetchClipboardData = objMSIE.document.parentwindow.clipboardData.GetData("text")
objMSIE.Quit



'Copy and paste clipboard data to active cell, but check for and remove line breaks first. 
Set objExcel = GetObject(,"Excel.Application")
lineAmount = 0
For i=1 To len(FetchClipboardData)
	If mid(FetchClipboardData,i,1) = vbCr Then 
		lineAmount = lineAmount + 1
		removedLineBreaks = Replace(FetchClipboardData, vbCrlf, " ")
		Set objExcel = GetObject(,"Excel.Application") 
		objExcel.ActiveCell.Value = removedLineBreaks
		Exit For 
	End If
Next 
'If no line breaks found, enter ClipBrd data to active cell and select adjacent cell to the right. 
If lineAmount = 0 Then 
	Set objExcel = GetObject(,"Excel.Application") 
	objExcel.ActiveCell.Value = FetchClipboardData
End If 



'Checking for which key word exist in the string.
 keyWords = array("Maximum.","maximum.","Greater than","greater than","Minimum.","minimum.","Less than","less than")
 count = 0
 For each word in split(FetchClipboardData)	
	Set re = new regexp
	re.Pattern = word
	'If word contains an illegal char, go to the next iteration. 
	On Error Resume Next
	If re.Test(keyWords(0)) = True or re.Test(keyWords(1)) = True Then 
		'Ignore this empty block. It's only syntactically needed. 
	End If 
	'Checking whether an error has occurred. 
	If Err.Number = 0 Then 
		'If string contains a "MAX" keyword, search for upper boundry value in string.
		If re.Test(keyWords(0)) = True or re.Test(keyWords(1)) = True or re.Test(keyWords(2)) = True or re.Test(keyWords(3)) = True Then
			For each token in split(FetchClipboardData)
				For i=1 To len(token)
					'If token is a number, paste upper and lower range values in spreadsheet. 
					If mid(token,i,1) = "0" or mid(token,i,1) = "1" or mid(token,i,1) = "2" or mid(token,i,1) = "3" or mid(token,i,1) = "4" or mid(token,i,1) = "5" or mid(token,i,1) = "6" or mid(token,i,1) = "7" or mid(token,i,1) = "8" or mid(token,i,1) = "9" Then
						objExcel.ActiveCell.Offset(0,-3).Value = 3
						objExcel.ActiveCell.Offset(0,1).Value = 0
						objExcel.ActiveCell.Offset(0,2).Value = token
						row = objExcel.ActiveCell.Row
						objExcel.Cells(row+1,2).Select
						Exit For 
					End If 
				Next 			
			Next				
		'If string contains a "MIN" keyword, search for lower boundry value in string.
		Elseif re.Test(keyWords(4)) = True or re.Test(keyWords(5)) = True or re.Test(keyWords(6)) = True or re.Test(keyWords(7)) = True Then
			For each token in split(FetchClipboardData)
				For i=1 To len(token)
					'If token is a number, paste upper and lower range values in spreadsheet. 
					If mid(token,i,1) = "0" or mid(token,i,1) = "1" or mid(token,i,1) = "2" or mid(token,i,1) = "3" or mid(token,i,1) = "4" or mid(token,i,1) = "5" or mid(token,i,1) = "6" or mid(token,i,1) = "7" or mid(token,i,1) = "8" or mid(token,i,1) = "9" Then
						objExcel.ActiveCell.Offset(0,-3).Value = 3
						objExcel.ActiveCell.Offset(0,1).Value = token
						objExcel.ActiveCell.Offset(0,2).Value = 1000000
						row = objExcel.ActiveCell.Row
						objExcel.Cells(row+1,2).Select
						Exit For 
					End If 
				Next 			
			Next 
		End If
	End If  
Next 















