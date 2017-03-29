'CHECK FOR MISMATCH BETWEEN FIELD TYPE AND THE ENTRIES IN COLUMNS F & G.
Set xlObj = GetObject(,"Excel.Application")
row = 2
errors = 0


Do Until xlObj.ActiveSheet.Cells(row,2).Value = ""
	'Check if type 3 items have range info in columns F & G. If no, then flag it down.
	If xlObj.ActiveSheet.Cells(row,2).Value = 3 Then 
		If xlObj.ActiveSheet.Cells(row,6) = "" and xlObj.ActiveSheet.Cells(row,7) = "" or xlObj.ActiveSheet.Cells(row,6) = "" or xlObj.ActiveSheet.Cells(row,7) = "" Then 
			xlObj.ActiveSheet.Range(xlObj.ActiveSheet.Cells(row,1), xlObj.ActiveSheet.Cells(row,7)).Interior.Color = 65535
			errors = errors + 1
		End If 
	'Check if type 2 items don't have range info in columns F & G. If no, then flag it down.
	ElseIf xlObj.ActiveSheet.Cells(row,2).Value = 2 Then 
		If xlObj.ActiveSheet.Cells(row,6) <> "" and xlObj.ActiveSheet.Cells(row,7) <> "" Then 	
			xlObj.ActiveSheet.Range(xlObj.ActiveSheet.Cells(row,1), xlObj.ActiveSheet.Cells(row,7)).Interior.Color = 65535
			errors = errors + 1
		End If
	'Check if type 4 items don't have range info in columns F & G. If no, then flag it down.
	ElseIf xlObj.ActiveSheet.Cells(row,2).Value = 4 Then 
		If xlObj.ActiveSheet.Cells(row,6) <> "" and xlObj.ActiveSheet.Cells(row,7) <> "" Then 	
			xlObj.ActiveSheet.Range(xlObj.Cells(row,1), xlObj.ActiveSheet.Cells(row,7)).Interior.Color = 65535
			errors = errors + 1
		End If	
	'If field type is not 2, 3, or 4, then flag it down. 
	ElseIf xlObj.ActiveSheet.Cells(row,2).Value <> 2 or xlObj.ActiveSheet.Cells(row,2).Value <> 3 or xlObj.ActiveSheet.Cells(row,2).Value <> 4 Then 
		xlObj.ActiveSheet.Range(xlObj.Cells(row,1), xlObj.ActiveSheet.Cells(row,7)).Interior.Color = 65535
		errors = errors + 1
	End If 
	
	row = row + 1
Loop 


If errors = 0 Then 
	Call MsgBox("No field type mismatch errors found.",,"Error Checking Script")
End If 