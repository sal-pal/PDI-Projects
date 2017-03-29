'PART 1: PARTITION DATA 


'Create three text files to prepare for data partitioning 
REM Set fso = CreateObject("Scripting.FileSystemObject")
Set alphaFile = fso.CreateTextFile("C:\Users\palomis\Desktop\Alpha.txt")
Set numeraFile = fso.CreateTextFile("C:\Users\palomis\Desktop\Numera.txt")
Set discardFile = fso.CreateTextFile("C:\Users\palomis\Desktop\Discard.txt") 


'Trim data of file path info and double quotes 
Set SL150 = fso.OpentextFile("C:\Users\palomis\Desktop\SL150FileList.txt")
Do until SL150.AtEndOfStream
	line = SL150.Readline()
	quotes = mid(line,1,1)
	removedQuotes = Replace(line,quotes,"")
	fileName = Replace(removedQuotes,"P:\CNC Programming\SL150\","")
	
	'Filter and partition data based on the value of the first char after trimming.
	firstTwoChars = mid(fileName,1,2)
	If firstTwoChars = "SL" or firstTwoChars = "sl" or firstTwoChars = "Sl" or firstTwoChars = "sL" Then
		partNum = Replace(fileName,firstTwoChars,"")
		If mid(partNum,1,1) = "0" or mid(partNum,1,1) = "1" or mid(partNum,1,1) = "2" or mid(partNum,1,1) = "3" or mid(partNum,1,1) = "4" or mid(partNum,1,1) = "5" or mid(partNum,1,1) = "6" or mid(partNum,1,1) = "7" or mid(partNum,1,1) = "8" or mid(partNum,1,1) = "9" Then 
			numeraFile.WriteLine(partNum)
		Else
			alphaFile.WriteLine(partNum)	
		End If
	Else															
		discardFile.WriteLine(fileName)
	End If 
Loop

SL150.close
alphaFile.close 
numeraFile.close


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'PART 2: REMOVE ILLEGAL CHARS FROM NUMERA 


Set numeraFile = fso.OpenTextFile("C:\Users\palomis\Desktop\Numera.txt")
Set trimedNumera = fso.CreateTextFile("C:\Users\palomis\Desktop\TrimedNumera.txt")

'Iterate over each line and test string for illegal chars 
Do Until numeraFile.AtEndOfStream
	line = numeraFile.Readline()
	Dim strVar
	For i=1 To len(line)
		charact = mid(line,i,1)
		If charact = "." Then 
			'If char next to period is a dash, remove only the period. Done to preserve a part number containing a dash. 
			If mid(line,(i+1),1) = "-" Then 
				strVar = Replace(line,charact,"") 					
				Exit For
			Else 
				'Remove period along with all chars proceeding it. Done to delete the operation sequence number proceeding the period. 
				strVar = Replace(line,charact & mid(line,(i+1),len(line)-i),"") 					
				Exit For
			End If 
			'If char not an int or a dash, remove it
		ElseIf charact <> "0" and charact <> "1" and charact <> "2" and charact <> "3" and charact <> "4" and charact <> "5" and charact <> "6" and charact <> "7" and charact <> "8" and charact <> "9" and charact <> "-" Then
			strVar = Replace(line,charact,"")												
			Exit For
		End If 
	Next
	
	'Second loop implemented in order to remove more than one char per string iteration. 
	count = 0
	Do While count < (len(line)-1)
		count = count + 1
		For i=1 To len(strVar)
			charact = mid(strVar,i,1)
			If charact <> "0" and charact <> "1" and charact <> "2" and charact <> "3" and charact <> "4" and charact <> "5" and charact <> "6" and charact <> "7" and charact <> "8" and charact <> "9" and charact <> "-" Then
				strVar = Replace(strVar,charact,"")
				Exit For 
			End If 
		Next
	Loop
	trimedNumera.WriteLine(strVar)
Loop

numeraFile.close
trimedNumera.close









