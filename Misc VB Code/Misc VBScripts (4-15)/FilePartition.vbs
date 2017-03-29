'FIRST ROUND OF FILE PARTITIONING: BASED ON THE TYPE OF THE FILE NAME'S FIRST CHARACTER  


Set fso = CreateObject("Scripting.FileSystemObject")
Set moriFolder = fso.GetFolder("C:\Users\palomis\Desktop\Mori Machines").Files

For each file in moriFolder		
	If file.type <> "Microsoft Excel 97-2003 Worksheet" Then 
		file.delete
	Else
		If mid(file.name,1,1) = "0" or mid(file.name,1,1) = "1" or mid(file.name,1,1) = "2" or mid(file.name,1,1) = "3" or mid(file.name,1,1) = "4" or mid(file.name,1,1) = "5" or mid(file.name,1,1) = "6" or mid(file.name,1,1) = "7" or mid(file.name,1,1) = "8" or mid(file.name,1,1) = "9" Then 
			file.move("C:\Users\palomis\Desktop\Numeric\")
		Else
			file.move("C:\Users\palomis\Desktop\Alphabetic\")
		End If
	End If
Next


