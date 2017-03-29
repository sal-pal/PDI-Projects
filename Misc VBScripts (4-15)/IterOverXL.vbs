'GOAL: Modify each XL file whose file path is provided in the given text file



Set fso = CreateObject("Scripting.FileSystemObject")
Set txtStrm = fso.OpenTextFile("C:\Users\palomis\Desktop\FilePaths.txt",1)
Set objXL = CreateObject("Excel.Application")


Do until txtStrm.AtEndOfStream
	
	Set objXL = CreateObject("Excel.Application")
	Set objWrkBook = objXL.WorkBooks.Open(txtStrm.ReadLine)
	
	with objXL 
		.Range("A2").Value = "PART NUMBER: 45035"
		.Range("K2").Value = "OP: A"
	End With 
	
	with objWrkBook
		.Save 
		.close
	End With 
	
	
	objXL.quit
	
Loop 


txtStrm.close