Set fso = CreateObject("Scripting.FileSystemObject")
Set textStrm = fso.OpenTextFile("C:\Users\palomis\Desktop\FilePaths.txt",1)
Set xlObj = CreateObject("Excel.Application")


Do until textStrm.AtEndOfStream
	filePath = textStrm.readline()
	fileName = Replace(filePath,"C:/Users/palomis/Desktop/ATR_Docs/","")
	
	Set WBObj = xlObj.WorkBooks.Open(filepath)
	Call WBObj.SaveAs("C:\Users\palomis\Desktop\LocallySavedCSV\" & fileName,6)
	
Loop





