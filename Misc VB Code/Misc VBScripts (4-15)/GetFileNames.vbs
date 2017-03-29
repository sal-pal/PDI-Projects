'GOAL: Get the file names of all files in the Mori Machines folder and write them to 
'a file. 

'Step 1: Get the names of all files in the Mori Machines folder. 
	'Access the files collection object of Mori Machines folder.
	'Iterate over each file and retrieve the file path. 

'Step 2: Write the names to a new file. 
	'Create a new file.
		'CreateTextFile method
		'Write each file path item to the text file. 
		
		


Set fso = CreateObject("Scripting.FileSystemObject") 
Set txtStrm = fso.CreateTextFile("C:\Users\palomis\Desktop\FilePaths.txt")
Set moriFiles = fso.GetFolder("C:\Users\palomis\Desktop\Mori Machines").Files

count = 0
For each file in moriFiles
	count = count + 1 
	txtStrm.writeline(file.path)
Next 

txtStrm.close
MsgBox(count)