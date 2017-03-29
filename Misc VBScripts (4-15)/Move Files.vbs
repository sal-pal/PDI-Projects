Set fso = CreateObject("Scripting.FileSystemObject")
Set Folder = fso.GetFolder("C:\Users\palomis\Desktop\Numeric").Files
Set moriFiles = fso.GetFolder("C:\Users\palomis\Desktop\Mori Machines").Files

For each file in Folder	
	file.move("C:\Users\palomis\Desktop\Mori Machines\")
Next