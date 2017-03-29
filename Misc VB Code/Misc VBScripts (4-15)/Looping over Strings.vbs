'HOW TO ITERATE OVER CHARS IN A STRING TO CHECK IF THEY MEET A CONDITION.


str = inputbox("Enter string")
For i=1 to len(str)
	symbol = mid(str,i,1)
	if symbol = "e" Then								
		dest = "C:\Users\palomis\Desktop\efile.txt"
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set file = fso.CreateTextFile(dest)
		file.WriteLine("You caused a text file to be created" & vbCrlf & "but now, I have succeeded in creating a new line!!")
		file.close
		Exit For 
	else
		msgbox(symbol)
	end if 
Next


'Knowing how to run an application through vbscript. 
Set newfso = CreateObject("WScript.Shell")
newfso.run "sbclient.exe sbclient 925"











