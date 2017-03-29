Set fso = CreateObject("Scripting.FileSystemObject")
Set txtStrm = fso.OpenTextFile("C:\Users\palomis\Desktop\KeepingTrack.txt",1)

oldDate = txtStrm.ReadLine()
quant = txtStrm.ReadLine()
txtStrm.close

Set txtStrm = fso.OpenTextFile("C:\Users\palomis\Desktop\KeepingTrack.txt",2)




If CDate(oldDate) = Date() Then
	newQuant = CInt(quant)+1
	txtStrm.Write(oldDate & vbCrlf)
	txtStrm.Write(newQuant & vbCrlf)
	txtStrm.Write(Time)
	
	If newQuant >= 15 Then
		Call MsgBox("Today's daily goal is met.",,"Keeping Track")
	ElseIf newQuant = "4" Then 
		Call MsgBox("1st quarter's goal is met.",,"Keeping Track")
	ElseIf newQuant = "8" and newQuant Mod 4 = 0 Then 
		Call MsgBox("2nd quarter's goal is met.",,"Keeping Track")
	ElseIf newQuant = "12" and newQuant Mod 4 = 0 Then 
		Call MsgBox("3rd quarter's goal is met.",,"Keeping Track")
	End If 

Else 
	txtStrm.Write(Date() & vbCrlf)
	txtStrm.Write(1 & vbCrlf)
	txtStrm.Write(Time)
End If 

txtStrm.close


Set shellObj = CreateObject("WScript.Shell")
shellObj.Run("C:\Users\palomis\Desktop\ATP'sLeft.pyw")





