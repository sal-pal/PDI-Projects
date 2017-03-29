'WRITING A PART NUMBER AND OPSEQ TO ITS SET UP SHEET


Set objXL = CreateObject("Excel.Application")
Set objWrkBook = objXL.WorkBooks.Open("C:\Users\palomis\Desktop\45035.A.xls")

objXL.Range("A2").Value = "PART NUMBER: 45035"
objXL.Range("K2").Value = "OP: A"

objWrkBook.Save 
objXL.quit

