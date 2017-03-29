'Input data and select first cell of next row. 
inputData = InputBox("Value of the cell")
Set objExcel = GetObject(,"Excel.Application") 
objExcel.ActiveCell = inputData

row = objExcel.ActiveCell.Row
objExcel.Cells(row+1,2).Select