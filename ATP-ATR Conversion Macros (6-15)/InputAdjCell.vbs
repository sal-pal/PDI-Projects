'Input data and select adjacent cell to the right.
inputData = InputBox("Value of the cell")
Set objExcel = GetObject(,"Excel.Application") 
objExcel.ActiveCell = inputData
objExcel.ActiveCell.Offset(0,1).Select

