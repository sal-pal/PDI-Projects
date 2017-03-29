'Copy contents from cell directly above and paste to the active cell. Then select
'first cell of next row. 
Set objExcel = GetObject(,"Excel.Application") 
Set activeCell = objExcel.ActiveCell
Set topCell = objExcel.ActiveCell.Offset(-1,0)

activeCell.Value = topCell.Value

row = objExcel.ActiveCell.Row
objExcel.Cells(row+1,2).Select