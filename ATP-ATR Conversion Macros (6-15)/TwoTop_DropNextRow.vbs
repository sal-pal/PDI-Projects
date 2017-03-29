'Copy contents from two cells directly above and paste to the target range. Then select
'first cell of next row. 
Set objExcel = GetObject(,"Excel.Application") 
Set topCells = objExcel.ActiveCell.Offset(0,1)
Set targetRng = objExcel.Range(objExcel.ActiveCell,topCells)
Set srcRng = objExcel.Range(objExcel.ActiveCell.Offset(-1,0),topCells.Offset(-1,0))

targetRng.Value = srcRng.Value

row = objExcel.ActiveCell.Row
objExcel.Cells(row+1,2).Select