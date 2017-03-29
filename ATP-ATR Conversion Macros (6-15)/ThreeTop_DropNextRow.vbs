'Copy contents from three cells directly above and paste to the target range. Then select
'first cell of next row. 
Set objExcel = GetObject(,"Excel.Application") 
Set lastCell = objExcel.ActiveCell.Offset(0,2)
Set srcRng = objExcel.Range(objExcel.ActiveCell.Offset(-1,0),lastCell.Offset(-1,0))
Set targetRng = objExcel.Range(objExcel.ActiveCell,lastCell)


targetRng.Value = srcRng.Value

row = objExcel.ActiveCell.Row
objExcel.Cells(row+1,2).Select



