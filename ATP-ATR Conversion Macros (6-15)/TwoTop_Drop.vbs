'Write the previous item's Field Type and Sec. Ref values to the current one.
Set objExcel = GetObject(,"Excel.Application") 
Set adjCell = objExcel.ActiveCell.Offset(0,1)
Set targetRng = objExcel.Range(objExcel.ActiveCell,adjCell)
Set srcRng = objExcel.Range(objExcel.ActiveCell.Offset(-1,0),adjCell.Offset(-1,0))

targetRng.Value = srcRng.Value
adjCell.Offset(0,1).Select