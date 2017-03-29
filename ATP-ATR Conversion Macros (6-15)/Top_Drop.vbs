'Copy contents from cell directly above and paste to the active cell. 
Set objExcel = GetObject(,"Excel.Application") 
Set activeCell = objExcel.ActiveCell
Set topCell = objExcel.ActiveCell.Offset(-1,0)

activeCell.Value = topCell.Value
activeCell.Offset(0,1).Select