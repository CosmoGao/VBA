Option Explicit
'修改自@鬼哥 的代码

Sub 批量插入多选控件()
    Dim i As Long, Cell As Range, CellEnd As Range
    Set Cell = Application.InputBox("请选择标记列的首行单元格", "选择单元格", , , , , , 8)
    Set CellEnd = Application.InputBox("请选择标记列的末行单元格", "选择单元格", , , , , , 8)
    For i = Cell.Row To CellEnd.Row
        ActiveSheet.CheckBoxes.Add(Cell.Left, Cells(i, 1).Top - 3, 24, 24).Select
        With Selection
            .Value = xlOff
            .LinkedCell = Cells(i, Cell.Column).Address
            .Caption = ""
        End With
    Next
    Range(Cell, Cells(ActiveSheet.UsedRange.Rows.Count, Cell.Column)).NumberFormat = ";;;"
End Sub

Sub 删除所有多选控件()
    Sheets(1).CheckBoxes.Delete
End Sub
