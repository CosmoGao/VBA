Sub 意见点前期拆分()

    Dim i, j, k As Integer
    
   '关闭屏幕更新
    Application.ScreenUpdating = False
    
    For i = 2 To 20
        Sheets(i).Name = i - 1
        
        For j = 1 To 16
            Sheets(i).Cells(1, j) = Sheets(1).Cells(1, j)
            
            For k = 2 To 10001
                Sheets(i).Cells(k, j) = Sheets(1).Cells(k + (i - 2) * 10000, j)
            Next
            
        Next
        
    Next

   '开启屏幕更新
    Application.ScreenUpdating = True

End Sub


Sub 意见点合并()

    Dim i, j, k As Integer
    Dim MyPath, MyName, AWbName
    Dim wb As Workbook
    
    MyPath = ActiveWorkbook.Path
    MyName = "更正后的数据-时间"
    AWbName = ActiveWorkbook.Name
    
    Application.ScreenUpdating = False
        
        For i = 1 To 19
            Set wb = Workbooks.Open(MyPath & "\" & MyName & i & ".xls")
            Workbooks(AWbName).Activate
            
            For j = 1 To 100
                
                For k = 2 To 10001
                    Cells(k + 10000 * (i - 1), j) = wb.Sheets(1).Cells(k, j)
                Next
            
            Next
            
            wb.Close False
        Next
    
    Application.ScreenUpdating = True


End Sub
