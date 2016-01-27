Sub 读取表格_街道()

    '条件确认
     If MsgBox("请确认要合并的文件与汇总表在同一文件夹，且文件夹内没有其他文件！", vbOKCancel, "重要提示") = vbCancel Then
     Exit Sub
     End If
     
    '声明变量
     Dim MyPath, MyName, AWbName
     Dim wb As Workbook
     Dim Num As Integer
     
    '关闭屏幕更新
     Application.ScreenUpdating = False
    
    '定义目录文件格式及变量值
     MyPath = ActiveWorkbook.Path
     MyName = Dir(MyPath & "\" & "*.xlsx")
     AWbName = ActiveWorkbook.Name
     Num = 0
        
    '循环处理
     Do While MyName <> ""                                                      '名称判断
        
        If MyName <> AWbName Then                                               '除本文件外
            Set wb = Workbooks.Open(MyPath & "\" & MyName)                      '定义变量
            Num = Num + 1
                Dim i As Integer
                    Workbooks(AWbName).Activate                                 '激活汇总表
                    i = Num + 2
                    Cells(i, 1) = Left(MyName, 3)                               '读取街道编码
                            
                    For j = 2 To 19
                        Cells(i, j) = wb.Sheets(1).Cells(j, 6)                  '各项三级指标得分
                    Next
                        
                        Cells(i, 20) = wb.Sheets(1).Cells(2, 4)                 '二级指标
                        Cells(i, 21) = wb.Sheets(1).Cells(14, 4)                '二级指标
                        Cells(i, 22) = wb.Sheets(1).Cells(2, 2)                 '一级指标
                        
            wb.Close False                                                      '关闭表
        End If
            
        MyName = Dir
     Loop
    
     '开启屏幕更新
      Application.ScreenUpdating = True
     
     '统计合并个数
      MsgBox "共合并了" & Num & "个工作薄表", vbInformation, "提示"
    
End Sub

Sub 读取表格_企业()

    '条件确认
     If MsgBox("请确认要合并的文件与汇总表在同一文件夹，且文件夹内没有其他文件！", vbOKCancel, "重要提示") = vbCancel Then
     Exit Sub
     End If
     
    '声明变量
     Dim MyPath, MyName, AWbName
     Dim wb As Workbook
     Dim Num As Integer
    
    '关闭屏幕更新
     Application.ScreenUpdating = False
    
    
    '定义目录文件格式及变量值
     MyPath = ActiveWorkbook.Path
     MyName = Dir(MyPath & "\" & "*.xlsx")
     AWbName = ActiveWorkbook.Name
     Num = 0
         
    '循环处理
     Do While MyName <> ""                                                      '名称判断
        
        If MyName <> AWbName Then                                               '除本文件外
            Set wb = Workbooks.Open(MyPath & "\" & MyName)                      '定义变量
            Num = Num + 1
                Dim i As Integer
                Workbooks(AWbName).Activate                                     '激活汇总表
                i = Num + 2
                Cells(i, 1) = Left(MyName, 3)                                   '读取部门编码
                            
                For j = 2 To 18
                    Cells(i, j) = wb.Sheets(1).Cells(j, 6)                      '各项三级指标得分
                Next
                                
                    Cells(i, 19) = wb.Sheets(1).Cells(2, 4)                     '二级指标
                    Cells(i, 20) = wb.Sheets(1).Cells(9, 4)                     '二级指标
                    Cells(i, 21) = wb.Sheets(1).Cells(2, 2)                     '一级指标
                
                For k = 22 To 31
                    Cells(i, k) = wb.Sheets(1).Cells(k - 13, 5)                 '三级指标名称
                Next
                
            wb.Close False                                                      '关闭表
        End If
        
        MyName = Dir
     Loop
    
    '开启屏幕更新
     Application.ScreenUpdating = True
     
    '统计合并个数
     MsgBox "共合并了" & Num & "个工作薄表", vbInformation, "提示"
       
End Sub



Sub 分解得分表_企业()
    
    Dim wb, awb As Workbook
    Dim i, j As Integer
    AWbName = ActiveWorkbook.Name
    Set awb = Workbooks(AWbName)
    For i = 2 To 18
        Set wb = Workbooks.Open("C:\Users\GaoYH\Desktop\东城\Database\数据\" & Cells(i, 1) & "满意得分.xlsx")
            
        For j = 14 To 19
            wb.Sheets(1).Cells(j, 8) = awb.Sheets(1).Cells(i, j - 10)
        Next
        
        wb.Save
        wb.Close False
    Next
    
End Sub



Sub 分解得分表_居民()
    
    Dim wb, awb As Workbook
    Dim i, j As Integer
    
    AWbName = ActiveWorkbook.Name
    Set awb = Workbooks(AWbName)
    
    For i = 2 To 18
        Set wb = Workbooks.Open("C:\Users\GaoYH\Desktop\东城\Database\数据\" & Cells(i, 1) & "满意得分.xlsx")
        
        For j = 2 To 13
            wb.Sheets(1).Cells(j, 8) = awb.Sheets(1).Cells(i, j + 2)
        Next
    
        wb.Save
        wb.Close False
    Next
    
End Sub


