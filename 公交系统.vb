Option Explicit
Sub Step2_1()
    
    '月度表导出到各个公司表格
    '此处打开某月意见率和某月意见点及模板中所有工作簿
    
    Dim i As Integer
    Dim month As String
    month = InputBox("请输入月份", "月份", "月")
    For i = 1 To 11
        
        Workbooks(month & "意见点.xlsx").Activate

        '复制到相应公司表
        Sheets(i).Copy before:=Workbooks("服务热线信息统计 - " & Sheets(i).[b1] & ".xlsx").Sheets(2)
        
        Workbooks(month & "意见率.xlsx").Activate

        Sheets(i).Copy before:=Workbooks("服务热线信息统计 - " & Sheets(i).[b1] & ".xlsx").Sheets(2)
        
        '定位单元格
        Sheets(1).Select
        
        [a1].Select
        
    Next
    
    '可以关闭某月意见率和意见点工作簿，对模板中的工作簿逐个处理
    
End Sub

Sub Step2_2()
    
    '读取公交线路，删除多余列
    
    '关闭屏幕更新
    Application.ScreenUpdating = False
    
    Dim i, n As Integer
    
    n = Application.CountA(Sheets(2).Range("4:4")) - 2
    
    
    For i = 1 To n
        
        '读取公交线路
        Sheets(1).Cells(3, 6 * i - 4) = Sheets(2).Cells(4, i + 1)
        
    Next

   
    '删除多余列
    Range(Columns(6 * n + 2), Columns(1801)).Delete

    '开启屏幕更新
    Application.ScreenUpdating = True

    [a1] = "线路已读取"
        [a1].Select
            With Selection.Interior
                .Pattern = xlSolid
                PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

End Sub


Sub Step2_3()
 
    '关闭屏幕更新
    Application.ScreenUpdating = False
    
    Dim i, j, n As Integer
    
    n = Application.CountA(Sheets(2).Range("4:4")) - 2
    
    For i = 7 To 88
        
        For j = 1 To n
        
            '意见率读取
            Cells(i, 6 * j - 4) = Application.IfNa(Application.VLookup(Cells(i, 1), Sheets(2).Cells, Application.Match(Cells(3, 6 * j - 4), Sheets(2).[4:4], 0), 0), 0)
            
            '意见点读取
            Cells(i, 6 * j - 1) = Application.IfNa(Application.VLookup(Cells(i, 1), Sheets(3).Cells, Application.Match(Cells(3, 6 * j - 4), Sheets(3).[4:4], 0), 0), 0)
        
        Next
    Next
    
    '开启屏幕更新
    Application.ScreenUpdating = True
    
    [a1] = "意见已查询"
        [a1].Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

End Sub


Sub Step2_4()
    
    '求和、百分比
    
    '关闭屏幕更新
    Application.ScreenUpdating = False
    
    Dim i, j, n As Integer
    
    n = Application.CountA(Range("5:5"))
    
    For i = 1 To n
    
        '第一部分，更改件数列求和公式
        
        '服务
        
            '投诉
            Cells(8, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(9, i * 3 - 1), Cells(14, i * 3 - 1)))
            
            '服务建议
            Cells(15, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(16, i * 3 - 1), Cells(30, i * 3 - 1)))
            
            '表扬
            Cells(31, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(32, i * 3 - 1), Cells(37, i * 3 - 1)))
        
        Cells(7, i * 3 - 1) = Cells(8, i * 3 - 1) + Cells(15, i * 3 - 1) + Cells(31, i * 3 - 1)
        
        '安全
        
            '投诉
            Cells(39, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(40, i * 3 - 1), Cells(46, i * 3 - 1)))
            
            '安全建议
            Cells(47, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(48, i * 3 - 1), Cells(56, i * 3 - 1)))
        
        Cells(38, i * 3 - 1) = Cells(39, i * 3 - 1) + Cells(47, i * 3 - 1)
        
        '运营
        
            '投诉
            Cells(58, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(59, i * 3 - 1), Cells(62, i * 3 - 1)))
            
            '运营建议
            Cells(63, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(64, i * 3 - 1), Cells(77, i * 3 - 1)))
        
        Cells(57, i * 3 - 1) = Cells(58, i * 3 - 1) + Cells(63, i * 3 - 1)
        
        '技术
            
            '技术投诉
            Cells(79, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(80, i * 3 - 1), Cells(81, i * 3 - 1)))
        
            '技术建议
            Cells(82, i * 3 - 1) = Application.WorksheetFunction.Sum(Range(Cells(83, i * 3 - 1), Cells(88, i * 3 - 1)))

            
        Cells(78, i * 3 - 1) = Cells(79, i * 3 - 1) + Cells(82, i * 3 - 1)
        
        
        '第二部分，更改百分比列计算公式
        
        For j = 7 To 88
            
            Cells(j, i * 3 + 1) = Cells(j, i * 3 - 1) / (Cells(7, i * 3 - 1) + Cells(38, i * 3 - 1) + Cells(57, i * 3 - 1) + Cells(78, i * 3 - 1))
            
        Next
    
    Next
    
    '开启屏幕更新
    Application.ScreenUpdating = True
    
    [a1] = "意见率已计算"
        [a1].Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

End Sub


Sub Step2_5()
    
    '定义文件标题
    
    '关闭屏幕更新
    Application.ScreenUpdating = False
    
    Dim i, n As Integer
    Dim year, month As String
    year = InputBox("请输入数据年度", "年度", "2015")
    month = InputBox("请输入数据时间点", "时间点", "月")
    n = Application.CountA(Sheets(2).[4:4]) - 2
    Sheets("年份").[a1] = year
    Sheets("时间点").[a1] = month
    Sheets(1).[b1] = Sheets(2).[b1] & "服务热线数据——" & Sheets("年份").[a1] & "年" & Sheets("时间点").[a1]
    Sheets(1).[b2] = Sheets(2).[b1] & "公司"
    
    Sheets("类型").Select
    
    [b2] = Sheets(2).[b1]
    [c:c].Clear
    [c1] = "线路"
    For i = 2 To n + 1
        
        Cells(i, 3) = Sheets(2).Cells(4, i)
        
    Next
    
    [d2] = n
    Sheets("数据行列数").[b2] = n * 6
    Sheets(2).Delete
    Sheets(2).Delete
    Sheets(1).Select
    
    [a1].Select
     '开启屏幕更新
     Application.ScreenUpdating = True
     
    [a1] = "指标"
        [a1].Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
End Sub