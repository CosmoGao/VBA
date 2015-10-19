Sub 计算得分()
'
'该宏应用于将SPSS导出的等距量表转换为百分制，并计算加权平均数
'高宇皓 编写

'将B列复制到G列
    Columns("B:B").Select
    Selection.Copy
    Columns("G:G").Select
    Selection.PasteSpecial
    Application.CutCopyMode = False
    Selection.ClearFormats
    Range("G:G,I:I,J:J,K:K,L:L").Select
    Selection.Clear

'对H列进行替换
    Columns("H:H").Select

'以下内容权重赋值为1
    Selection.Replace What:="很好", Replacement:="1", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="很重视", Replacement:="1", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="完善", Replacement:="1", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="很有效", Replacement:="1", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="有力度", Replacement:="1", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="好", Replacement:="1", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'以下内容权重赋值为0.8
    Selection.Replace What:="较好", Replacement:="0.8", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="比较重视", Replacement:="0.8", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="比较完善", Replacement:="0.8", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="比较有效", Replacement:="0.8", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="力度较大", Replacement:="0.8", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
  
'以下内容权重赋值为0.6
    Selection.Replace What:="一般", Replacement:="0.6", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 
'以下内容权重赋值为0.4
    Selection.Replace What:="不太好", Replacement:="0.4", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="不太重视", Replacement:="0.4", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="不太完善", Replacement:="0.4", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="不太有效", Replacement:="0.4", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="力度较小", Replacement:="0.4", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'以下内容权重赋值为0.2
    Selection.Replace What:="不好", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="不重视", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="不完善", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="没效果", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="没有力度", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="不了解", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="没听说", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="不好说", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="未开展工作", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="本单位不设纪委（纪工委、纪检组）", Replacement:="0.2", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'清除合计项
    Selection.Replace What:="合计", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'根据权重进行相乘
    For i = 1 To 300
   '    If Cells(i, 8) = "" Then Cells(i, 9) = ""
        If Cells(i, 8) <> "" Then Cells(i, 9) = Cells(i, 8) * Cells(i, 4)
    Next

'根据题号进行求和，求和向下的六项
    For i = 4 To 200
        If Cells(i, 1) Like "A*" Then Cells(i, 10) = Application.WorksheetFunction.Sum(Range(Cells(i + 2, 9), Cells(i + 7, 9)))
    Next

'若果求和结果大于100，说明产生叠加，改变求和公式为求和向下的4项
    For j = 4 To 200
        If Cells(j, 10) > 100 Then Cells(j, 10) = Application.WorksheetFunction.Sum(Range(Cells(j + 2, 9), Cells(j + 5, 9)))
        Next
    
'对Sheet重命名
    Sheets(1).Name = "导出结果及计算"
    Sheets.Add After:=Sheets(1)
    Sheets(2).Name = "计算结果"

'纪检组织选择框
    Dim x As Integer
    x = MsgBox("是否有纪检组织", vbYesNo, "选择")
    If x = 6 Then
    GoTo a:
    Else
    GoTo b:

'有纪检组织的单位
a:
    Workbooks.Open ("E:/OneDrive/CMMR/未归挡项目/2015丰台区党风廉政建设民意调查项目资料/计算/指标体系和权重定稿.xlsx")
    ActiveWindow.WindowState = xlMinimized
    Workbooks("指标体系和权重定稿.xlsx").Sheets(3).Range("A1:I18").Copy
    Sheets(2).Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
       , SkipBlanks:=False, Transpose:=False
    For i = 1 To 200
    If Sheets(1).Cells(i, 1) Like "A4*" Then Sheets(2).Cells(4, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*①*" Then Sheets(2).Cells(5, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*②*" Then Sheets(2).Cells(6, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*③*" Then Sheets(2).Cells(7, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*④*" Then Sheets(2).Cells(8, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*⑤*" Then Sheets(2).Cells(9, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A9*" Then Sheets(2).Cells(10, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A10*" Then Sheets(2).Cells(11, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A2*" Then Sheets(2).Cells(12, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A3*" Then Sheets(2).Cells(13, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A7*" Then Sheets(2).Cells(14, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A6*" Then Sheets(2).Cells(15, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A8*" Then Sheets(2).Cells(16, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A12*" Then Sheets(2).Cells(17, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A13*" Then Sheets(2).Cells(18, 9) = Sheets(1).Cells(i, 10)
        Next
   
    Workbooks("指标体系和权重定稿.xlsx").Close
    End
     
    
 '无纪检组织的单位
b:
    Workbooks.Open ("E:/OneDrive/CMMR/未归挡项目/2015丰台区党风廉政建设民意调查项目资料/计算/指标体系和权重定稿.xlsx")
    ActiveWindow.WindowState = xlMinimized
    Workbooks("指标体系和权重定稿.xlsx").Sheets(4).Range("A1:I17").Copy
    Sheets(2).Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
       , SkipBlanks:=False, Transpose:=False
    For i = 1 To 200
    If Sheets(1).Cells(i, 1) Like "A4*" Then Sheets(2).Cells(4, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*①*" Then Sheets(2).Cells(5, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*②*" Then Sheets(2).Cells(6, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*③*" Then Sheets(2).Cells(7, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*④*" Then Sheets(2).Cells(8, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A5*⑤*" Then Sheets(2).Cells(9, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A10*" Then Sheets(2).Cells(10, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A2*" Then Sheets(2).Cells(11, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A3*" Then Sheets(2).Cells(12, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A7*" Then Sheets(2).Cells(13, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A6*" Then Sheets(2).Cells(14, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A8*" Then Sheets(2).Cells(15, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A12*" Then Sheets(2).Cells(16, 9) = Sheets(1).Cells(i, 10)
    If Sheets(1).Cells(i, 1) Like "A13*" Then Sheets(2).Cells(17, 9) = Sheets(1).Cells(i, 10)
        Next
    
    Workbooks("指标体系和权重定稿.xlsx").Close
    
    End If
    
    Sheets(2).Range("A1").Select
    
End Sub
