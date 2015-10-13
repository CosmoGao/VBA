Option Explicit
 Sub score()
 Range("A:E").Clear
 Cells(1, 1) = "学号"
 Cells(1, 2) = "姓名"
 Cells(1, 3) = "分数"
 Cells(1, 4) = "成绩"
 Cells(1, 5) = "等级"
 Dim arr1(1 To 8)
 arr1(1) = ("赵")
 arr1(2) = ("钱")
 arr1(3) = ("孙")
 arr1(4) = ("李")
 arr1(5) = ("周")
 arr1(6) = ("吴")
 arr1(7) = ("郑")
 arr1(8) = ("王")
 Dim arr2(1 To 9)
 arr2(1) = ("大")
 arr2(2) = ("二")
 arr2(3) = ("三")
 arr2(4) = ("四")
 arr2(5) = ("五")
 arr2(6) = ("六")
 arr2(7) = ("七")
 arr2(8) = ("八")
 arr2(9) = ("九")
 
 Dim ipt%
 ipt = InputBox("请输入模拟数据数量", "模拟数据", 50, 1)
 Dim i As Integer
    For i = 2 To ipt + 1
    Cells(i, 1) = 2015123000 + i - 1
    Range(Cells(2, 1), Cells(i, 1)).Select
    Selection.NumberFormatLocal = "0000000000"

    Cells(i, 2) = arr1(7 * Rnd + 1) & arr2(8 * Rnd + 1)
    Cells(i, 3) = Int(50 * Rnd + 50)
    Dim x As Integer
    Next
    
 x = MsgBox("模拟数据生成完毕,是否进行模拟计算?", vbYesNo, "模拟计算")
 If x = 6 Then
    For i = 2 To ipt + 1
    Dim score As Integer
       score = Cells(i, 3)
    
    If score >= 60 Then
       Cells(i, 4) = "及格"
    Else
       Cells(i, 4) = "不及格"
       End If
    
    If score >= 90 Then
    Cells(i, 5) = "优秀"
    Else
        If score >= 80 Then
        Cells(i, 5) = "良好"
        Else
           If score >= 70 Then
           Cells(i, 5) = "中等"
           Else
               If score >= 60 Then
               Cells(i, 5) = "及格"
               Else
               Cells(i, 5) = "不及格"
               End If
           End If
       End If
    End If
    Next
 Else
    End If
 
 End Sub