Option Explicit
Sub 合并汇总表()

'本程序适用于合并多个格式一致的Excel表格至一个汇总表  by 高宇皓

'条件确认
 If MsgBox("请确认要合并的文件与汇总表在同一文件夹，且文件夹内没有其他文件！", vbOKCancel, "重要提示") = vbCancel Then
 Exit Sub
 End If

 '声明变量
 Dim MyPath, MyName, AWbName
 Dim Wb As Workbook
 Dim Num As Integer

 '关闭屏幕更新
    Application.ScreenUpdating = False
 
 '定义目录文件格式及变量值
	MyPath = ActiveWorkbook.Path
    MyName = Dir(MyPath & "\" & "*.xlsx")
    AWbName = ActiveWorkbook.Name
    Num = 0

 '循环处理

 Do While MyName <> ""														'名称判断
    If MyName <> AWbName Then												'除本文件外
        Set Wb = Workbooks.Open(MyPath & "\" & MyName)						'定义变量
        Num = Num + 1
            Dim i As Integer
                Workbooks(AWbName).Activate									'激活汇总表
                    i = Num + 2
                        Cells(i, 1) = i - 2
                        Cells(i, 2) = Wb.Sheets(1).Range("C3")
                        Cells(i, 3) = Wb.Sheets(1).Range("F10")
                        Cells(i, 4) = Wb.Sheets(1).Range("C10")
                        Cells(i, 5) = Wb.Sheets(1).Range("C6")
                        Cells(i, 6) = Wb.Sheets(1).Range("F6")
                        Cells(i, 7) = Wb.Sheets(1).Range("C7")
                        Cells(i, 8) = Wb.Sheets(1).Range("F4")
                        Cells(i, 9) = Wb.Sheets(1).Range("C8")
                        Cells(i, 10) = Wb.Sheets(1).Range("F7")
                        Cells(i, 11) = Wb.Sheets(1).Range("C13")
                        Cells(i, 12) = Wb.Sheets(1).Range("F25")
                        Cells(i, 12).NumberFormatLocal = "yyyy/m/d"
                        Cells(i, 13) = Wb.Sheets(1).Range("C14")
                        Cells(i, 14) = Wb.Sheets(1).Range("F14")
                        Cells(i, 15) = Wb.Sheets(1).Range("C23")
        Wb.Close False														'关闭表
    End If
        MyName = Dir
 Loop

 '开启屏幕更新
 Application.ScreenUpdating = True
 
 '统计合并个数
 MsgBox "共合并了" & Num & "个工作薄表", vbInformation, "提示"

End Sub


