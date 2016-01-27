Option Explicit
Sub 生成意见()

Dim lc(), td(), hj(), js(), fy()
Dim a, b, c, d, e As Integer
a = Application.WorksheetFunction.CountIf(Sheets(5).Range("A:A"), Sheets(4).Range("i3"))
b = Application.WorksheetFunction.CountIf(Sheets(5).Range("A:A"), Sheets(4).Range("i4"))
c = Application.WorksheetFunction.CountIf(Sheets(5).Range("A:A"), Sheets(4).Range("i5"))
d = Application.WorksheetFunction.CountIf(Sheets(5).Range("A:A"), Sheets(4).Range("i6"))
e = Application.WorksheetFunction.CountIf(Sheets(5).Range("A:A"), Sheets(4).Range("i7"))

ReDim lc(1 To a), td(1 To b), hj(1 To c), js(1 To d), fy(1 To e)
Dim i%, ai%, bi%, ci%, di%, ei%, n%

For di = 1 To d
js(di) = Sheets(5).Cells(di, 2)
Next
For ci = 1 To c
hj(ci) = Sheets(5).Cells(ci + d, 2)
Next
For bi = 1 To b
td(bi) = Sheets(5).Cells(bi + c + d, 2)
Next
For ai = 1 To a
lc(ai) = Sheets(5).Cells(ai + b + c + d, 2)
Next
For ei = 1 To e
fy(ei) = Sheets(5).Cells(ei + a + b + c + d, 2)
Next


n = Application.WorksheetFunction.CountIf(Range("A:A"), "*")
Range(Cells(4, 6), Cells(n + 1, 6)).Clear
For i = 4 To n + 1
    If Cells(i, 2) = Range("i3") Then
        Cells(i, 6) = lc(Int((a - 1) * Rnd()) + 1)
        Else
            If Cells(i, 2) = Range("i4") Then
                Cells(i, 6) = td(Int((b - 1) * Rnd()) + 1)
                Else
                    If Cells(i, 2) = Range("i5") Then
                        Cells(i, 6) = hj(Int((c - 1) * Rnd()) + 1)
                        Else
                            If Cells(i, 2) = Range("i6") Then
                                Cells(i, 6) = js(Int((d - 1) * Rnd()) + 1)
                                Else
                                    If Cells(i, 2) = Range("i7") Then
                                        Cells(i, 6) = fy(Int((e - 1) * Rnd()) + 1)
                                        Else
                                    End If
                            End If
                    End If
            End If
    End If
Next

End Sub

