Sub test()

Sheets("链接对应地址").Select
rowmax = Sheets("链接对应地址").UsedRange.Rows.Count

arr_link = Range("A2:A" & rowmax)
arr_remark = Range("B2:B" & rowmax)

Sheets("数据源").Select
Application.DisplayStatusBar = True

For i = 1 To 23
    For j = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(j, i) = "" Then
            Cells(j, i) = Cells(j - 1, i)
        End If
    Next
    Application.StatusBar = "补齐数据" & GetProgress(i, 23)
Next

For j = 0 To rowmax - 2
    k = arr_link(j + 1, 1)
    l = arr_remark(j + 1, 1)
    Call find_ins("Z", k, l)
    Application.StatusBar = "清洗中" & GetProgress(j, rowmax - 1)
    Application.DisplayStatusBar = True
Next

MsgBox "清洗完成"
End Sub


Function find_ins(a, b, c)
'a是列号，如A；b表示要替换的字符串；c是中文备注

Debug.Print a, b, c
rowmax1 = Sheets("数据源").UsedRange.Rows.Count
col = Range(a & 1).Column

For i = 1 To rowmax1
    If InStr(Range(a & i), b) > 0 Then '假设数据都在A列中，从第一行开始
        Cells(i, col + 1) = b
        Cells(i, col + 2) = c
    End If
Next

End Function

Function GetProgress(curValue, maxValue)
Dim i As Single, j As Integer, s As String
    i = maxValue / 20
    j = curValue / i
    For M = 1 To j
        s = s & "■"
    Next M
    For n = 1 To 20 - j
        s = s & "□"
    Next n
    GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function


