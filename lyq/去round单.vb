Sub 去round()
'获取最大行列
rowmax = ActiveSheet.UsedRange.Rows.Count
colmax = ActiveSheet.UsedRange.Columns.Count

'目前就是行列循环一下，后面可以先定为到所有数值类的，创建数组之后再循环
For j = 1 To colmax
    For i = 1 To rowmax
        cName = Split(Cells(i, j).Address, "$")(1)
        '判断是否是数字且长度大于0
        If IsNumeric(Range(cName & i)) = True And Len(Range(cName & i).Formula) > 0 Then
                Debug.Print cName & i
                '去掉round和位数2
                a = Replace(Range(cName & i).Formula, "=ROUND(", "=", 1, 1)
                b = Len(a)
                c = Left(a, b - 3)
                Range(cName & i) = c
        End If
    Next
    Application.StatusBar = "整体进度" & GetProgress(j, colmax)
Next
MsgBox "操作完成"
Application.StatusBar = False
End Sub
Function GetProgress(curValue, maxValue)
Dim i As Single, j As Integer, s As String
i = maxValue / 20
j = curValue / i
'进度条
For m = 1 To j
    s = s & "■"
Next
For n = 1 To 20 - j
    s = s & "□"
Next
GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function
