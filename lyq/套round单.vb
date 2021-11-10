Sub 套round1()
'获取最大范围》行列循环》单元格呈现是否为数字》Y套round（其中公式需要特殊处理）
                                        'N不动（包括空与文本均不动）
'获取最大行列
rowmax = ActiveSheet.UsedRange.Rows.Count
colmax = ActiveSheet.UsedRange.Columns.Count

'目前就是行列循环一下，后面可以先定为到所有数值类的，创建数组之后再循环
For j = 1 To colmax
    For i = 1 To rowmax
        cName = Split(Cells(i, j).Address, "$")(1)
        a = Range(cName & i).Formula
        b = Range(cName & i)
        '如果公式结果是数字的，就继续判断，如果不是的，就不管
        If IsNumeric(b) = True Then
            '如果是公式，等于号个数大于0，就替换掉第一个等于号，套round
            If UBound(Split(a, "=")) > 0 Then
                c = Replace(a, "=", "", 1, 1)
                Range(cName & i) = "=ROUND(" & c & ",2)"
            '不为空是纯数字，就套round
            ElseIf IsNumeric(a) = True Then
                Range(cName & i) = "=ROUND(" & b & ",2)"
            End If
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


