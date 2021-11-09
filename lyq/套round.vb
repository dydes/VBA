Sub test()

'目前就是行列循环一下，后面可以先定为到所有数值类的，创建数组之后再循环
For i = 5 To 10
    a = Range("D" & i).Formula
    b = Range("D" & i)
    Debug.Print a, b

    '如果是公式，等于号个数大于0，就替换掉等于号，套round
    If UBound(Split(a, "=")) > 0 Then
        c = Replace(a, "=", "")
        Debug.Print c
        Range("H" & i) = "=ROUND(" & c & ",2)"

    '如果不是公式，还需要在判断一下是不是空，是空就啥也不要
    Else
        Range("H" & i) = "=ROUND(" & b & ",2)"
    End If
Next

End Sub