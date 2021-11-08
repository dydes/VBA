Sub D_R()

Application.DisplayAlerts = False '暂停通知
rowmax = ActiveSheet.UsedRange.Rows.Count
colmax = ActiveSheet.UsedRange.Columns.Count

'逐列去重并降序
rowmax1 = ActiveSheet.UsedRange.Rows.Count
For i = 1 To colmax
    col_n = Split(Cells(2, i).Address, "$")(1)
    j = col_n & "2:" & col_n & rowmax1
    ActiveSheet.Range(Cells(1, i), Cells(rowmax, i)).RemoveDuplicates Columns:=1, Header:=xlYes
    Call rank(i, j)
Next

Application.DisplayAlerts = True '重启通知

End Sub


Function rank(a, rng)

    Columns(a).Select
    b = ActiveSheet.Name
    ActiveWorkbook.Worksheets(b).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(b).Sort.SortFields.Add2 Key:=Range(rng), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(b).Sort
        .SetRange Range(rng)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Function
