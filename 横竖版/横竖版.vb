Sub 横竖版()

'得处理一下缺考
'得校验一下考试名称是否填写
'得校验一下总表下面是否有汇总

'获取最大行数
    Sheets("总表").Select
    rowmax = ActiveSheet.UsedRange.Rows.Count
    
'DEF三列转为数值格式
    For i = 1 To rowmax
        Sheets("总表").Range("G" & i) = Range("D" & i).Value
        Sheets("总表").Range("H" & i) = Range("E" & i).Value
        Sheets("总表").Range("I" & i) = Range("F" & i).Value
    Next

'删除DEF三列
    Sheets("总表").Select
    Range("D:F").Select
    Selection.Delete Shift:=xlToLeft
    
'按班级升序，按总分降序
    Sheets("总表").Select
    Columns("A:F").Select
    ActiveWorkbook.Worksheets("总表").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("总表").Sort.SortFields.Add2 Key:=Range("B2:B" & rowmax), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("总表").Sort.SortFields.Add2 Key:=Range("D2:D" & rowmax), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("总表").Sort
        .SetRange Range("A1:F" & rowmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'JKL列增加辅助公式
    For i = 2 To rowmax
        Sheets("总表").Range("G" & i).Formula = "=IF(B" & i & "=B" & i - 1 & ","""",1)"
        If Sheets("总表").Range("G" & i) = 1 Then
            Sheets("总表").Range("H" & i) = i
        End If
    Next
    
'复制班级列粘贴到说明sheet中并去重
    Sheets("总表").Select
    Range("B2:B" & rowmax).Select
    Selection.Copy
    Sheets("说明").Select
    Range("A5").Select
    ActiveSheet.Paste
    Sheets("总表").Select
    Range("H2:H" & rowmax).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("说明").Select
    Range("B5").Select
    ActiveSheet.Paste
    ActiveSheet.Range("$A$4:$C$" & rowmax).RemoveDuplicates Columns:=1, Header:=xlYes
    Sheets("说明").Range("A4") = "班级信息"
    Sheets("说明").Range("B4") = "起始行数"
    Sheets("说明").Range("C4") = "结束行数"

'生成各班起始行数与结束行数
'    Sheets("说明").Range("B5") = 2
'    rowmax1 = ActiveSheet.UsedRange.Rows.Count
'    For i = 5 To rowmax1
'        j = Sheets("说明").Range("A" & i)
'            For k = 2 To rowmax
'                If Sheets("总表").Range("B" & k) = j Then
'
'
End Sub

