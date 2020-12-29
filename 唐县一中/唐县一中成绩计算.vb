Public 原始分数_rowmax, 原始分数_colmax, arr_sub, arr_col_pos, arr_region
Sub data()

    tim1 = Timer
    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知

'获取原始分数表最大行数列数
    原始分数_rowmax = Sheets("原始分数").UsedRange.Rows.Count
    原始分数_colmax = Sheets("原始分数").UsedRange.Columns.Count

'定义自定义区域名称数组
    arr_region = Array("bc", "score", _
                        "_f1", "_f2", "_f3", "_f4", "_f5", "_f6", "_f7", "_f8", "_f9", _
                        "mc", "mc1", "mc2", "mc3", "mc4", "mc5", "mc6", "mc7", "mc8", "mc9", _
                        "bmc", "bmc1", "bmc2", "bmc3", "bmc4", "bmc5", "bmc6", "bmc7", "bmc8", "bmc9")

'清除已有并新建名称
    Call dele_new_vars
    
'新建可调节参数的名称
    ActiveWorkbook.Names.Add Name:="pjrs", RefersToR1C1:="='平均成绩一览'!R2C7:R2C7"
    ActiveWorkbook.Names.Add Name:="yxs", RefersToR1C1:="='优秀生分布'!R2C7:R2C7"
    ActiveWorkbook.Names.Add Name:="fd", RefersToR1C1:="='分数段统计'!R2C11:R2C11"
    ActiveWorkbook.Names.Add Name:="dxf", RefersToR1C1:="='分数段统计'!R2C14:R2C14"
   
'创建班级参数矩阵，用于后面的计算
    Sheets("使用说明").Range("A7") = "班级"
    Sheets("使用说明").Range("B7") = "起始行号"
    Sheets("使用说明").Range("C7") = "结束行号"
    
'将班级内容复制到参数表中
    Sheets("原始分数").Select
    Range("C2:C" & 原始分数_rowmax).Copy
    Sheets("使用说明").Select
    Range("A8").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
'找到各班起始行号，填写结束行号
    使用说明_rowmax = Sheets("使用说明").UsedRange.Rows.Count
    For i = 8 To 使用说明_rowmax
        If Range("A" & i) <> Range("A" & i - 1) Then
            Range("B" & i) = i - 6
        End If
    Next
    
'去重，得到不重复的班号
    ActiveSheet.Range("$A$8:$C$" & 原始分数_rowmax + 7).RemoveDuplicates Columns:=1, Header:=xlNo
    
'填写结束行号
    使用说明_rowmax_distinct = Sheets("使用说明").UsedRange.Rows.Count
    For i = 8 To 使用说明_rowmax_distinct
        Range("C" & i) = Range("B" & i + 1) - 1
    Next
    Range("C" & 使用说明_rowmax_distinct) = 原始分数_rowmax

'制作各科班名次()
    Sheets("原始分数").Select
    For i = 4 To 13
        Call crange_sort(i, 1)
        Call crange_sort(3, 0)
        For k = 8 To 使用说明_rowmax_distinct
            a = Sheets("使用说明").Range("B" & k)
            b = Sheets("使用说明").Range("C" & k)
            For j = 2 To 原始分数_rowmax
                If Cells(j, i) <> "" Then
                    Cells(j, i + 20) = WorksheetFunction.Rank(Cells(j, i), Range(Cells(a, i), Cells(b, i)))
                End If
            Next
        Next
    Next
    
''定义bmc(f1)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("E2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("Y2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("Y3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-22]=R[-1]C[-22],R[-1]C+1,1)"
'    Range("Y3").Select
'    Selection.AutoFill Destination:=Range("Y3:Y" & 原始分数_rowmax)
'    Range("Y2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f2)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("F2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("Z2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("Z3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-23]=R[-1]C[-23],R[-1]C+1,1)"
'    Range("Z3").Select
'    Selection.AutoFill Destination:=Range("Z3:Z" & 原始分数_rowmax)
'    Range("Z2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f3)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("G2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AA2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AA3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-24]=R[-1]C[-24],R[-1]C+1,1)"
'    Range("AA3").Select
'    Selection.AutoFill Destination:=Range("AA3:AA" & 原始分数_rowmax)
'    Range("AA2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f4)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("H2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AB2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AB3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-25]=R[-1]C[-25],R[-1]C+1,1)"
'    Range("AB3").Select
'    Selection.AutoFill Destination:=Range("AB3:AB" & 原始分数_rowmax)
'    Range("AB2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f5)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("I2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AC2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AC3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-26]=R[-1]C[-26],R[-1]C+1,1)"
'    Range("AC3").Select
'    Selection.AutoFill Destination:=Range("AC3:AC" & 原始分数_rowmax)
'    Range("AC2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f6)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("J2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AD2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AD3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-27]=R[-1]C[-27],R[-1]C+1,1)"
'    Range("AD3").Select
'    Selection.AutoFill Destination:=Range("AD3:AD" & 原始分数_rowmax)
'    Range("AD2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f7)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("K2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AE2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AE3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-28]=R[-1]C[-28],R[-1]C+1,1)"
'    Range("AE3").Select
'    Selection.AutoFill Destination:=Range("AE3:AE" & 原始分数_rowmax)
'    Range("AE2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f8)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("L2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AF2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AF3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-29]=R[-1]C[-29],R[-1]C+1,1)"
'    Range("AF3").Select
'    Selection.AutoFill Destination:=Range("AF3:AF" & 原始分数_rowmax)
'    Range("AF2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'
''定义bmc(f9)
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("C2"), order1:=xlAscending, Key2:= _
'        Range("M2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
'    Range("AG2").Select
'    ActiveCell.FormulaR1C1 = "1"
'    Range("AG3").Select
'    ActiveCell.FormulaR1C1 = "=IF(RC[-30]=R[-1]C[-30],R[-1]C+1,1)"
'    Range("AG3").Select
'    Selection.AutoFill Destination:=Range("AG3:AG" & 原始分数_rowmax)
'    Range("AG2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
    
''回到原记录顺序
''
'    Range("A1:AG" & 原始分数_rowmax).Sort key1:=Range("A2"), order1:=xlAscending, Header:=xlYes, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, _
'        DataOption1:=xlSortNormal


'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True '重启刷新
    Application.DisplayAlerts = True '重启通知
    MsgBox "计算完成，用时" & Format(using_time, "0.0秒")
    
End Sub

Function dele_new_vars()
'根据名称数组，逐个删除已有名称，新建对应名称的区域
    For i = 0 To UBound(arr_region)
        ActiveWorkbook.Names(arr_region(i)).Delete
        ActiveWorkbook.Names.Add Name:=arr_region(i), RefersToR1C1:="='原始分数'!R2C" & i + 3 & ":R" & 原始分数_rowmax & "C" & i + 3
    Next
End Function

Function sub_col()
'判断科目是否存在，存在记录具体列号，不存在记0
    arr_col_pos = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0) '创建一个用于接收列号的空数组
    For i = 0 To UBound(arr_sub) '逐个学科查找
        Sheets("原始分数").Select
        col = Application.Match(arr_sub(i), Range(Cells(1, 1), Cells(1, 原始分数_colmax)), 0)
            If IsNumeric(col) = True Then '如果科目名称存在，就记列号
                arr_col_pos(i) = col
            Else
                arr_col_pos(i) = 0 '如果不存在，列号为0
            End If
    Next
End Function

Function crange_sort(key, order) '三个参数：ref表示排序的区域，随便给个A1就行，key表示排序的关键字是哪个字段，order0升序，1降序
'如果是希望排多个字段，需要把权重最高的放在最后
    If order = 0 Then '升序
        Range("A1").CurrentRegion.Sort key1:=Cells(1, key), order1:=xlAscending, Header:=xlYes
    ElseIf order = 1 Then '降序
        Range("A1").CurrentRegion.Sort key1:=Cells(1, key), order1:=xlDescending, Header:=xlYes
    ElseIf order <> 0 Or order <> 1 Then
        Exit Function
    End If
End Function
