Public ofile, ofName, cfile, cfName, 成绩排名_colmax, 成绩排名_rowmax, 参数_rowmax, 参数_rowmax_distinct, arr_sub
Sub 等级赋分()
    
    '读取相关参数

    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知
    
    '调用选文件函数，输出ofile是带路径的文件名，ofName是不带路径的文件名
    Call file_open_name("请选择年级全科文件", "D:\会通\VBA\等级赋分\")
    
    '调用保存文件函数，，输出cfile是带路径的文件夹名，cfName是不带路径的文件夹名
    Call file_save_name("请选择要保存的文件夹", "D:\会通\VBA\等级赋分\")
    
    '打开文件
    Workbooks.Open (ofile)
    
    '获取当前时间
    dat = Format(Date, "yyyy年mm月dd日") '当前年月日
    tim = Format(Time, "hh时mm分ss秒") '当前时间
    tim1 = Timer
    
    '选择“成绩排名”工作表，复制
    Windows(ofName).Activate
    Sheets("成绩排名").Select
    Sheets("成绩排名").Copy
    
    '新建文件并保存新文件
    new_file = cfile & "\等级赋分-" & dat & "-" & tim & "生成.xlsx"
    ChDir cfile
    ActiveWorkbook.SaveAs Filename:=new_file, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    
    '关闭源文件
    Windows(ofName).Activate
    ActiveWorkbook.Close savechanges:=False

    '删除顶部多余行
    Sheets("成绩排名").Select
    top_rows = Application.Match("学生学号", Range("A1:A100"), 0) - 1
        If top_rows > 0 Then
            Rows("1:" & top_rows).Delete Shift:=xlUp
        ElseIf top_rows = 0 Then
            Range("A1").Select
        Else
            MsgBox "未找到学生学号字段，请确认文件是否正确"
        End If
        
    '删除不需要的列
    For i = 1 To 12
        Call del_col_key("班级排名", "A1:AF1")
        Call del_col_key("年级排名", "A1:AF1")
    Next
    Call del_col_key("语数外总分", "A1:AF1")
    
    '求成绩排名工作表最大行数列数
    成绩排名_colmax = Sheets("成绩排名").UsedRange.Columns.Count
    成绩排名_rowmax = Sheets("成绩排名").UsedRange.Rows.Count
    
    '在生成的工作簿中新建工作表，用于记录中间值和参数
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Name = "参数"
    Sheets("参数").Move Before:=Sheets(1) '拖拽到第一的位置
    
    '在参数工作表中填写列标题
    Range("A1") = "科目"
    Range("B1") = "是否考试"
    
    '创建科目数组，转置并填写到科目字段中
    arr_sub = Array("总分", "语文", "数学", "英语", "物理", "化学", "生物", "历史", "地理", "政治")
    Range("A2").Resize(10, 1) = Application.Transpose(arr_sub)
    
    '创建列标题数组，用于插列内容的判断
    arr_insert = Array("原始班次", "原始级次", "等级", "赋分", "赋分班次", "赋分级次")
    
    '计算考试相关参数，用于后面的判断
    arr_info = Array("所在列号", "应考人数", "缺考人数", "实考人数", "前1%多少人")
    Range("C1").Resize(1, 5) = arr_info
    
    '创建班级参数矩阵，用于后面的计算
    arr_til = Array("班级", "起始行号", "结束行号")
    Range("A13").Resize(1, 3) = arr_til
        
    '调用科目获取列号函数
    Call colpos
    
    '统计各科应考人数
    For i = 2 To 11
        If Sheets("参数").Range("B" & i) = 1 Then
            Sheets("参数").Range("D" & i) = 成绩排名_rowmax
        Else
            Sheets("参数").Range("D" & i) = 0
        End If
    Next
    
    '统计各科缺考人数
    For i = 2 To 11
        j = 0
        l = Sheets("参数").Range("C" & i)
        If l <> 0 Then
            Sheets("成绩排名").Select
            For k = 2 To 成绩排名_rowmax
                If IsNumeric(Sheets("成绩排名").Range(l & k)) = False Then
                    j = j + 1
                End If
            Next
            Sheets("参数").Range("E" & i) = j
        Else
            Sheets("参数").Range("E" & i) = 0
        End If
    Next
    
    '统计各科实考人数
    For i = 2 To 11
        Sheets("参数").Range("F" & i) = Sheets("参数").Range("D" & i) - Sheets("参数").Range("E" & i)
    Next
    
    '计算前1%是多少人
    For i = 2 To 11
        Sheets("参数").Range("G" & i) = WorksheetFunction.RoundUp(Sheets("参数").Range("F" & i) * 0.01, 0)
    Next
    
    '调用替换函数替换--
    Sheets("成绩排名").Select
    Call replace("A1:" & Split(Cells(1, 成绩排名_colmax).Address, "$")(1) & 成绩排名_rowmax, "--", "")
    
    '逐列文本转数值
    For i = 4 To 成绩排名_colmax
        Range(Cells(2, i), Cells(成绩排名_rowmax, i)).TextToColumns FieldInfo:=Array(1, 1)
    Next
    
    '替换班级列的班字
    Call replace("B2:B" & 成绩排名_rowmax, "班", "")
    
    '先按总分降序排列，再按班级升序排列
    Call crange_sort("A1", "D", 1)
    Call crange_sort("A1", "B", 0)
    
    '将班级内容复制到参数表中
    Sheets("成绩排名").Select
    Range("B2:B" & 成绩排名_rowmax).Copy
    Sheets("参数").Select
    Range("A14").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    '找到各班起始行号，填写结束行号
    参数_rowmax = Sheets("参数").UsedRange.Rows.Count
    For i = 14 To 参数_rowmax
        If Range("A" & i) <> Range("A" & i - 1) Then
            Range("B" & i) = i - 12
        End If
    Next
    Range("B14") = 2
    
    '去重，得到不重复的班号
    ActiveSheet.Range("$A$14:$C$" & 成绩排名_rowmax + 13).RemoveDuplicates Columns:=1, Header:=xlNo
    
    '填写结束行号
    参数_rowmax_distinct = Sheets("参数").UsedRange.Rows.Count
    For i = 14 To 参数_rowmax_distinct
        Range("C" & i) = Range("B" & i + 1) - 1
    Next
    Range("C" & 参数_rowmax_distinct) = 成绩排名_rowmax
    
    '调用插列函数
    For i = 2 To 11 '参数这个sheet，循环A列的各个科目
        Sheets("参数").Select
        If Range("B" & i) = 1 Then '如果B列是1，那么该科目存在，就插列
            j = Range("A" & i) '取出对应的科目名称
            Sheets("成绩排名").Select
            Call insert_subcol(j, 6) '调用插列函数，在对应的列后面插入6列，等后面再把多余的删掉！！！！
            l = Application.Match(j, Range(Cells(1, 1), Cells(1, Columns.Count)), 0) '定位该科目的表头位置
            For k = 0 To 5
                Cells(1, l).Offset(0, k + 1) = j & arr_insert(k) '插完列之后要选中对应的列头，逐个填充标题
            Next
        End If
    Next
    
    '删除不需要的列
    arr_delcol = Array("总分", "语文", "数学", "英语")
    For i = 0 To 3
        If Sheets("参数").Range("B" & i + 2) = 1 Then
            Call del_qcol(arr_delcol(i) & "等级", "A1:DD1")
            Call del_qcol(arr_delcol(i) & "赋分", "A1:DD1")
            Call del_qcol(arr_delcol(i) & "赋分班次", "A1:DD1")
            Call del_qcol(arr_delcol(i) & "赋分级次", "A1:DD1")
        End If
    Next
    
    '计算原始班次和级次
    For i = 2 To 11
        If Sheets("参数").Range("B" & i) <> 0 Then
            Sheets("成绩排名").Select
            Call group_rank(Sheets("参数").Range("A" & i) & "原始班次")
            Call grade_rank(Sheets("参数").Range("A" & i) & "原始级次")
        End If
    Next
    
    '计算等级赋分
    For i = 6 To 11
        If Sheets("参数").Range("B" & i) <> 0 Then
            Sheets("成绩排名").Select
            Call level_score(Sheets("参数").Range("A" & i))
        End If
    Next
    
    '计算赋分班次和级次
    For i = 6 To 11
        If Sheets("参数").Range("B" & i) <> 0 Then
            Sheets("成绩排名").Select
            Call group_rank(Sheets("参数").Range("A" & i) & "赋分班次")
            Call grade_rank(Sheets("参数").Range("A" & i) & "赋分级次")
        End If
    Next
    
    '设置边框线
    With Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
    End With
    
    '设置表头样式
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    Selection.Font.Bold = True
    
    '重新求成绩排名工作表最大行数列数
    Debug.Print "这里重新计算了列数"
    成绩排名_colmax = Sheets("成绩排名").UsedRange.Columns.Count
    
    '重新定位各科所在列，调用科目获取列号函数
    Call colpos
    
    '关键列上色
    For i = 2 To 11
        If Sheets("参数").Range("B" & i) <> 0 Then
            Sheets("成绩排名").Select
            Range(Sheets("参数").Range("C" & i) & 1).Select
            With Selection.Interior
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.799981688894314
            End With
        End If
    Next
    
    '创建各科sheet，复制总分列内容
    
    
    '完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True '重启刷新
    Application.DisplayAlerts = True '重启通知
    MsgBox "计算完成，用时" & Format(using_time, "0.0秒")
    
End Sub

Function file_open_name(til, ifilname) 'til是文件选择器标题，ifilname是默认打开路径
'先选择文件，获取路径，若未选择任何文件，终止程序，输出ofile是带路径的文件名，ofName是不带路径的文件名
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = til
        .InitialFileName = ifilname
        If .Show Then
            ofile = .SelectedItems(1)
            Else: Exit Function
        End If
    End With
'用斜杠分割文件路径，创建数组，选取数组最后一个元素做为不含路径的文件名
    splfile = Split(ofile, "\")
    ofName = splfile(UBound(splfile))
End Function

Function file_save_name(til, ifilname) 'til是文件选择器标题，ifilname是默认保存路径
'选择要保存的文件路径，若未选择任何文件夹，终止程序，输出cfile是带路径的文件夹名，cfName是不带路径的文件夹名
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = til
        .InitialFileName = ifilname
        If .Show = -1 Then
            cfile = .SelectedItems(1)
            Else: Exit Function
        End If
    End With
'用斜杠分割文件路径，创建数组，选取数组最后一个元素做为不含路径的文件名
    splfile = Split(cfile, "\")
    cfName = splfile(UBound(splfile))
End Function

'rbef表示替换什么，rlat表示替换后是什么，数字字符均可，字符用双引号'
Function replace(rang, rbef, rlat)
    Range(rang).Select
    Selection.replace What:=rbef, Replacement:=rlat, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function

Function insert_subcol(a, b) '参数a是需要找的列标题，参数b是需要插入的列数，在a后面插入b列
    For i = 1 To b
        Columns(Application.Match(a, Range(Cells(1, 1), Cells(1, Columns.Count)), 0) + 1).Insert _
        Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
End Function

Function del_col_key(a, b) '参数a表示要查找的中文，如"订单状态"，参数b表示一个查询范围，如："A1:AZ1"
Do
  colx = Application.Match(a, Range(b), 0)
  If IsNumeric(colx) = False Then Exit Do
    Columns(colx).Delete
Loop
End Function

Function crange_sort(ref, key, order) '三个参数：ref表示排序的区域，随便给个A1就行，key表示排序的关键字是哪个字段，order0升序，1降序
'如果是希望排多个字段，需要把权重最高的放在最后
    If order = 0 Then '升序
        Range(ref).CurrentRegion.Sort key1:=Range(key & "1"), order1:=xlAscending, Header:=xlYes
    ElseIf order = 1 Then '降序
        Range(ref).CurrentRegion.Sort key1:=Range(key & "1"), order1:=xlDescending, Header:=xlYes
    ElseIf order <> 0 Or order <> 1 Then
        Exit Function
    End If
End Function

Function del_qcol(a, b) '参数a是需要搜索的列名，如："付款时间"，参数b是查找范围，如："A1:AZ1"
    Columns(Application.Match(a, Range(b), 0)).Select
    Selection.Delete Shift:=xlToLeft    '删除列，默认右侧单元格左移
End Function

Function group_rank(x)
    '计算原始班次,传入一个列标题名称，定位到该列，循环计算各班班级名次，如：总分原始班次
    y = Application.Match(x, Range("A1:DD1"), 0)
    For i = 14 To 参数_rowmax_distinct
        a = Sheets("参数").Range("B" & i)
        b = Sheets("参数").Range("C" & i)
        For j = a To b
            Sheets("成绩排名").Select
            If Cells(j, y - 1) <> "" Then
                Cells(j, y) = WorksheetFunction.Rank(Cells(j, y - 1), Range(Cells(a, y - 1), Cells(b, y - 1)))
            End If
        Next
    Next
End Function

Function grade_rank(x)
    '计算原始级次,传入一个列标题名称，定位到该列，计算年级名次，如：总分原始级次
    y = Application.Match(x, Range("A1:DD1"), 0)
    a = Sheets("参数").Range("B14")
    b = Sheets("参数").Range("C" & 参数_rowmax_distinct)
    For j = a To b
        Sheets("成绩排名").Select
        If Cells(j, y - 2) <> "" Then
            Cells(j, y) = WorksheetFunction.Rank(Cells(j, y - 2), Range(Cells(a, y - 2), Cells(b, y - 2)))
        End If
    Next
End Function

Function level_score(x)
    y = Application.Match(x & "等级", Range("A1:DD1"), 0)
    Z = Application.Match(x, Sheets("参数").Range("A1:A11"), 0)
    a = Sheets("参数").Range("B14")
    b = Sheets("参数").Range("C" & 参数_rowmax_distinct)
    '计算相对排名
    For i = a To b
        rnk = Sheets("成绩排名").Cells(i, y - 1)
        num = Sheets("参数").Range("F" & Z)
        If num = 0 Then
            rnk_rate = 0
        Else
            rnk_rate = rnk / num
        End If
        '根据相对排名计算等级和赋分
        Select Case rnk_rate
            Case Is = 0
                Sheets("成绩排名").Cells(i, y) = ""
                Sheets("成绩排名").Cells(i, y + 1) = ""
            Case Is <= 0.01
                Sheets("成绩排名").Cells(i, y) = "A1"
                Sheets("成绩排名").Cells(i, y + 1) = 100
            Case Is <= 0.03
                Sheets("成绩排名").Cells(i, y) = "A2"
                Sheets("成绩排名").Cells(i, y + 1) = 97
            Case Is <= 0.06
                Sheets("成绩排名").Cells(i, y) = "A3"
                Sheets("成绩排名").Cells(i, y + 1) = 94
            Case Is <= 0.1
                Sheets("成绩排名").Cells(i, y) = "A4"
                Sheets("成绩排名").Cells(i, y + 1) = 91
            Case Is <= 0.15
                Sheets("成绩排名").Cells(i, y) = "A5"
                Sheets("成绩排名").Cells(i, y + 1) = 88
            Case Is <= 0.21
                Sheets("成绩排名").Cells(i, y) = "B1"
                Sheets("成绩排名").Cells(i, y + 1) = 85
            Case Is <= 0.28
                Sheets("成绩排名").Cells(i, y) = "B2"
                Sheets("成绩排名").Cells(i, y + 1) = 82
            Case Is <= 0.36
                Sheets("成绩排名").Cells(i, y) = "B3"
                Sheets("成绩排名").Cells(i, y + 1) = 79
            Case Is <= 0.43
                Sheets("成绩排名").Cells(i, y) = "B4"
                Sheets("成绩排名").Cells(i, y + 1) = 76
            Case Is <= 0.5
                Sheets("成绩排名").Cells(i, y) = "B5"
                Sheets("成绩排名").Cells(i, y + 1) = 73
            Case Is <= 0.57
                Sheets("成绩排名").Cells(i, y) = "C1"
                Sheets("成绩排名").Cells(i, y + 1) = 70
            Case Is <= 0.64
                Sheets("成绩排名").Cells(i, y) = "C2"
                Sheets("成绩排名").Cells(i, y + 1) = 67
            Case Is <= 0.71
                Sheets("成绩排名").Cells(i, y) = "C3"
                Sheets("成绩排名").Cells(i, y + 1) = 64
            Case Is <= 0.78
                Sheets("成绩排名").Cells(i, y) = "C4"
                Sheets("成绩排名").Cells(i, y + 1) = 61
            Case Is <= 0.84
                Sheets("成绩排名").Cells(i, y) = "C5"
                Sheets("成绩排名").Cells(i, y + 1) = 58
            Case Is <= 0.89
                Sheets("成绩排名").Cells(i, y) = "D1"
                Sheets("成绩排名").Cells(i, y + 1) = 55
            Case Is <= 0.93
                Sheets("成绩排名").Cells(i, y) = "D2"
                Sheets("成绩排名").Cells(i, y + 1) = 52
            Case Is <= 0.96
                Sheets("成绩排名").Cells(i, y) = "D3"
                Sheets("成绩排名").Cells(i, y + 1) = 49
            Case Is <= 0.98
                Sheets("成绩排名").Cells(i, y) = "D4"
                Sheets("成绩排名").Cells(i, y + 1) = 46
            Case Is <= 0.99
                Sheets("成绩排名").Cells(i, y) = "D5"
                Sheets("成绩排名").Cells(i, y + 1) = 43
            Case Is <= 1
                Sheets("成绩排名").Cells(i, y) = "E"
                Sheets("成绩排名").Cells(i, y + 1) = 40
        End Select
    Next
End Function

Function colpos()
'判断科目是否存在，1表示存在，0表示不存在，填写在参数表中
arr_col_pos = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0) '创建一个用于接收列号的空数组
For i = 0 To UBound(arr_sub) '逐个学科查找
    Sheets("成绩排名").Select
    col = Application.Match(arr_sub(i), Range(Cells(1, 1), Cells(1, 成绩排名_colmax)), 0)
        If IsNumeric(col) = True Then '如果科目名称存在，就在参数表对应科目后面写1
            Sheets("参数").Select
            Row = Application.Match(arr_sub(i), Range("A2:A11"), 0)
            Range("B" & Row + 1) = 1
            col_a = Split(Cells(1, col).Address, "$")(1) '将列号转换为列标
            arr_col_pos(i) = col_a '如果存在，逐个接收列号位置
        Else
            Sheets("参数").Select
            Row = Application.Match(arr_sub(i), Range("A2:A11"), 0) '如果科目名称不存在，就在参数表对应科目后面写0
            Range("B" & Row + 1) = 0
            arr_col_pos(i) = 0 '如果不存在，列号为0
        End If
Next
'填充列号
Range("C2").Resize(10, 1) = Application.Transpose(arr_col_pos)
End Function