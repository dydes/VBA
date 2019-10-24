Sub 河南加权分()

'定义文件处理相关变量
    Dim splfile As Variant, fName As String, fPath As String, file As String
'定义运行时间相关变量
    Dim dat As String, tim As String, tim1 As String, tim2 As String
'定义最大行数相关变量
    Dim rowmax As Integer, colmax As Integer
'定义新增列相关变量
    Dim col_a As Integer, col_e As Integer
    
    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知

'先选择文件，获取路径，若未选择任何文件，终止程序
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "请选择年级全科文件"
        .InitialFileName = "D:\会通\VBA\加权分\"
            If .Show Then
                file = .SelectedItems(1)
            Else: Exit Sub
            End If
    End With

'用斜杠分割文件路径，创建数组，选取数组最后一个元素做为不含路径的文件名
    splfile = Split(file, "\")
    fName = splfile(UBound(splfile))

'选择要保存的文件路径，若未选择任何文件夹，终止程序
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择要保存的文件夹"
        .InitialFileName = "D:\会通\VBA\加权分\"
        If .Show = -1 Then
            fPath = .SelectedItems(1)
        Else: Exit Sub
        End If
    End With

'打开指定工作簿
    Workbooks.Open (file)

'获取当前时间
    dat = Format(Date, "yyyy年mm月dd日") '当前年月日
    tim = Format(Time, "hh时mm分ss秒") '当前时间
    tim1 = Timer

'选择“成绩排名”工作表，复制并新建文件，保存新文件，关闭源文件
    Windows(fName).Activate
        Sheets("成绩排名").Select
        Sheets("成绩排名").Copy
    ChDir fPath
        ActiveWorkbook.SaveAs Filename:=fPath & "\年级全科1.25加权后-" & dat & "-" & tim & "生成.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows(fName).Activate
        ActiveWorkbook.Close savechanges:=False
        
'删除顶部1或2行
    Sheets("成绩排名").Select
    If Range("A3").Value = "学生学号" Then
        Rows("1:2").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
    End If
    
'判断是否有英语
    ss = 0
    For tt = 6 To 16
        If Cells(1, tt).Value = "英语" Then
            ss = ss + 1
        End If
    Next
    If ss > 0 Then
        GoTo 1
    Else
        MsgBox "未找到英语学科"
        Exit Sub
    End If

'文本转数值格式
1:  colmax = ActiveSheet.UsedRange.Columns.Count
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="--", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    For i = 1 To colmax
    Cells(2, i).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Cells(2, i), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Next

'加列
    '查找并增加总分、总分班次、总分级次3列
    col_a = Rows("1:1").Find(What:="总分").Column
    col_a_Name = Split(Cells(1, col_a).Address, "$")(1) 'col_array(col_a - 1) '总分列名
    For i = 1 To 5 Step 2
        Columns(col_a + i).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    '查找并增加英语加权、英语加权年级排名2列
    col_e = Rows("1:1").Find(What:="英语").Column
    col_e_Name = Split(Cells(1, col_e).Address, "$")(1) 'col_array(col_e - 1) '英语列号
    For i = 1 To 3 Step 2
        Columns(col_e + i).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next

'列标转换矩阵
    '最大行列
    colmax = ActiveSheet.UsedRange.Columns.Count
    rowmax = ActiveSheet.UsedRange.Rows.Count
    '列号及列名
    col_a125 = col_a + 1 '总分加权列号
    col_a125_Name = Split(Cells(1, col_a125).Address, "$")(1)

    col_a125_cr = col_a + 3 '总分加权班级排名列号
    col_a125_cr_Name = Split(Cells(1, col_a125_cr).Address, "$")(1)

    col_a125_gr = col_a + 5 '总分加权年级排名列号
    col_a125_gr_Name = Split(Cells(1, col_a125_gr).Address, "$")(1)

    col_c = col_a - 1 '班级名称列号
    col_c_Name = Split(Cells(1, col_c).Address, "$")(1)
    
    col_a_cr = col_a + 2 '总分班级排名列号
    col_a_cr_Name = Split(Cells(1, col_a_cr).Address, "$")(1)
    
    col_a_gr = col_a + 4 '总分年级排名列号
    col_a_gr_Name = Split(Cells(1, col_a_gr).Address, "$")(1)

    col_e125 = col_e + 1 '英语加权列号
    col_e125_Name = Split(Cells(1, col_e125).Address, "$")(1)

    col_e125_gr = col_e + 3 '英语加权年级排名列号
    col_e125_gr_Name = Split(Cells(1, col_e125_gr).Address, "$")(1)

    col_e_gr = col_e + 2 '英语年级排名列号
    col_e_gr_Name = Split(Cells(1, col_e_gr).Address, "$")(1)

    colmaxN = Split(Cells(1, colmax).Address, "$")(1) '最大列名
    
    colmax1 = colmax + 1
    colmaxN1 = Split(Cells(1, colmax1).Address, "$")(1)  '最大+1列名
    
    colmax2 = colmax + 2
    colmaxN2 = Split(Cells(1, colmax2).Address, "$")(1) '最大+2列名
    
    colmax3 = colmax + 3
    colmaxN3 = Split(Cells(1, colmax3).Address, "$")(1)  '最大+3列名
    
    colmax4 = colmax + 4
    colmaxN4 = Split(Cells(1, colmax4).Address, "$")(1)  '最大+4列名

'命名表头
    Cells(1, col_a125_Name) = "总分加权"
    Cells(1, col_a125_cr) = "加权班级排名"
    Cells(1, col_a125_gr) = "加权年级排名"
    Cells(1, col_e125) = "英语加权"
    Cells(1, col_e125_gr) = "加权年级排名"

'计算英语学科加权分及排名
    '计算英语加权分
    For i = 2 To rowmax
        Cells(i, col_e125).Value = Cells(i, col_e).Value * 1.25
        Cells(i, col_e125).NumberFormatLocal = "0.0"
    Next
    '按加权分降序计算加权年级排名
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Add2 Key:=Range(col_e125_Name & "2:" & col_e125_Name & rowmax) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("成绩排名").Sort
        .SetRange Range("A1:" & colmaxN & rowmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '计算年级排名
    For i = 2 To rowmax
        Cells(i, col_e125_gr) = Application.Rank(Cells(i, col_e125), Range(col_e125_Name & "2:" & col_e125_Name & rowmax))
    Next

'创建科目数组
    subject_array = Array("语文", "数学", "英语加权", "物理", "化学", "生物", "历史", "地理", "政治")
    subject_col_array = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    Dim col_istrue As Range
    For i = 0 To 8
        Set col_istrue = Rows("1:1").Find(What:=subject_array(i))
        If Not col_istrue Is Nothing Then
            subject_col_array(i) = Split(col_istrue.Address, "$")(1)
        End If
    Next
    
'清洗科目数组，过滤掉不存在的列名，创建纯列名数组
    col_isNotNull = Join(subject_col_array, ",")
    col_isNotNull = Replace(col_isNotNull, "0,", "")
    col_isNotNull = Replace(col_isNotNull, ",0", "")
    col_isNotNull = Replace(col_isNotNull, "0", "")
    col_isNotNull = Replace(col_isNotNull, ",,", ",")
    subject_col_array = Split(col_isNotNull, ",")
    
'计算加权总分
    For h = 2 To rowmax
        a = 0
        For Each i In subject_col_array
            a = a + Cells(h, i).Value
        Next
        Cells(h, col_a125).Value = a
        Cells(h, col_a125).NumberFormatLocal = "0.0"
    Next

'计算加权总分及排名
    '按总分加权降序排
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Add2 Key:=Range(col_a125_Name & "2:" & col_a125_Name & rowmax) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("成绩排名").Sort
        .SetRange Range("A1:" & colmaxN & rowmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '计算总分加权年级排名
    For i = 2 To rowmax
        Cells(i, col_a125_gr) = Application.Rank(Cells(i, col_a125), Range(col_a125_Name & "2:" & col_a125_Name & rowmax))
    Next
    '按班升序再按总分降序排列
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Add2 Key:=Range("C2:C" & rowmax) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Add2 Key:=Range("E2:E" & rowmax) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("成绩排名").Sort
        .SetRange Range("A1:AA" & rowmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'成绩排名工作表添加班级辅助列并去重
    Range(colmaxN1 & "2:" & colmaxN1 & rowmax).Value = Range("C2:C" & rowmax).Value
    Range(colmaxN1 & "2:" & colmaxN1 & rowmax).RemoveDuplicates 1
    cn = Application.CountA(Range(colmaxN1 & ":" & colmaxN1)) '去重后的班级数量
    
'统计班级人数、计算起始截止行数
    x = 2
    For i = 2 To cn + 1
        Range(colmaxN2 & i) = Application.CountIf(Range("C1:C" & rowmax), Range(colmaxN1 & i))
        Range(colmaxN3 & i) = x
        Range(colmaxN4 & i) = x + Range(colmaxN2 & i) - 1
        x = Range(colmaxN4 & i).Value + 1
    Next
    
'计算总分加权班级排名
    For j = 2 To cn + 1
        For i = Range(colmaxN3 & j).Value To Range(colmaxN4 & j).Value
            Cells(i, 7) = Application.Rank(Cells(i, 5), Range("E" & Range(colmaxN3 & j).Value & ":E" & Range(colmaxN4 & j).Value)) '按班级排名
        Next
    Next

'逐个取辅助列的值，命名新建sheet，复制内容
    For k = 2 To cn + 1
        clName = Sheets("成绩排名").Range(colmaxN1 & k).Value
        sheetNum = Sheets.Count
        Sheets.Add after:=Sheets(sheetNum)
        Sheets(sheetNum + 1).Name = clName
    Next
    
'逐个填充内容
    For t = 2 To cn + 1
        Sheets("成绩排名").Select
        Range("A" & Range(colmaxN3 & t).Value & ":" & colmaxN & Range(colmaxN4 & t).Value).Select
        Selection.Copy
        Sheets(Range(colmaxN1 & t).Value).Select
        Range("A2:" & colmaxN & Range(colmaxN2 & t).Value + 1).Select
        ActiveSheet.Paste
        Rows("1:1").Select
        Application.CutCopyMode = False
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets("成绩排名").Select
        Range("A1:" & colmaxN & 1).Select
        Selection.Copy
        Sheets(Range(colmaxN1 & t).Value).Select
        Range("A1:" & colmaxN & 1).Select
        ActiveSheet.Paste
    Next
 
'删除辅助列
    Sheets("成绩排名").Select
    Columns(colmaxN1 & ":" & colmaxN4).Select
    Selection.Delete Shift:=xlToLeft
    
'按总分加权降序排列
    Columns("A:" & colmaxN).Select
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("成绩排名").Sort.SortFields.Add2 Key:=Range("E2:E" & rowmax) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("成绩排名").Sort
        .SetRange Range("A1:" & colmaxN & rowmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'逐个调整行高列宽，设置标题字体样式
    For s = 1 To sheetNum
        Sheets(s).Select
        Columns("A:" & col_c_Name).ColumnWidth = 13
        Rows("1:" & rowmax).EntireRow.AutoFit
        Range("A1:" & colmaxN & "1").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = True
    Next

'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
ActiveWindow.WindowState = xlMaximized
ActiveWorkbook.Save

Application.ScreenUpdating = True '重启刷新
MsgBox "计算完成，用时" & Format(using_time, "0.0秒")

End Sub

