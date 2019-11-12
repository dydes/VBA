Sub 生成全科等级赋分表()

'定义变量
    Dim title As String, subTitle As String, rowmax As Integer, file As String, fPath As String, splfile As Variant, fName As String, grade As String

'先选择文件，获取路径，若未选择任何文件，终止程序，选择的不是全年级-所有工作簿，终止程序
With Application.FileDialog(msoFileDialogFilePicker)
    .title = "请选择全年级文件夹下全科的工作簿"
    .InitialFileName = "c:\"
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
    .title = "请选择要保存的文件夹"
    .InitialFileName = "c:\"
    If .Show = -1 Then
        fPath = .SelectedItems(1)
    Else: Exit Sub
    End If
End With

'打开指定工作簿
Workbooks.Open (file)

'选择成绩排名工作表，复制并新建文件，保存新文件，关闭源文件
Windows(fName).Activate
    Sheets("成绩排名").Select
    Sheets("成绩排名").Copy
ChDir fPath
    ActiveWorkbook.SaveAs Filename:=fPath & "\全年级-全科.xlsx", FileFormat:= _
    xlOpenXMLWorkbook, CreateBackup:=False
Windows(fName).Activate
    ActiveWorkbook.Close savechanges:=False

'开始处理成绩排名工作表
grade = "全年级-全科.xlsx"
Windows(grade).Activate
Sheets("成绩排名").Select

'临时存储标题（标题合并处理的话，目前有点问题，有待研究）
    Range("A1:Z2").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge

'存储标题内容
    title = Range("A1").Value
    subTitle = Range("A2").Value

'删除标题行
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp

'文本转数值格式
    For i = 1 To 26
    Cells(2, i).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Cells(2, i), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Next

'单独插入总分赋分列
    Columns(5).Select
    Cells(3, 5).Activate
    Selection.Insert Shift:=xlToRight
    
'循环插入各单科赋分列
    For i = 16 To 40 Step 4
        Columns(i).Select
        Selection.Insert Shift:=xlToRight
        Selection.Insert Shift:=xlToRight
    Next

'填入数组
    Range("A1:M1") = Array("学生学号", "班级", "姓名", "原始总分", "赋分总分", "原始分班级排名", "原始分年级排名", "语文", "年级排名", "数学", "年级排名", "英语", "年级排名")
    Range("N1:Y1") = Array("物理原始分", "年级排名", "物理等级", "物理赋分", "化学原始分", "年级排名", "化学等级", "化学赋分", "生物原始分", "年级排名", "生物等级", "生物赋分")
    Range("Z1:AK1") = Array("历史原始分", "年级排名", "历史等级", "历史赋分", "地理原始分", "年级排名", "地理等级", "地理赋分", "政治原始分", "年级排名", "政治等级", "政治赋分")
    Range("AL1:AM1") = Array("语数外总分", "年级排名")
    
'替横线向--
    Columns("A:A").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Replace What:="--", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'工作表改名
    Sheets("成绩排名").Select
    Sheets("成绩排名").Name = "全科"

'计算等级及赋分
    rowmax = ActiveSheet.UsedRange.Rows.Count
    Sheets("全科").Range("P2:P" & rowmax).Formula = "=IF(O2<>"""",LOOKUP(O2/(COUNTA(O:O)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{""A1"",""A2"",""A3"",""A4"",""A5"",""B1"",""B2"",""B3"",""B4"",""B5"",""C1"",""C2"",""C3"",""C4"",""C5"",""D1"",""D2"",""D3"",""D4"",""D5"",""E""}),0)"
    Sheets("全科").Range("Q2:Q" & rowmax).Formula = "=IF(O2<>"""",LOOKUP(O2/(COUNTA(O:O)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{100,97,94,91,88,85,82,79,76,73,70,67,64,61,58,55,52,49,46,43,40}),0)"
    Sheets("全科").Range("T2:T" & rowmax).Formula = "=IF(S2<>"""",LOOKUP(S2/(COUNTA(S:S)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{""A1"",""A2"",""A3"",""A4"",""A5"",""B1"",""B2"",""B3"",""B4"",""B5"",""C1"",""C2"",""C3"",""C4"",""C5"",""D1"",""D2"",""D3"",""D4"",""D5"",""E""}),0)"
    Sheets("全科").Range("U2:U" & rowmax).Formula = "=IF(S2<>"""",LOOKUP(S2/(COUNTA(S:S)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{100,97,94,91,88,85,82,79,76,73,70,67,64,61,58,55,52,49,46,43,40}),0)"
    Sheets("全科").Range("X2:X" & rowmax).Formula = "=IF(W2<>"""",LOOKUP(W2/(COUNTA(W:W)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{""A1"",""A2"",""A3"",""A4"",""A5"",""B1"",""B2"",""B3"",""B4"",""B5"",""C1"",""C2"",""C3"",""C4"",""C5"",""D1"",""D2"",""D3"",""D4"",""D5"",""E""}),0)"
    Sheets("全科").Range("Y2:Y" & rowmax).Formula = "=IF(W2<>"""",LOOKUP(W2/(COUNTA(W:W)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{100,97,94,91,88,85,82,79,76,73,70,67,64,61,58,55,52,49,46,43,40}),0)"
    Sheets("全科").Range("AB2:AB" & rowmax).Formula = "=IF(AA2<>"""",LOOKUP(AA2/(COUNTA(AA:AA)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{""A1"",""A2"",""A3"",""A4"",""A5"",""B1"",""B2"",""B3"",""B4"",""B5"",""C1"",""C2"",""C3"",""C4"",""C5"",""D1"",""D2"",""D3"",""D4"",""D5"",""E""}),0)"
    Sheets("全科").Range("AC2:AC" & rowmax).Formula = "=IF(AA2<>"""",LOOKUP(AA2/(COUNTA(AA:AA)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{100,97,94,91,88,85,82,79,76,73,70,67,64,61,58,55,52,49,46,43,40}),0)"
    Sheets("全科").Range("AF2:AF" & rowmax).Formula = "=IF(AE2<>"""",LOOKUP(AE2/(COUNTA(AE:AE)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{""A1"",""A2"",""A3"",""A4"",""A5"",""B1"",""B2"",""B3"",""B4"",""B5"",""C1"",""C2"",""C3"",""C4"",""C5"",""D1"",""D2"",""D3"",""D4"",""D5"",""E""}),0)"
    Sheets("全科").Range("AG2:AG" & rowmax).Formula = "=IF(AE2<>"""",LOOKUP(AE2/(COUNTA(AE:AE)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{100,97,94,91,88,85,82,79,76,73,70,67,64,61,58,55,52,49,46,43,40}),0)"
    Sheets("全科").Range("AJ2:AJ" & rowmax).Formula = "=IF(AI2<>"""",LOOKUP(AI2/(COUNTA(AI:AI)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{""A1"",""A2"",""A3"",""A4"",""A5"",""B1"",""B2"",""B3"",""B4"",""B5"",""C1"",""C2"",""C3"",""C4"",""C5"",""D1"",""D2"",""D3"",""D4"",""D5"",""E""}),0)"
    Sheets("全科").Range("AK2:AK" & rowmax).Formula = "=IF(AI2<>"""",LOOKUP(AI2/(COUNTA(AI:AI)-1),{0,1,3,6,10,15,22,30,39,47,55,62,68,74,80,85,89,93,96,98,99}%,{100,97,94,91,88,85,82,79,76,73,70,67,64,61,58,55,52,49,46,43,40}),0)"

'所有数都贴为死数
    Columns("A:A").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'求和
    Sheets("全科").Range("E2:E" & rowmax).Formula = "=H2+J2+L2+Q2+U2+Y2+AC2+AG2+AK2"

'赋分总分列都贴为死数
    Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'增加赋分年级排名列
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "赋分年级排名"
    Range("H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
'计算赋分年级排名
    Sheets("全科").Range("H2:H" & rowmax).Formula = "=RANK.EQ(E2,E:E,0)"
    
'赋分列排名粘贴为死数
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'按赋分年级排名排序
    Columns("H:H").Select
    ActiveWorkbook.Worksheets("全科").Sort.SortFields.Add2 Key:=Range("H1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("全科").Sort
        .SetRange Range("A2:AN339")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'把赋分列的0替换为空
    Range("Q:R,U:V,Y:Z,AC:AD,AG:AH,AK:AL").Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'创建各单科Sheet
    Sheets.Add After:=ActiveSheet
    Sheets(2).Name = "语文"
    Sheets.Add After:=ActiveSheet
    Sheets(3).Name = "数学"
    Sheets.Add After:=ActiveSheet
    Sheets(4).Name = "英语"
    Sheets.Add After:=ActiveSheet
    Sheets(5).Name = "物理"
    Sheets.Add After:=ActiveSheet
    Sheets(6).Name = "化学"
    Sheets.Add After:=ActiveSheet
    Sheets(7).Name = "生物"
    Sheets.Add After:=ActiveSheet
    Sheets(8).Name = "历史"
    Sheets.Add After:=ActiveSheet
    Sheets(9).Name = "地理"
    Sheets.Add After:=ActiveSheet
    Sheets(10).Name = "政治"
    
'将全科表对应内容复制到单科表中
    Sheets("全科").Select
    Columns("A:J").Select
    Selection.Copy
    Sheets("语文").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("全科").Select
    Range("A:H,K:L").Select
    Selection.Copy
    Sheets("数学").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,M:N").Select
    Selection.Copy
    Sheets("英语").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,O:R").Select
    Selection.Copy
    Sheets("物理").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,S:V").Select
    Selection.Copy
    Sheets("化学").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,W:Z").Select
    Selection.Copy
    Sheets("生物").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,AA:AD").Select
    Selection.Copy
    Sheets("历史").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,AE:AH").Select
    Selection.Copy
    Sheets("地理").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("全科").Select
    Range("A:H,AI:AL").Select
    Selection.Copy
    Sheets("政治").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'设置打印格式
    '循环选中工作表
    For i = 1 To 10 Step 1
    Sheets(i).Select
        '自适应各列
        Columns("A:A").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Columns("A:L").EntireColumn.AutoFit
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        '隔行标色
        For j = 3 To rowmax Step 2
            Rows(j).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -4.99893185216834E-02
                .PatternTintAndShade = 0
            End With
        Next
    Next

    '设置打印区域
    For i = 1 To 10 Step 1
    Sheets(i).Select
    ActiveWindow.View = xlPageBreakPreview
    If ActiveSheet.VPageBreaks.Count > 0 Then
            ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        End If '判断了一下，如果是一页就不调整了
    ActiveWindow.View = xlNormalView
    Next
    
    '全科调整列宽
    Sheets("全科").Select
    Columns("D:D").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ColumnWidth = 5
    Range("D1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'赋分排名列标题标色
    For i = 1 To 4 Step 1
    Sheets(i).Select
    Range("H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Next
    For i = 5 To 10 Step 1
    Sheets(i).Select
    Range("H1,K1,L1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Next

'冻结窗口
    For i = 1 To 10 Step 1
    Sheets(i).Select
    Range("D2").Select
    ActiveWindow.FreezePanes = True
    Next
    
'最后按班拆分成多个工作簿

'完成提示
ActiveWorkbook.Save
    MsgBox "生成完成"
    
End Sub