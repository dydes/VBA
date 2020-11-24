Sub 横竖版()

'校验考试名称
    If Len(Sheets("说明").Range("B2")) = 0 Then
        MsgBox "请在第1个sheet的B2单元格填写考试名称"
        GoTo 1
    End If

'校验科目
    If Len(Sheets("说明").Range("B3")) = 0 Then
        MsgBox "请在第1个sheet的B3单元格填写考试科目"
        GoTo 1
    End If
    
'判断总分表表头
    If Sheets("总表").Range("A1") <> "学号" Then
        MsgBox "请勿更改总表表头：学号、班级、姓名、总分、客观题、主观题"
        GoTo 1
    End If

'获取当前时间
    dat = Format(Date, "yyyy年mm月dd日") '当前年月日
    tim = Format(Time, "hh时mm分ss秒") '当前时间
    tim1 = Timer
    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知

'选择要保存的文件路径，若未选择任何文件夹，终止程序
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择要保存的文件夹"
        .InitialFileName = "D:\会通\VBA\横竖版\"
        If .Show = -1 Then
            fPath = .SelectedItems(1)
        Else: Exit Sub
        End If
    End With
    fName = "横竖版成绩-" & dat & tim & ".xlsx" '新建工作簿的名称
    gName = ActiveWorkbook.Name
    
'新建工作簿
    Workbooks.Add
    ChDir fPath
        ActiveWorkbook.SaveAs Filename:=fPath & "\" & fName, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        
'工作表命名
    Sheets("Sheet1").Name = "说明"
    
'将模板中说明和总表的内容转移到新建的文件中
    Windows(gName).Activate
    Sheets("说明").Activate
    Sheets("说明").Range("A1:B3").Select
    Selection.Copy
    Windows(fName).Activate
    Sheets("说明").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Windows(gName).Activate
    Sheets("总表").Select
    Sheets("总表").Copy After:=Workbooks(fName).Sheets(1)

'获取总表最大行数
    Windows(fName).Activate
    Sheets("总表").Select
    rowmax = ActiveSheet.UsedRange.Rows.Count
    
'班级列纯数值排序
    Range("B2:B" & rowmax).Select
    Selection.Replace What:="班", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

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

'处理缺考
    For i = 2 To rowmax
        If Range("D" & i) = "-" Then
            Range("D" & i) = 0
            Range("E" & i) = "缺考"
            Range("F" & i) = "缺考"
        End If
    Next

'按班级升序，按总分降序
    Sheets("总表").Select
    Columns("A:F").Select
    ActiveWorkbook.Worksheets("总表").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("总表").Sort.SortFields.Add Key:=Range("B2:B" & rowmax), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("总表").Sort.SortFields.Add Key:=Range("D2:D" & rowmax), _
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
    rowmax1 = ActiveSheet.UsedRange.Rows.Count '总表粘贴过来的最大行数
    ActiveSheet.Range("$A$4:$C$" & rowmax1).RemoveDuplicates Columns:=1, Header:=xlYes
    Sheets("说明").Range("A4") = "班号"
    Sheets("说明").Range("B4") = "总表起始行号"
    Sheets("说明").Range("C4") = "总表结束行号"
    Sheets("说明").Range("D4") = "总表行数"
    Sheets("说明").Range("E4") = "双左起始行号"
    Sheets("说明").Range("F4") = "双左结束行号"
    Sheets("说明").Range("G4") = "双右起始行号"
    Sheets("说明").Range("H4") = "双右结束行号"
    Sheets("说明").Range("I4") = "三左起始行号"
    Sheets("说明").Range("J4") = "三左结束行号"
    Sheets("说明").Range("K4") = "三中起始行号"
    Sheets("说明").Range("L4") = "三中结束行号"
    Sheets("说明").Range("M4") = "三右起始行号"
    Sheets("说明").Range("N4") = "三右结束行号"
    
'调整列宽
    Call cwidth("B:N", 5)

'生成各班结束行数
    Windows(fName).Activate
    Sheets("说明").Select
    rowmax2 = ActiveSheet.UsedRange.Rows.Count '班级矩阵的最大行数
    For i = 5 To rowmax2 - 1
        Sheets("说明").Range("C" & i) = Sheets("说明").Range("B" & i + 1) - 1
    Next
    Sheets("说明").Range("C" & rowmax2) = rowmax

'计算总表行号
    For i = 5 To rowmax2
        Sheets("说明").Range("D" & i) = Sheets("说明").Range("C" & i) - Sheets("说明").Range("B" & i) + 1
    Next
    
'计算双栏行号
    For i = 5 To rowmax2
        If Sheets("说明").Range("D" & i) Mod 2 = 1 Then
            j = (Sheets("说明").Range("D" & i) + 1) / 2
        Else
            j = Sheets("说明").Range("D" & i) / 2
        End If
        Sheets("说明").Range("E" & i) = Sheets("说明").Range("B" & i) '左侧起始就是班级起始行号
        Sheets("说明").Range("H" & i) = Sheets("说明").Range("C" & i) '右侧结束就是班级结束行号
        Sheets("说明").Range("F" & i) = Sheets("说明").Range("B" & i) + j - 1 '左侧结束行号是起始加一半行数减1
        Sheets("说明").Range("G" & i) = Sheets("说明").Range("F" & i) + 1 '右侧起始行号是左侧结束行号加1
    Next
    
'计算三栏折行
    For i = 5 To rowmax2
        Sheets("说明").Range("I" & i) = Sheets("说明").Range("B" & i) '左侧起始就是班级起始行号
        Sheets("说明").Range("N" & i) = Sheets("说明").Range("C" & i) '右侧结束就是班级结束行号
        
        If Sheets("说明").Range("D" & i) Mod 3 = 2 Then '如果被3除2
            j = (Sheets("说明").Range("D" & i) + 1) / 3
            Sheets("说明").Range("J" & i) = Sheets("说明").Range("B" & i) + j - 1 '左侧结束行号是起始加三分之一行数减1
            Sheets("说明").Range("K" & i) = Sheets("说明").Range("J" & i) + 1 '中间起始行号是左侧结束行号加1
            k = Sheets("说明").Range("K" & i)
            Sheets("说明").Range("L" & i) = k + j - 1 '中间结束是中间起始加三分之一行数减1
            Sheets("说明").Range("M" & i) = Sheets("说明").Range("L" & i) + 1 '右侧开始是中间结束行号加1
        ElseIf Sheets("说明").Range("D" & i) Mod 3 = 1 Then '如果被3除余1
            j = (Sheets("说明").Range("D" & i) - 1) / 3
            Sheets("说明").Range("J" & i) = Sheets("说明").Range("B" & i) + j '左侧结束行号是起始加三分之一行数
            Sheets("说明").Range("K" & i) = Sheets("说明").Range("J" & i) + 1 '中间起始行号是左侧结束行号加1
            k = Sheets("说明").Range("K" & i)
            Sheets("说明").Range("L" & i) = k + j '中间结束是中间起始加三分之一行数
            Sheets("说明").Range("M" & i) = Sheets("说明").Range("L" & i) + 1 '右侧开始是中间结束行号加1
        Else '如果被3整除
            j = Sheets("说明").Range("D" & i) / 3
            Sheets("说明").Range("J" & i) = Sheets("说明").Range("B" & i) + j - 1 '左侧结束行号是起始加三分之一行数减1
            Sheets("说明").Range("K" & i) = Sheets("说明").Range("J" & i) + 1 '中间起始行号是左侧结束行号加1
            k = Sheets("说明").Range("K" & i)
            Sheets("说明").Range("L" & i) = k + j - 1 '中间结束是中间起始加三分之一行数减1
            Sheets("说明").Range("M" & i) = Sheets("说明").Range("L" & i) + 1 '右侧开始是中间结束行号加1
        End If
    Next

'删除总表GH列，调整列顺序
    Windows(fName).Activate
    Sheets("总表").Select
    Range("G:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select '班级列调整到A
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select '总分列调整到F
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    
'逐个创建各班总分sheet
    For i = 5 To rowmax2
        '生成新sheet
        clName = Sheets("说明").Range("A" & i).Value & "班总"
        sheetNum = Sheets.Count
        Sheets.Add After:=Sheets(sheetNum)
        Sheets(sheetNum + 1).Name = clName
        '复制表头
        Sheets("总表").Range("B1:F1").Copy
        Sheets(clName).Select
        Range("A1").Select
        ActiveSheet.Paste
        '复制内容
        bg = Sheets("说明").Range("B" & i).Value
        en = Sheets("说明").Range("C" & i).Value
        Sheets("总表").Range("B" & bg & ":F" & en & "").Copy
        Sheets(clName).Select
        Range("A2").Select
        ActiveSheet.Paste
        '插入表头
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1") = Sheets("说明").Range("A3")
        Subject = "考试科目：" & Sheets("说明").Range("B3")
        Group = "班级：" & Sheets("说明").Range("A" & i) & "班"
        Range("A2") = Subject & "   " & Group
        
        '调用cmerge函数合并标题
        cmerge ("A1:E1")
        cmerge ("A2:E2")
        
        '调用列宽自适应函数
        Call cfit("A:E")
        
        '调用表格画线函数
        Call tline("A1", "E")

        '调用标题上色函数
        Call tcolor("A1:E3")

        '调用显示打印边界函数
        Call psetting
    Next
    
'逐个创建各班双栏sheet
    For i = 5 To rowmax2
        '生成新sheet
        clName = Sheets("说明").Range("A" & i).Value & "班竖"
        sheetNum = Sheets.Count
        Sheets.Add After:=Sheets(sheetNum)
        Sheets(sheetNum + 1).Name = clName
        '复制表头
        Sheets("总表").Range("B1:F1").Copy
        Sheets(clName).Select
        Range("A1").Select
        ActiveSheet.Paste
        Range("G1").Select
        ActiveSheet.Paste
        '复制内容
        bgl = Sheets("说明").Range("E" & i).Value
        enl = Sheets("说明").Range("F" & i).Value
        bgr = Sheets("说明").Range("G" & i).Value
        enr = Sheets("说明").Range("H" & i).Value
        Sheets("总表").Range("B" & bgl & ":F" & enl & "").Copy
        Sheets(clName).Select
        Range("A2").Select
        ActiveSheet.Paste
        Sheets("总表").Range("B" & bgr & ":F" & enr & "").Copy
        Sheets(clName).Select
        Range("G2").Select
        ActiveSheet.Paste
        Range("F1") = 1
        '插入表头
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1") = Sheets("说明").Range("B2")
        Subject = "考试科目：" & Sheets("说明").Range("B3")
        Group = "班级：" & Sheets("说明").Range("A" & i) & "班"
        Range("A2") = Subject & "        " & Group
        
        '调用cmerge函数合并标题
        cmerge ("A1:K1")
        cmerge ("A2:K2")
        
        '调用列宽自适应函数
        Call cfit("A:K")
        
        '调用表格画线函数
        Call tline("A1", "K")
        
        '调用标题上色函数
        Call tcolor("A1:K3")
        
        '调用标题隐色函数
        Call tlhide("F3")
        
        '调用显示打印边界函数
        Call psetting
    Next

'逐个创建各班三栏sheet
    Windows(fName).Activate
    For i = 5 To rowmax2
        '生成新sheet
        clName = Sheets("说明").Range("A" & i).Value & "班横"
        sheetNum = Sheets.Count
        Sheets.Add After:=Sheets(sheetNum)
        Sheets(sheetNum + 1).Name = clName
        '复制表头
        Sheets("总表").Range("B1:F1").Copy
        Sheets(clName).Select
        Range("A1").Select
        ActiveSheet.Paste
        Range("G1").Select
        ActiveSheet.Paste
        Range("M1").Select
        ActiveSheet.Paste
        '复制内容
        bgl = Sheets("说明").Range("I" & i).Value
        enl = Sheets("说明").Range("J" & i).Value
        bgm = Sheets("说明").Range("K" & i).Value
        enm = Sheets("说明").Range("L" & i).Value
        bgr = Sheets("说明").Range("M" & i).Value
        enr = Sheets("说明").Range("N" & i).Value
        Sheets("总表").Range("B" & bgl & ":F" & enl & "").Copy
        Sheets(clName).Select
        Range("A2").Select
        ActiveSheet.Paste
        Sheets("总表").Range("B" & bgm & ":F" & enm & "").Copy
        Sheets(clName).Select
        Range("G2").Select
        ActiveSheet.Paste
        Sheets("总表").Range("B" & bgr & ":F" & enr & "").Copy
        Sheets(clName).Select
        Range("M2").Select
        ActiveSheet.Paste
        Range("F1") = 1
        Range("L1") = 1
        '插入表头
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1") = Sheets("说明").Range("B2")
        Subject = "考试科目：" & Sheets("说明").Range("B3")
        Group = "班级：" & Sheets("说明").Range("A" & i) & "班"
        Range("A2") = Subject & "        " & Group
        
        '调用cmerge函数合并标题
        cmerge ("A1:Q1")
        cmerge ("A2:Q2")
        
        '调用列宽自适应函数
        Call cfit("A:Q")

        '调用表格画线函数
        Call tline("A1", "Q")
        
        '调用标题上色函数
        Call tcolor("A1:Q3")

        '调用标题隐色函数
        Call tlhide("F3")
        Call tlhide("L3")
        
        '调整为横向打印
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 1200
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
        
        '调用显示打印边界函数
        Call psetting
    Next

'保存
    Windows(fName).Activate
    ThisWorkbook.Save

'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True '重启刷新
    MsgBox "计算完成，用时" & Format(using_time, "0.0秒")
    
1
End Sub
Function cmerge(a) '单元格合并函数，根据传入的区域合并单元格

'传入的a是一个区域，需要用双引号，比如："A1:Q1"
    Range(a).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Function

Function cwidth(a, b) '列宽调整函数
    Columns(a).Select '传入的a需要是带引号的连续列，如："B:N"
    Selection.ColumnWidth = b '传入的b需要是一个数字，表示列宽
End Function

Function cfit(a) '列宽自适应函数
    Columns(a).Select '传入的a需要是带引号的连续列，如："A:Q"
    Columns(a).EntireColumn.AutoFit
End Function

Function tline(a, b) '可见区域画表格线函数
    maxrow = ActiveSheet.UsedRange.Rows.Count
    Range(a & ":" & b & maxrow).Select 'a表示左上角起始位置需要带引号，如："A1"，b表示结束位置的列号需要带引号，如："Q"，函数会自动计算右下角结束位置
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Function

Function tcolor(a) '标题上色函数
    Range(a).Select '传入的a需要是带引号的一个区域，如："A1:Q3"，目前颜色固定为蓝色
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Function

Function tlhide(a) '将单元格文字颜色改为与底纹颜色一致达到隐藏但又能列宽自适应的目的
    Range(a).Select '传入的a是一个带引号的单元格，如："F3"
    With Selection.Font
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
End Function

Function psetting() '显示打印边界函数
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 1200
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ActiveWindow.View = xlNormalView
End Function
