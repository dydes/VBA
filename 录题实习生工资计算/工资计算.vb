Sub 工资计算()

'定义变量
    Dim splfile As Variant, fName As String, fPath As String, file As String
    Dim dat As String, tim As String, tim1 As String, tim2 As String
    Dim rowmax As Integer, rowmax1 As Integer, rowmax2 As Integer, rowmax3 As Integer, rowmax4 As Integer
    Dim pid1 As double, pid2 As double, man_hour As double, fee_per_day As double, fee_per_hour As double

'读取配置
    pid1 = Sheets("说明").Range("B2").value
    pid2 = Sheets("说明").Range("B3").value
    man_hour = Sheets("说明").Range("B4").value
    fee_per_day = Sheets("说明").Range("B5").value
    fee_per_hour = Sheets("说明").Range("B6").value

'先选择文件，获取路径，若未选择任何文件，终止程序
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "请选择工资文件"
        .InitialFileName = "C:\"
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
        .InitialFileName = "C:\"
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

'选择“每日统计”工作表，复制并新建文件，保存新文件，关闭源文件
    Windows(fName).Activate
        Sheets("每日统计").Select
        Sheets("每日统计").Copy
    ChDir fPath
        ActiveWorkbook.SaveAs Filename:=fPath & "\实习生工资表-" & dat & "-" & tim & "生成.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows(fName).Activate
        ActiveWorkbook.Close savechanges:=False

'删除顶部2行
    Sheets("每日统计").Select
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp

'插入“计费方式”、“人名+计费方式”式2列
    Columns("B:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'插入“来时段”、“走时段”、“时间差”、“实际工时”4列
    Columns("N:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'合并单元格、填充标题、标色
    Merge_list = Array("B", "C", "N", "O", "P", "Q")
    Title = Array("计费方式", "key", "来时段", "走时段", "时间差", "实际工时")
    h = LBound(Merge_list)
    For Each i In Merge_list
        Range(Cells(1, i), Cells(2, i)).Merge
        Range(Cells(1, i), Cells(2, i)) = Title(h)
        Range(Cells(1, i), Cells(2, i)).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
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
        h = h + 1
    Next

'获取最大行数去除底色
    rowmax = ActiveSheet.UsedRange.Rows.Count
    Range("N3:Q" & rowmax).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'判断是否有空值，是否有“次日 ”

'判断来时段、走时段、计算时间差、实际工时、判断计费方式、拼接人员计费类型
    Sheets("每日统计").Range("B3:B" & rowmax).Formula = "=IF(Q3>=" & man_hour & ",1,2)"
    Sheets("每日统计").Range("C3:C" & rowmax).Formula = "=A3&B3"
    Sheets("每日统计").Range("N3:N" & rowmax).Formula = "=IF(J3="""",0,IF((HOUR(J3)+MINUTE(J3)/60)<="& pid1 &",1,IF(AND((HOUR(J3)+MINUTE(J3)/60)>"& pid1 &",(HOUR(J3)+MINUTE(J3)/60)<="& pid2 &"),2,3)))"
    Sheets("每日统计").Range("O3:O" & rowmax).Formula = "=IF(L3="""",0,IF((HOUR(L3)+MINUTE(L3)/60)<=12,1,IF(AND((HOUR(L3)+MINUTE(L3)/60)>12,(HOUR(L3)+MINUTE(L3)/60)<="& pid2 &"),2,3)))"
    Sheets("每日统计").Range("P3:P" & rowmax).Formula = "=value(Text((HOUR(L3)+MINUTE(L3)/60)-(HOUR(J3)+MINUTE(J3)/60),""0.0""))"
    Sheets("每日统计").Range("Q3:Q" & rowmax).Formula = "=VALUE(TEXT(IF(N3&O3=""11"",P3,IF(N3&O3=""12"",P3-1,IF(N3&O3=""13"",P3-2,IF(N3&O3=""22"",P3,IF(N3&O3=""23"",P3-1,IF(N3&O3=""33"",P3,0)))))),""0.0""))"

'创建两个工作表
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "工资第1步"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "工资第2步"

'粘贴姓名、计费方式、key三列到工资第1步
    Sheets("每日统计").Select
    Columns("A:C").Select
    Selection.Copy
    Sheets("工资第1步").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False

'按人员+计费方式列去重
    rowmax1 = ActiveSheet.UsedRange.Rows.Count
    Range("A1:C2").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
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
        .WrapText = True
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
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("C:C").Select
    ActiveSheet.Range("$A$1:$C$" & rowmax1).RemoveDuplicates Columns:=3, Header:=xlYes

    rowmax2 = ActiveSheet.UsedRange.Rows.Count

'插入“计费单价”、“计费基数”、“费用总计”3列，填充标题
    Merge_list1 = Array("D", "E", "F")
    Title1 = Array("计费单价", "计费基数", "费用总计")
    h = LBound(Merge_list1)
    For Each i In Merge_list1
        Cells(1, i) = Title1(h)
        Cells(1, i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        h = h + 1
    Next
    Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("A1:F" & rowmax2).Select
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

'判断计费单价、计费基数、费用总计
    Sheets("工资第1步").Range("D2:D" & rowmax2).Formula = "=IF(B2=2," & fee_per_hour & "," & fee_per_day & ")"
    Sheets("工资第1步").Range("E2:E" & rowmax2).Formula = "=IF(B2=1,TEXT(COUNTIF(每日统计!C:C,""=""&工资第1步!C2),0)&"" 天"",TEXT(SUMIF(每日统计!C:C,""=""&工资第1步!C2,每日统计!Q:Q),0)&"" 时"")"
    Sheets("工资第1步").Range("F2:F" & rowmax2).Formula = "=IF(B2=1,COUNTIF(每日统计!C:C,""=""&工资第1步!C2)*150,SUMIF(每日统计!C:C,""=""&工资第1步!C2,每日统计!Q:Q)*15)"

'粘贴姓名列到工资第2步
    Sheets("工资第1步").Select
    Columns("A").Select
    Selection.Copy
    Sheets("工资第2步").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False

'姓名列去重
    rowmax3 = ActiveSheet.UsedRange.Rows.Count
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$A$" & rowmax3).RemoveDuplicates Columns:=1, Header:=xlYes

    rowmax4 = ActiveSheet.UsedRange.Rows.Count

'增加最终工资列
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "最终工资"
    Range("B1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("A1:B" & rowmax4).Select
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

'计算最终工资
    Sheets("工资第2步").Range("B2:B" & rowmax4).Formula = "=SUMIF(工资第1步!A:A,""=""&工资第2步!A2,工资第1步!F:F)"

'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
ActiveWindow.WindowState = xlMaximized
MsgBox "计算完成，用时" & Format(using_time, "0.0秒")

End Sub