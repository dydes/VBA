Sub 计算各小时并发峰值()

    
    rowmax = ActiveSheet.UsedRange.Rows.Count

'新增标题
    Range("G1") = "起始小时"
    Range("H1") = "结束小时"
    Range("I1") = "小时数"
    For i = 1 To 5
        Cells(1, 9 + i) = "第" & i & "小时"
    Next
    
'计算起止小时及小时数和各小时
    For i = 2 To rowmax
        Range("G" & i).FormulaR1C1 = "=HOUR(RC[-5])"
        Range("H" & i).FormulaR1C1 = "=HOUR(RC[-5])"
        Range("I" & i).FormulaR1C1 = "=RC[-1]-RC[-2]+1"
        col_num = Range("I" & i)
        If col_num >= 1 Then
            Range("J" & i) = Range("G" & i)
        End If
        If col_num >= 2 Then
            Range("K" & i) = Range("G" & i) + 1
        End If
        If col_num >= 3 Then
            Range("L" & i) = Range("G" & i) + 2
        End If
        If col_num >= 4 Then
            Range("M" & i) = Range("G" & i) + 3
        End If
        If col_num >= 5 Then
            Range("N" & i) = Range("G" & i) + 4
        End If
    Next
    
'新增sheet
    ActiveSheet.Name = "源数据"
    On Error Resume Next
    If Sheets("合并") Is Nothing Then
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "合并"
    Else
        Sheets("合并").Select
        Cells.Select
        Selection.ClearContents
    End If

'计算最多几小时
    Sheets("源数据").Select
    key_col = Application.WorksheetFunction.Max(Range("I2:I" & i))
    
'选择并粘贴
    Application.Wait Now + TimeValue("00:00:01")
    If key_col >= 1 Then
        Sheets("源数据").Range("A1:A" & rowmax).Select
        Selection.Copy
        Sheets("合并").Select
        Range("A1").Select
        ActiveSheet.Paste
        Sheets("源数据").Range("D1:D" & rowmax).Copy
        Sheets("合并").Select
        Range("B1").Select
        ActiveSheet.Paste
        Sheets("源数据").Range("F1:F" & rowmax).Copy
        Sheets("合并").Select
        Range("C1").Select
        ActiveSheet.Paste
        Sheets("源数据").Range("J1:J" & rowmax).Copy
        Sheets("合并").Select
        Range("D1").Select
        ActiveSheet.Paste
    End If
    Application.Wait Now + TimeValue("00:00:01")
    If key_col >= 2 Then
        Sheets("源数据").Range("A2:A" & rowmax).Copy
        Sheets("合并").Select
        Range("A" & rowmax + 1).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("D2:D" & rowmax).Copy
        Sheets("合并").Select
        Range("B" & rowmax + 1).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("F2:F" & rowmax).Copy
        Sheets("合并").Select
        Range("C" & rowmax + 1).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("K2:K" & rowmax).Copy
        Sheets("合并").Select
        Range("D" & rowmax + 1).Select
        ActiveSheet.Paste
    End If
    Application.Wait Now + TimeValue("00:00:01")
    If key_col >= 3 Then
        Sheets("源数据").Range("A2:A" & rowmax).Copy
        Sheets("合并").Select
        Range("A" & rowmax * 2).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("D2:D" & rowmax).Copy
        Sheets("合并").Select
        Range("B" & rowmax * 2).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("F2:F" & rowmax).Copy
        Sheets("合并").Select
        Range("C" & rowmax * 2).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("L2:L" & rowmax).Copy
        Sheets("合并").Select
        Range("D" & rowmax * 2).Select
        ActiveSheet.Paste
    End If
    Application.Wait Now + TimeValue("00:00:01")
    If key_col >= 4 Then
        Sheets("源数据").Range("A2:A" & rowmax).Copy
        Sheets("合并").Select
        Range("A" & rowmax * 3 - 1).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("D2:D" & rowmax).Copy
        Sheets("合并").Select
        Range("B" & rowmax * 3 - 1).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("F2:F" & rowmax).Copy
        Sheets("合并").Select
        Range("C" & rowmax * 3 - 1).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("M2:M" & rowmax).Copy
        Sheets("合并").Select
        Range("D" & rowmax * 3 - 1).Select
        ActiveSheet.Paste
    End If
    Application.Wait Now + TimeValue("00:00:01")
    If key_col = 5 Then
        Sheets("源数据").Range("A2:A" & rowmax).Copy
        Sheets("合并").Select
        Range("A" & rowmax * 4 - 2).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("D2:D" & rowmax).Copy
        Sheets("合并").Select
        Range("B" & rowmax * 4 - 2).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("F2:F" & rowmax).Copy
        Sheets("合并").Select
        Range("C" & rowmax * 4 - 2).Select
        ActiveSheet.Paste
        Sheets("源数据").Range("N2:N" & rowmax).Copy
        Sheets("合并").Select
        Range("D" & rowmax * 4 - 2).Select
        ActiveSheet.Paste
    End If
    Columns("A:D").EntireColumn.AutoFit
    Range("D1") = "小时"
    
'插入透视表
    Columns("A:D").Select
    Range("A334").Activate
    Application.CutCopyMode = False
    Sheets.Add
    ActiveSheet.Name = "透视"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "合并!R1C1:R1048576C4", Version:=6).CreatePivotTable TableDestination:= _
        "透视!R3C1", TableName:="数据透视表1", DefaultVersion:=6
    Sheets("透视").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("时间")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("小时")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("听课人数"), "求和项:听课人数", xlSum

End Sub
