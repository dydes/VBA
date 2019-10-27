Sub 数据整理()

'省份还没清洗，如果区是空的就把前两列内容移过来，目前都是直辖市，有一些街道名得注意，有北京和北京市，得统一
'原价得放到后面的表里一块处理

Dim a As Integer

'定位新数据截止行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        If Len(Cells(i, 4).Value) < 12 Then
            a = i - 1
            Exit For
        End If
    Next

'剪切数据块，分列
    Sheets("excelReport").Select
    Range("E2:AL" & a).Cut
    Range("F2").Select
    ActiveSheet.Paste
    Range("D2:D" & a).Select
    Selection.TextToColumns Destination:=Range("D2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    
    Range("G2:AM" & a).Cut
    Range("H2").Select
    ActiveSheet.Paste
    Range("F2:F" & a).Select
    Selection.TextToColumns Destination:=Range("F2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

'型号去重
    '复制新增订单的商品信息
    Range("H2:J" & a).Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H2:J" & a).Copy
    '粘贴到商品信息表的末尾
    Sheets("ProductCode").Select
    rowmax1 = ActiveSheet.UsedRange.Rows.Count
    Range("A" & rowmax1 + 1).Select
    ActiveSheet.Paste
    '商品信息去重
    rowmax2 = ActiveSheet.UsedRange.Rows.Count
    For i = rowmax1 + 1 To rowmax2
        Range("D" & i) = Range("C" & i) & "||" & Range("B" & i)
    Next
    Columns("D:D").Select
    Range("D" & rowmax1).Activate
    ActiveSheet.Range("$A$1:$G$" & rowmax2).RemoveDuplicates Columns:=4, Header:=xlYes
    
'型号排序
    rowmax3 = ActiveSheet.UsedRange.Rows.Count
    Columns("A:G").Select
    ActiveWorkbook.Worksheets("productCode").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("productCode").Sort.SortFields.Add2 Key:=Range( _
        "D2:D" & rowmax3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("productCode").Sort
        .SetRange Range("A1:G" & rowmax3)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    MsgBox "完成"
End Sub

Sub 移动数据()
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        If Range("Y" & i) = "" Then
            Range("Y" & i) = Range("X" & i)
            Range("X" & i) = Range("W" & i)
        End If
    Next
    
    Columns("W:W").Select
    Selection.Replace What:="省", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("W:W").Select
    Selection.Replace What:="市", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("W:W").Select
    Selection.Replace What:="维吾尔自治区", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("W:W").Select
    Selection.Replace What:="自治区", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("X:X").Select
    Selection.Replace What:="上海", Replacement:="上海市", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("X:X").Select
    Selection.Replace What:="北京", Replacement:="北京市", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("X:X").Select
    Selection.Replace What:="重庆", Replacement:="重庆市", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
        
End Sub




