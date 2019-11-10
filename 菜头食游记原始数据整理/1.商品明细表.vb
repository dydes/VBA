Sub 数据整理()

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
    
    Range("W1") = "省"
        
        
End Sub
