Sub 商品明细()

'确定结束位置
    a = 110

'粘贴换位置
    Range("E2:AK" & a).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("F2").Select
    ActiveSheet.Paste
    Range("G2:AK" & a).Select
    Selection.Cut
    Range("H2").Select
    ActiveSheet.Paste

    
'分列
    Range("D2:D" & a).Select
    Selection.TextToColumns Destination:=Range("D2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

    Range("F2:F" & a).Select
    Selection.TextToColumns Destination:=Range("F2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    
End Sub

