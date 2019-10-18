Sub 普通订单()

'确定结束位置
    a = 109

'粘贴换位置
    Range("H2:BG" & a).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("I2").Select
    ActiveSheet.Paste
    Range("J2:BG" & a).Select
    Selection.Cut
    Range("K2").Select
    ActiveSheet.Paste
    Range("L2:BG" & a).Select
    Selection.Cut
    Range("M2").Select
    ActiveSheet.Paste
    Range("N2:BG" & a).Select
    Selection.Cut
    Range("O2").Select
    ActiveSheet.Paste
    
'分列
    Range("G2:G" & a).Select
    Selection.TextToColumns Destination:=Range("G2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

    Range("I2:I" & a).Select
    Selection.TextToColumns Destination:=Range("I2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    
    Range("K2:K" & a).Select
    Selection.TextToColumns Destination:=Range("K2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Range("M2:M" & a).Select
    Selection.TextToColumns Destination:=Range("M2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

End Sub