Sub 普通订单()

'插入列
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'分列
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

'改标题
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "下单日期"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "下单时间"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "付款日期"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "付款时间"

End Sub

