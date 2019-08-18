Sub 普通订单()

'插入列
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("N:N").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'分列
    Columns("G:G").Select
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("I:I").Select
    Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("K:K").Select
    Selection.TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("M:M").Select
    Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'改标题
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "下单日期"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "下单时间"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "付款日期"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "付款时间"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "发货日期"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "发货时间"
    Range("L2").Select
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "买家确认收货时间"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "买家确认收货日期"

End Sub