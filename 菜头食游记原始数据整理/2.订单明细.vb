Sub 订单明细()
Dim rowmax As Integer

'求最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count

'重命名表头
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


'确定结束位置
    a = 109

'粘贴换位置
    Range("E2:F" & a).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("F2").Select
    ActiveSheet.Paste

'填充
    Range("E" & a + 1).Select
    Selection.AutoFill Destination:=Range("E2:E" & a + 1), Type:=xlFillDefault
    Range("H110:L" & a + 1).Select
    Selection.AutoFill Destination:=Range("H2:L" & a + 1), Type:=xlFillDefault
    Range("H2:L" & a + 1).Select
    ActiveWorkbook.Save
End Sub
