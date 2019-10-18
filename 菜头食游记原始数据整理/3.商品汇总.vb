Sub 商品汇总()
    a = 209
    b = 222
    
    Range("C" & a & ":G" & b).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("D" & a).Select
    ActiveSheet.Paste
    Range("F" & a & ":H" & b).Select
    Selection.Cut
    Range("I" & a).Select
    ActiveSheet.Paste
    Range("F" & a - 1).Select
    Selection.AutoFill Destination:=Range("F" & a - 1 & ":F" & b)
    Range("F" & a - 1 & ":F" & b).Select
    Range("G" & a - 1).Select
    Selection.AutoFill Destination:=Range("G" & a - 1 & ":G" & b)
    Range("G" & a - 1 & ":G" & b).Select
    Range("H" & a - 1).Select
    Selection.AutoFill Destination:=Range("H" & a - 1 & ":H" & b)
    Range("H" & a - 1 & ":H" & b).Select
    Range("L" & a - 1 & ":M" & a - 1).Select
    Selection.AutoFill Destination:=Range("L" & a - 1 & ":M" & b)
    Range("L" & a - 1 & ":M" & b).Select
End Sub