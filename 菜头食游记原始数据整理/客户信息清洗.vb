Sub test()
    rowmax = Range("A1").CurrentRegion.Rows.Count
    '省清洗
    Call replace("I2:I" & rowmax, "广西壮族自治区", "广西省")
    Call replace("I2:I" & rowmax, "内蒙古自治区", "内蒙省")
    Call replace("I2:I" & rowmax, "新疆维吾尔自治区", "新疆省")
    Call replace("I2:I" & rowmax, "宁夏回族自治区", "宁夏省")
    Call replace("I2:I" & rowmax, "内蒙古", "内蒙省")
    For i = 2 To rowmax
        Debug.Print "i=" & i
        pos_省 = InStr(Range("I" & i), "省")
        If pos_省 <> 0 Then
            Range("F" & i) = Left(Range("I" & i), pos_省 - 1)
        Else
            Range("F" & i) = Left(Range("I" & i), 2)
        End If
        
        pos_市 = InStr(Range("I" & i), "市")
        If pos_市 <> 0 Then
            Range("G" & i) = Mid(Range("I" & i), pos_省 + 1, pos_市 - pos_省)
        Else
            pos_自治 = InStr(Range("I" & i), "自治")
            Range("G" & i) = Mid(Range("I" & i), pos_省 + 1, pos_自治 - 1)
        End If
        
        pos_区 = InStr(Range("I" & i), "区")
        pos_县 = InStr(Range("I" & i), "县")
        pos_市2 = InStr(Len(Range("F" & i)) + 1 + Len(Range("G" & i)) + 1, Range("I" & i), "市")
        sta = InStr(Range("I" & i), Range("G" & i)) + Len(Range("G" & i))
        If pos_区 = 0 And pos_县 <> 0 Then
            Range("H" & i) = Right(Left(Range("I" & i), pos_县), Len(Left(Range("I" & i), pos_县)) - (sta - 1))
        ElseIf pos_区 <> 0 And pos_县 = 0 Then
            Range("H" & i) = Right(Left(Range("I" & i), pos_区), Len(Left(Range("I" & i), pos_区)) - (sta - 1))
        ElseIf pos_市2 <> 0 Then
            Range("H" & i) = Right(Left(Range("I" & i), pos_市2), pos_市2 - pos_市)
        Else
            Range("H" & i) = Right(Range("I" & i), (Len(Range("I" & i)) - Len(Range("G" & i)) - Len(Range("F" & i)) - 1))
        End If
        
    Next
    
End Sub
'三个参数分别表示，连续的范围如："A1:K33"，替换什么，替换后是什么，数字字符均可，字符用双引号'
Function replace(rang, rbef, rlat)
    Range(rang).Select
    Selection.replace What:=rbef, Replacement:=rlat, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function

