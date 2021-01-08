Sub test()
    Call ad_to_los(1, Range("A1").CurrentRegion.Rows.Count)
End Sub
Function ad_to_los(a, b)
    Application.DisplayAlerts = False '暂停通知
    For i = a To b
        Range("C" & i).Formula = "=FILTERXML(WEBSERVICE(""https://restapi.amap.com/v3/geocode/geo?address=" & Range("A" & i) & "&city=" & Range("B" & i) & "&output=XML&key=fb0ec65db400a59b31bb9a059b22d378""),""//location"")"
        Range("C" & i).Copy
        Range("C" & i).PasteSpecial Paste:=xlPasteValues
        Range("C" & i).TextToColumns Destination:=Range("C" & i), Comma:=True
    Call delay(0.01)
    Next
    Application.DisplayAlerts = True '重启通知
End Function

Function delay(ts)
    Dim t, t1
    t = Timer
    Do
        t1 = Timer
        If t1 < t Then t1 = 86400 + t1
        DoEvents
    Loop Until t1 - ts > t
End Function