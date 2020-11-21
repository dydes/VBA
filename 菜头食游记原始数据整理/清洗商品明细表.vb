'清洗商品明细数据
Sub 数据整理()
    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知

'求列号
    col_1 = Application.Match("退款金额", Range("A1:AZ1"), 0)
    col_2 = Application.Match("key", Range("A1:AZ1"), 0)
    col_3 = Application.Match("商品简称", Range("A1:AZ1"), 0)
    col_4 = Application.Match("型号简称", Range("A1:AZ1"), 0)
    col_5 = Application.Match("商品+型号", Range("A1:AZ1"), 0)
    col_6 = Application.Match("供货价", Range("A1:AZ1"), 0)
    col_7 = Application.Match("运费", Range("A1:AZ1"), 0)
    col_8 = Application.Match("销售额", Range("A1:AZ1"), 0)
    col_9 = Application.Match("批次", Range("A1:AZ1"), 0)
    col_10 = Application.Match("利润", Range("A1:AZ1"), 0)
    col_11 = Application.Match("商品id", Range("A1:AZ1"), 0)
    col_12 = Application.Match("SKU ID", Range("A1:AZ1"), 0)

'定位新数据截止行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        If Len(Cells(i, col_2)) > 0 Then
            a = i - 1
            Exit For
        End If
    Next
Debug.Print "a=" & a
    
'填充数据
    For i = 2 To a
        If Cells(i, col_1) = "" Then
            Cells(i, col_1) = 0
        End If
    Cells(i, col_9) = 1
    Cells(i, col_2) = LTrim(Str(Cells(i, col_11))) & LTrim(Str(Cells(i, col_12))) & LTrim(Str(Cells(i, col_9)))
    Next
Application.ScreenUpdating = True '重启刷新
MsgBox "完成"
End Sub

