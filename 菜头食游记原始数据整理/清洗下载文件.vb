'清洗下载的文件
Sub clean_download()
'删除不需要的列
    Call del_col_key("订单状态", "A1:AZ1")
    Call del_col_key("订单类型", "A1:AZ1")
    Call del_col_key("型号编码", "A1:AZ1")
    Call del_col_key("商品编码", "A1:AZ1")
    Call del_col_key("发货状态", "A1:AZ1")
    Call del_col_key("发货方", "A1:AZ1")
    Call del_col_key("物流公司", "A1:AZ1")
    Call del_col_key("物流单号", "A1:AZ1")
    Call del_col_key("订单描述", "A1:AZ1")
    Call del_col_key("买家留言", "A1:AZ1")
    Call del_col_key("下单模板信息", "A1:AZ1")
    Call del_col_key("备注", "A1:AZ1")
    Call del_col_key("分销商店铺ID", "A1:AZ1")
    Call del_col_key("分销店铺名称", "A1:AZ1")
    Call del_col_key("分销商注册姓名", "A1:AZ1")
    Call del_col_key("分销商手机号", "A1:AZ1")
    Call del_col_key("分销佣金", "A1:AZ1")
    Call del_col_key("是否已成团", "A1:AZ1")
    Call del_col_key("身份证号", "A1:AZ1")
    Call del_col_key("支付方式", "A1:AZ1")
    Call del_col_key("是否自提", "A1:AZ1")
    Call del_col_key("网店名称", "A1:AZ1")
    Call del_col_key("代理店铺ID", "A1:AZ1")
    Call del_col_key("代理店铺名称", "A1:AZ1")
    
'定位商品id与SKU ID列标
    col_sp = Application.Match("商品id", Range("A1:AZ1"), 0)
    col_sp_A = Split(Cells(1, col_sp).Address, "$")(1)
    col_sku_A = Split(Cells(1, col_sp + 1).Address, "$")(1)
    col_sn_A = Split(Cells(1, col_sp - 2).Address, "$")(1)
    col_skun_A = Split(Cells(1, col_sp - 1).Address, "$")(1)
    
'删除商品id或SKUID为空的行
    Call del_row_key("", col_sp_A & ":" & col_sp_A)
    Call del_row_key("", col_sku_A & ":" & col_sku_A)
    Call del_row_key("", col_sn_A & ":" & col_sn_A)
    
'插入列
    Columns(Application.Match("付款时间", Range("A1:AZ1"), 0)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns(Application.Match("商品名称", Range("A1:AZ1"), 0)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'定位到对应列修改标题
    col_1 = Application.Match("下单时间", Range("A1:AZ1"), 0)
    Cells(1, col_1) = "下单日期"
    Cells(1, col_1 + 2) = "付款日期"
'得出列标
    col_1_A = Split(Cells(1, col_1).Address, "$")(1)
    col_2_A = Split(Cells(1, col_1 + 2).Address, "$")(1)
'分列函数
    Call col_divide(col_1_A, 4)
    Call col_divide(col_2_A, 4)
'修改标题
    Cells(1, col_1 + 1) = "下单时间"
    Cells(1, col_1 + 3) = "付款时间"
'求最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
'定位省市区三列
    col_3 = Application.Match("省", Range("A1:AZ1"), 0)
    col_3_A = Split(Cells(1, col_3).Address, "$")(1) '省
    col_4_A = Split(Cells(1, col_3 + 1).Address, "$")(1) '市
    col_5_A = Split(Cells(1, col_3 + 2).Address, "$")(1) '区
'循环平移数据
    For i = 1 To rowmax
        If Range(col_3_A & i) = "北京" _
            Or Range(col_3_A & i) = "北京市" _
            Or Range(col_3_A & i) = "天津市" _
            Or Range(col_3_A & i) = "天津" _
            Or Range(col_3_A & i) = "上海市" _
            Or Range(col_3_A & i) = "上海" _
            Or Range(col_3_A & i) = "广州市" _
            Or Range(col_3_A & i) = "广州" _
            Or Range(col_3_A & i) = "重庆市" _
            Or Range(col_3_A & i) = "重庆" Then
            Range(col_5_A & i) = Range(col_4_A & i)
            Range(col_4_A & i) = Range(col_3_A & i)
        End If
    Next
'省清洗
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "省", "")
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "市", "")
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "广西壮族自治区", "广西")
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "内蒙古自治区", "内蒙")
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "新疆维吾尔自治区", "新疆")
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "宁夏回族自治区", "宁夏")
    Call replace(col_3_A & "2:" & col_3_A & rowmax, "内蒙古", "内蒙")
'市清洗
    Call replace(col_4_A & "2:" & col_4_A & rowmax, "市", "")
MsgBox "完成"
End Sub
Function del_col_key(a, b) '参数a表示要查找的中文，如："订单状态"，参数b表示一个查询范围，如："A1:AZ1"
Do
  colx = Application.Match(a, Range(b), 0)
  If IsNumeric(colx) = False Then Exit Do
    Columns(colx).Delete
Loop
End Function
Function del_row_key(a, b) '参数a表示要查找的中文，如"订单状态"，参数b表示一个查询范围，如："A:AZ"
Do
  rowx = Application.Match(a, Range(b), 0)
  If IsNumeric(rowx) = False Then Exit Do
    Rows(rowx).Delete
Loop
End Function
'两个参数分别表示，带引号单列列标，如："A"；分隔符，1用tab，2用分号，3用逗号，4用空格，5用其他需要加引号，如："|"'
Function col_divide(a, c)
    d = a & ":" & a
    b = a & 1
    Columns(d).Select
    If c = 1 Then
        Selection.TextToColumns Destination:=Range(b), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
            Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf c = 2 Then
        Selection.TextToColumns Destination:=Range(b), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
            Tab:=False, Semicolon:=True, Comma:=False, Space:=False, Other:=False, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf c = 3 Then
        Selection.TextToColumns Destination:=Range(b), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
            Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf c = 4 Then
        Selection.TextToColumns Destination:=Range(b), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
            Tab:=False, Semicolon:=False, Comma:=False, Space:=True, Other:=False, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf c <> 1 Or c <> 2 Or c <> 3 Or c <> 4 Then
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
            Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:=c, _
            FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    End If
End Function
'三个参数分别表示，连续的范围如："A1:K33"，替换什么，替换后是什么，数字字符均可，字符用双引号'
Function replace(rang, rbef, rlat)
    Range(rang).Select
    Selection.replace What:=rbef, Replacement:=rlat, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function


