Sub 统计时长()

'删除辅助列
Sheets("数据明细").Range("L:M").Delete Shift:=xlToLeft
Sheets("订单信息").Range("A:D").Delete Shift:=xlToLeft

'获取最大行数
Sheets("数据明细").Select
rowmax = Sheets("数据明细").UsedRange.Rows.Count

'增加辅助列
Sheets("数据明细").Range("L1") = "ID-状态"
Sheets("数据明细").Range("M1") = "ID-驳回"
For i = 2 To rowmax
    Range("L" & i) = Range("A" & i) & "-" & Range("H" & i)
    Range("M" & i) = Range("A" & i) & "-" & Range("G" & i)
Next
Sheets("订单信息").Range("B1") = "科目"
Sheets("订单信息").Range("D1") = "起始时间"
Sheets("订单信息").Range("E1") = "结束时间"

'订单ID复制到订单信息表并去重
    Columns("A:C").Copy
    Sheets("订单信息").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$B$" & rowmax).RemoveDuplicates Columns:=1, Header:=xlYes

''开始判断
'1.在订单信息表里循环订单号判断，如果是人工创建，查找ID-结构对应的起始时间
'2.如果不是人工创建，判断ID-有驳回是否存在，存在找到id-结构对应的起始时间，否则找到ID-审核的起始时间
'3.ID-二校且ID-四校均存在，则比较哪个结束时间晚
'4.ID-二校或ID-四校存在一个，则取那个结束时间
'5.ID-二校或ID-四校均不存在，是否存在等待发布，存在就取，不存在取-审核的时间
'Sheets("数据明细").Select
''那么合并其A列和H列在L列中搜索对应的行号
'For i = 1 To rowmax
'    If Range("G" & i) = "有" Then '如果驳回状态列是“有”
'        id_value = Range("A" & i) '先获取
'        key_value = id_value & "-" & Range("H" & i)
'        rowx = Application.Match(key_value, Range("L:L"), 0)
'        '找到对应的行号，对应的I列是起始时间，在订单信息表中找到该行，填入数据
'        strat_t = Range("I" & rowx)
'        Subject = Range("C" & rowx)
'        rowy = Application.Match(id_value, Sheets("订单信息").Range("A:A"), 0)
'        Sheets("订单信息").Range("B" & rowy) = Subject
'        Sheets("订单信息").Range("C" & rowy) = start_t
'    ElseIf Range("G" & i) = "否" And Range("B" & i) = "订单创建" Then
'
'
'
'

End Sub
