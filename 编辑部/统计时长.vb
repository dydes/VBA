Sub 统计时长()
'1.在订单信息表里循环订单号判断
'2.判断起始时间
'   2.1 如果是人工创建，查找ID-结构对应的起始时间
'   2.2 如果是订单创建，判断ID-有驳回是否存在
'       2.2.1 有驳回，判断ID-结构是否存在
'           2.2.1.1 结构存在则结构对应的起始时间
'           2.2.1.2 结构不存在找到审核对应的起始时间
'       2.2.2 无驳回，否则找到ID-审核的起始时间
'3.判断结束时间
'   3.1 ID-二校且ID-四校均存在，则比较哪个结束时间晚
'   3.2 ID-二校或ID-四校存在一个，则取那个结束时间
'   3.3 ID-二校或ID-四校均不存在
'       3.3.1 是否存在等待发布，存在就取
'       3.3.2 不存在就看录题是否存在，存在就取
'       3.3.3 录题不存在取-审核的时间

Application.ScreenUpdating = False '暂停刷新
Application.DisplayAlerts = False '暂停通知

'删除辅助列（删除上一次的数据）
Sheets("数据明细").Range("L:M").Delete Shift:=xlToLeft
Sheets("订单信息").Range("A:G").Delete Shift:=xlToLeft

'获取数据明细表最大行数
rowmax1 = Sheets("数据明细").UsedRange.Rows.Count

'数据明细表增加辅助列
Sheets("数据明细").Range("L1") = "ID-状态"
Sheets("数据明细").Range("M1") = "ID-驳回"
Sheets("数据明细").Select
For i = 2 To rowmax1
    Range("L" & i) = Range("A" & i) & "-" & Range("H" & i)
    Range("M" & i) = Range("A" & i) & "-" & Range("G" & i)
Next

'订单信息表增加辅助列
Sheets("订单信息").Range("B1") = "科目"
Sheets("订单信息").Range("D1") = "起始时间"
Sheets("订单信息").Range("E1") = "结束时间"
Sheets("订单信息").Range("F1") = "时长-小时"
Sheets("订单信息").Range("G1") = "时长等级"

'订单ID复制到订单信息表并去重
Sheets("数据明细").Columns("A:C").Copy
Sheets("订单信息").Select
Range("A1").Select
ActiveSheet.Paste
Columns("A:A").Select
Application.CutCopyMode = False
ActiveSheet.Range("$A$1:$C$" & rowmax1).RemoveDuplicates Columns:=1, Header:=xlYes

'获取订单信息表最大行数
rowmax2 = Sheets("订单信息").UsedRange.Rows.Count

'填充时间、计算时长、计算时长等级
Sheets("订单信息").Select
Dim rng1, rng2, rng3, rng4 As Range
For i = 2 To rowmax2 '循环订单信息表中去重后的订单号
    Id_value = Sheets("订单信息").Range("A" & i) '逐个订单ID判断
    key1 = Id_value & "-结构" 'ID-结构
    key2 = Id_value & "-审核" 'ID-审核
    key3 = Id_value & "-有" 'ID-有
    key4 = Id_value & "-二校" 'ID-二校
    key5 = Id_value & "-四校" 'ID-四校
    key6 = Id_value & "-等待发布任务" 'ID-等待发布任务
    key7 = Id_value & "-结构" 'ID-结构
    key8 = Id_value & "-录题" 'ID-录题
    Set rng1 = Sheets("数据明细").Range("M:M").Find(key3, lookat:=xlWhole) '在数据明细表中查找ID-驳回列，看是否能找到ID-有
    Set rng2 = Sheets("数据明细").Range("L:L").Find(key4, lookat:=xlWhole) '在数据明细表中查找ID-状态列，看是否能找到ID-二校
    Set rng3 = Sheets("数据明细").Range("L:L").Find(key5, lookat:=xlWhole) '在数据明细表中查找ID-状态列，看是否能找到ID-四校
    Set rng4 = Sheets("数据明细").Range("L:L").Find(key6, lookat:=xlWhole) '在数据明细表中查找ID-状态列，看是否能找到ID-等待发布任务
    set rng5 = Sheets("数据明细").Range("L:L").Find(key7, lookat:=xlWhole) '在数据明细表中查找ID-状态列，看是否能找到ID-结构
    set rng6 = Sheets("数据明细").Range("L:L").Find(key8, lookat:=xlWhole) '在数据明细表中查找ID-状态列，看是否能找到ID-录题
    If Sheets("订单信息").Range("B" & i) = "订单创建" And rng1 Is Nothing Then
        rowx = Application.Match(key2, Sheets("数据明细").Range("L:L"), 0) '如果是订单创建且没有找到驳回，取ID-审核的行号
    elseif Sheets("订单信息").Range("B" & i) = "订单创建" And not rng1 Is Nothing and rng5 is nothing Then
        rowx = Application.Match(key2, Sheets("数据明细").Range("L:L"), 0) '如果是订单创建有找到驳回但没结构，也取ID-审核的行号
    Else
        rowx = Application.Match(key1, Sheets("数据明细").Range("L:L"), 0) '其他情况取ID-结构的行号
    End If
    start_t = Sheets("数据明细").Range("I" & rowx) '根据行号取起始时间
    If rng2 Is Nothing And rng3 Is Nothing Then '如果二校四校都不存在
        if rng4 is nothing and rng6 is nothing then
            rowy = Application.Match(key2, Sheets("数据明细").Range("L:L"), 0) '如果等待发布任务和录题都不存在，取ID-审核的行号
        elseif rng4 is nothing and not rng6 is nothing then
            rowy = Application.Match(key8, Sheets("数据明细").Range("L:L"), 0) '如果等待发布任务不存在，录题存在，取ID-录题的行号
        else
            rowy = Application.Match(key6, Sheets("数据明细").Range("L:L"), 0) '如果等待发布任务存在，取ID-等待发布任务的行号
        endif
    ElseIf rng2 Is Nothing And Not rng3 Is Nothing Then
        rowy = Application.Match(key5, Sheets("数据明细").Range("L:L"), 0) '二校不存在、四校存在，取ID-四校的行号
    ElseIf Not rng2 Is Nothing And rng3 Is Nothing Then
        rowy = Application.Match(key4, Sheets("数据明细").Range("L:L"), 0) '二校存在、四校不存在，取ID-二校的行号
    ElseIf Not rng2 Is Nothing And Not rng3 Is Nothing Then
        row_2j = Application.Match(key4, Sheets("数据明细").Range("L:L"), 0) '取ID-二校的行号
        row_4j = Application.Match(key5, Sheets("数据明细").Range("L:L"), 0) '取ID-四校的行号
        et_2j = Sheets("数据明细").Range("I" & row_2j) '取二校结束时间
        et_4j = Sheets("数据明细").Range("I" & row_4j) '取四校结束时间
        If et_2j > et_4j Then
            rowy = row_2j
        Else
            rowy = row_4j
        End If
    End If
    end_t = Sheets("数据明细").Range("J" & rowy) '根据行号取结束时间
    '填充时间
    Sheets("订单信息").Range("D" & i) = start_t '填充起始时间
    Sheets("订单信息").Range("E" & i) = end_t '填充起始时间
    dur_t = Application.Round((end_t - start_t) * 24, 2)
    Sheets("订单信息").Range("F" & i) = dur_t '时长=结束时间-起始时间
    '判断时长等级
    Select Case dur_t
        Case Is <=24
            Sheets("订单信息").Range("G" & i) = "时长≤24h"
        Case Is <=48
            Sheets("订单信息").Range("G" & i) = "24h<时长≤48h"
        Case Is <=72
            Sheets("订单信息").Range("G" & i) = "48h<时长≤72h"
        Case Is >72
            Sheets("订单信息").Range("G" & i) = "时长>72h"
    End Select
    Application.StatusBar = "整体进度" & GetProgress(i, rowmax2)
Next

'自适应列宽
sheets("订单信息").Columns("A:G").EntireColumn.AutoFit

'数据统计
sheets("数据看板").Select
sheets("数据看板").PivotTables("数据透视表1").PivotCache.Refresh
arr=Array("时长≤24h","24h<时长≤48h","48h<时长≤72h","时长>72h")
brr = Array("语文","数学","英语","物理","化学","生物","政治","地理","历史","通用技术","信息技术","理科综合","文科综合","技术")
Dim rngc, rngs As Range
for i=0 to 3
    keyc = arr(i)
    Set rngc = Sheets("订单信息").Range("G:G").Find(keyc, lookat:=xlWhole) '在订单信息表中查找时长等级
    If not rngc Is Nothing then
        sheets("数据看板").PivotTables("数据透视表1").PivotFields("时长等级").PivotItems(keyc).Position = i+1
    endif
Next
for i=0 to 13
    keys = brr(i)
    Set rngs = Sheets("订单信息").Range("C:C").Find(keys, lookat:=xlWhole) '在订单信息表中查找科目
    If not rngs Is Nothing then
        sheets("数据看板").PivotTables("数据透视表1").PivotFields("科目").PivotItems(keys).Position = i+1
    endif
Next

'处理完成
ActiveWorkbook.Save
    
Application.ScreenUpdating = True '重启刷新
Application.DisplayAlerts = True '重启通知

MsgBox "操作完成"
Application.StatusBar = False

End Sub

Function GetProgress(curValue, maxValue)
Dim i As Single, j As Integer, s As String
i = maxValue / 20
j = curValue / i
'进度条
For m = 1 To j
    s = s & "■"
Next
For n = 1 To 20 - j
    s = s & "□"
Next
GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function


