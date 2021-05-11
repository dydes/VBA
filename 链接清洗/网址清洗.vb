Sub clean()

'插入列
Call insert_col("入口页面", "A1:AZ1")
Call insert_col("入口页面", "A1:AZ1")
Call insert_col("最后停留在", "A1:AZ1")
Call insert_col("最后停留在", "A1:AZ1")

'切换到链接工作表并计算最大行数
Sheets("链接对应地址").Select
rowmax = Sheets("链接对应地址").UsedRange.Rows.Count

'构建两个数组，分别是链接数组和注释数组
arr_link = Range("A2:A" & rowmax)
arr_remark = Range("B2:B" & rowmax)

'切换到数据源工作表
Sheets("数据源").Select

'从第1列到最后一个需要补齐数据的列
For i = 1 To Application.Match("是否支持JAVA", Range("A1:AZ1"), 0)
    For j = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(j, i) = "" Then
            Cells(j, i) = Cells(j - 1, i)
        End If
    Next
    Application.StatusBar = "补齐数据" & GetProgress(i, 23)
Next

'清洗链接
For j = 0 To rowmax - 2
    Debug.Print rowmax
    k = arr_link(j + 1, 1)
    l = arr_remark(j + 1, 1)
    Call find_ins("入口页面", k, l)
    Call find_ins("最后停留在", k, l)
    Call find_ins("页面地址", k, l)
    Application.StatusBar = "清洗中" & GetProgress(j, rowmax - 1)
Next

'修改日期格式
Columns(Application.Match("访问时间", Range("A1:AZ1"), 0)).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
Columns(Application.Match("上一次访问时间", Range("A1:AZ1"), 0)).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"

'分列
Call insert_col("访问时间", "A1:AZ1")
Call insert_col("访问时间", "A1:AZ1")

d = Application.Match("访问时间", Range("A1:AZ1"), 0)
Columns(d).Select
d1 = Split(Cells(1, d).Address, "$")(1)
Selection.TextToColumns Destination:=Range(d1 & 1), Space:=True
Columns(d + 2).Delete Shift:=xlToLeft
Cells(1, d) = "访问日期"
Cells(1, d + 1) = "访问时间"

'修改日期格式
Columns(Application.Match("访问日期", Range("A1:AZ1"), 0)).NumberFormatLocal = "yyyy/mm/dd"
Columns(Application.Match("访问时间", Range("A1:AZ1"), 0)).NumberFormatLocal = "hh:mm:ss"

Call insert_col("上一次访问时间", "A1:AZ1")
Call insert_col("上一次访问时间", "A1:AZ1")

e = Application.Match("上一次访问时间", Range("A1:AZ1"), 0)
Columns(e).Select
e1 = Split(Cells(1, e).Address, "$")(1)
Selection.TextToColumns Destination:=Range(e1 & 1), Space:=True
Columns(e + 2).Delete Shift:=xlToLeft
Cells(1, e) = "上一次访问日期"
Cells(1, e + 1) = "上一次访问时间"

'修改日期格式
Columns(Application.Match("上一次访问日期", Range("A1:AZ1"), 0)).NumberFormatLocal = "yyyy/mm/dd"
Columns(Application.Match("上一次访问时间", Range("A1:AZ1"), 0)).NumberFormatLocal = "hh:mm:ss"

'清洗访问时长、停留时长
Columns(Application.Match("访问时长", Range("A1:AZ1"), 0)).Select
Selection.Replace What:="正在访问s", Replacement:="0"
Selection.Replace What:="s", Replacement:=""
    
Columns(Application.Match("停留时长", Range("A1:AZ1"), 0)).Select
Selection.Replace What:="正在访问", Replacement:="0"
Selection.Replace What:="s", Replacement:=""

'填写列标题
Range("A1") = "序号"

f = Application.Match("入口页面", Range("A1:AZ1"), 0)
Cells(1, f + 1) = "入口清洗"
Cells(1, f + 2) = "入口名称"

g = Application.Match("最后停留在", Range("A1:AZ1"), 0)
Cells(1, g + 1) = "最后停留清洗"
Cells(1, g + 2) = "最后停留名称"

h = Application.Match("页面地址", Range("A1:AZ1"), 0)
Cells(1, h + 1) = "页面清洗"
Cells(1, h + 2) = "页面名称"

MsgBox "清洗完成"
End Sub
Function insert_col(a, b) '参数a是需要搜索的列名，如："付款时间"，参数b是查找范围，如："A1:AZ1"'
    Columns(Application.Match(a, Range(b), 0) + 1).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Function

Function find_ins(a, b, c)
'a是标题，如A；b表示要替换的字符串；c是中文备注

m = Application.Match(a, Range("A1:AZ1"), 0) '得到标题的列号
n = Split(Cells(1, m).Address, "$")(1) '得到标题的列标
rowmax1 = Sheets("数据源").UsedRange.Rows.Count

For i = 1 To rowmax1
    If InStr(Range(n & i), b) > 0 Then '假设数据都在A列中，从第一行开始
        Cells(i, m + 1) = b
        Cells(i, m + 2) = c
    End If
Next

End Function

Function GetProgress(curValue, maxValue)
Dim i As Single, j As Integer, s As String
    i = maxValue / 20
    j = curValue / i
    For m = 1 To j
        s = s & "■"
    Next m
    For n = 1 To 20 - j
        s = s & "□"
    Next n
    GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function



