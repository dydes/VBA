Public tNAme, ofile, ofName, cfile, cfName, rowmax, colmax, arr_region
Sub auto_open()

'系统提示
'
    mb = "· 成绩管理员：" + Chr(13) + Chr(13)
    mb = mb + "  1.请从会课教学平台导出“xxxx考试（全年级）-全科”表；" + Chr(13)
    mb = mb + "  2.按组合键“ctrl + J”后，请选择导出的表格；" + Chr(13)
    mb = mb + "  3.请不要更改表头结构及标题，以免影响整个系统的使用。" + Chr(13) + Chr(13)
    mb = mb + "· 学校及年级主管领导：" + Chr(13) + Chr(13)
    mb = mb + "  1.“确定”后，可直接浏览成绩分析；" + Chr(13)
    mb = mb + "  2.根据需要，还可更改“平均成绩一览”“优秀生分布”“分数段统计”中红颜色的数据。" + Chr(13)
    MsgBox mb, , "系统提示"
End Sub

Sub data()

    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知

'获取cj表最大行数列数
    rowmax = Sheets("cj").UsedRange.Rows.Count
    colmax = Sheets("cj").UsedRange.Columns.Count
    
'获取工作簿名称
    tNAme = ThisWorkbook.Name

'清除现有分数数据
    Range("A2:M" & rowmax).ClearContents

'调用选文件函数，输出ofile是带路径的文件名，ofName是不带路径的文件名
    Call file_open_name("请选择年级全科文件", "D:\会通\VBA\唐县一中\")

'计时
    tim1 = Timer

'打开文件
    Workbooks.Open (ofile)

'选择“成绩排名”工作表
    Windows(ofName).Activate
    Sheets("成绩排名").Select

'删除顶部多余行
    top_rows = Application.Match("学生学号", Range("A1:A100"), 0) - 1
        If top_rows > 0 Then
            Rows("1:" & top_rows).Delete Shift:=xlUp
        ElseIf top_rows = 0 Then
            Range("A1").Select
        Else
            MsgBox "未找到学生学号字段，请确认文件是否正确"
        End If

'更新最大行数列数
    rowmax = Sheets("成绩排名").UsedRange.Rows.Count
    colmax = Sheets("成绩排名").UsedRange.Columns.Count
    
'判断内容多少
    Windows(tNAme).Activate
    Sheets("cj").Select
    rowmax1 = Sheets("cj").UsedRange.Rows.Count
    If rowmax1 > rowmax Then
        Rows(rowmax & ":" & rowmax1).Delete Shift:=xlUp
    ElseIf rowmax1 = rowmax Then GoTo 2
    Else
        Range("N" & rowmax1 & ":W" & rowmax1).Copy
        Range("N" & rowmax1 & ":W" & rowmax).Paste
    End If
2
    
'定义科目数组
    arr_sub_huike = Array("学生学号", "姓名", "班级", "总分", "语文", "数学", "英语", "物理", "化学", "生物", "历史", "地理", "政治")
    arr_sub_cj = Array("考号", "姓名", "班级代码", "总成绩", "语文", "数学", "外语", "物理", "化学", "生物", "历史", "地理", "政治")
    
'逐列复制内容
    For i = 0 To 12
        Windows(ofName).Activate
        Sheets("成绩排名").Select
            colx = Application.Match(arr_sub_huike(i), Range("A1:AZ1"), 0)
            If IsNumeric(colx) = False Then GoTo 1
                Range(Cells(2, colx), Cells(rowmax, colx)).Copy
                Windows(tNAme).Activate
                Sheets("cj").Select
                coly = Application.Match(arr_sub_cj(i), Range("A1:AZ1"), 0)
                Cells(2, coly).Select
                Selection.PasteSpecial Paste:=xlPasteValues
1
    Application.StatusBar = "自动粘贴中" & GetProgress(i, 13)
    Next

'关闭源文件且不保存
    Windows(ofName).Activate
    ActiveWorkbook.Close savechanges:=False
    
'替换--
    Call replace("D2:M" & rowmax, "--", 0)

'定义自定义区域名称数组
    Windows(tNAme).Activate
    Sheets("cj").Select
    arr_region = Array("bc", "score", _
                        "_f1", "_f2", "_f3", "_f4", "_f5", "_f6", "_f7", "_f8", "_f9", _
                        "mc", "mc1", "mc2", "mc3", "mc4", "mc5", "mc6", "mc7", "mc8", "mc9", _
                        "bmc", "bmc1", "bmc2", "bmc3", "bmc4", "bmc5", "bmc6", "bmc7", "bmc8", "bmc9")

'清除已有并新建名称
    Call dele_new_vars
    
'新建可调节参数的名称
    ActiveWorkbook.Names.Add Name:="pjrs", RefersToR1C1:="='平均成绩一览'!R2C7:R2C7"
    ActiveWorkbook.Names.Add Name:="yxs", RefersToR1C1:="='优秀生分布'!R2C7:R2C7"
    ActiveWorkbook.Names.Add Name:="fd", RefersToR1C1:="='分数段统计'!R2C11:R2C11"
    ActiveWorkbook.Names.Add Name:="dxf", RefersToR1C1:="='分数段统计'!R2C14:R2C14"

'制作各科班名次()
'
    Sheets("cj").Select
    arr_f = Array("D", "E", "F", "G", "H", "I", "J", "K", "L", "M")
    arr_bmc = Array("X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG")
    arr_loc = Array(-21, -22, -23, -24, -25, -26, -27, -28, -29, -30)

    For i = 0 To 9
        Range("A1:AG" & rowmax).Sort _
            Key1:=Range("C2"), Order1:=xlAscending, _
            Key2:=Range(arr_f(i) & "2"), Order2:=xlDescending, _
            Header:=xlYes
        Range(arr_bmc(i) & 2).FormulaR1C1 = "1"
        Range(arr_bmc(i) & 3).FormulaR1C1 = "=IF(RC[" & arr_loc(i) & "]=R[-1]C[" & arr_loc(i) & "],R[-1]C+1,1)"
        Range(arr_bmc(i) & 3).AutoFill Destination:=Range(arr_bmc(i) & "3:" & arr_bmc(i) & rowmax)
        Range(arr_bmc(i) & 2).Select
        Range(Selection, Selection.End(xlDown)).Copy
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    Application.StatusBar = "计算班级名称" & GetProgress(i, 10)
    Next
    
'回到原记录顺序
'
    Range("A1:AG" & rowmax).Sort _
        Key1:=Range("A2"), Order1:=xlAscending, _
        Header:=xlYes

'设置边框线
    With Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
    End With
    
'设置表头样式
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    Selection.Font.Bold = True

'调整列宽和行高
    Columns("A:AG").EntireColumn.AutoFit
    Rows("1:" & rowmax).EntireRow.AutoFit

'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True '重启刷新
    Application.DisplayAlerts = True '重启通知
    Application.StatusBar = "已完成"
    MsgBox "数据制作完成，谢谢使用！！用时" & Format(using_time, "0.0秒")

End Sub

Function file_open_name(til, ifilname) 'til是文件选择器标题，ifilname是默认打开路径
'先选择文件，获取路径，若未选择任何文件，终止程序，输出ofile是带路径的文件名，ofName是不带路径的文件名
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = til
        .InitialFileName = ifilname
        If .Show Then
            ofile = .SelectedItems(1)
            Else: Exit Function
        End If
    End With
'用斜杠分割文件路径，创建数组，选取数组最后一个元素做为不含路径的文件名
    splfile = Split(ofile, "\")
    ofName = splfile(UBound(splfile))
End Function

Function dele_new_vars()
'根据名称数组，逐个删除已有名称，新建对应名称的区域
    Windows(tNAme).Activate
    For i = 0 To UBound(arr_region)
        ActiveWorkbook.Names(arr_region(i)).Delete
        ActiveWorkbook.Names.Add Name:=arr_region(i), RefersToR1C1:="='cj'!R2C" & i + 3 & ":R" & rowmax & "C" & i + 3
    Application.StatusBar = "自动命名" & GetProgress(i, UBound(arr_region) + 1)
    Next
End Function

Function GetProgress(curValue, maxValue)
Dim i As Single, j As Integer, s As String
    i = maxValue / 20
    j = curValue / i
    For M = 1 To j
        s = s & "■"
    Next M
    For n = 1 To 20 - j
        s = s & "□"
    Next n
    GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function

'rbef表示替换什么，rlat表示替换后是什么，数字字符均可，字符用双引号'
Function replace(rang, rbef, rlat)
    Range(rang).Select
    Selection.replace What:=rbef, Replacement:=rlat, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function
