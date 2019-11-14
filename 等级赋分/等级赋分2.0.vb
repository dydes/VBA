Sub 等级赋分2.0()

'定义文件处理相关变量
    Dim splfile As Variant, fName As String, fPath As String, file As String
'定义运行时间相关变量
    Dim dat As String, tim As String, tim1 As String, tim2 As String
    
    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知

'先选择文件，获取路径，若未选择任何文件，终止程序
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "请选择年级全科文件"
        .InitialFileName = "D:\会通\VBA\等级赋分\"
            If .Show Then
                file = .SelectedItems(1)
            Else: Exit Sub
            End If
    End With

'用斜杠分割文件路径，创建数组，选取数组最后一个元素做为不含路径的文件名
    splfile = Split(file, "\")
    fName = splfile(UBound(splfile))

'选择要保存的文件路径，若未选择任何文件夹，终止程序
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择要保存的文件夹"
        .InitialFileName = "D:\会通\VBA\等级赋分\"
        If .Show = -1 Then
            fPath = .SelectedItems(1)
        Else: Exit Sub
        End If
    End With

'打开指定工作簿
    Workbooks.Open (file)

'获取当前时间
    dat = Format(Date, "yyyy年mm月dd日") '当前年月日
    tim = Format(Time, "hh时mm分ss秒") '当前时间
    tim1 = Timer

'选择“成绩排名”工作表，复制并新建文件，保存新文件，关闭源文件
    Windows(fName).Activate
        Sheets("成绩排名").Select
        Sheets("成绩排名").Copy
    ChDir fPath
        ActiveWorkbook.SaveAs Filename:=fPath & "\等级赋分-" & dat & "-" & tim & "生成.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows(fName).Activate
        ActiveWorkbook.Close savechanges:=False
        
'删除顶部1或2行
    Sheets("成绩排名").Select
    If Range("A3").Value = "学生学号" Then
        Rows("1:2").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
    End If

'求最大行数列数
    colmax = Sheets("成绩排名").UsedRange.Columns.Count
    rowmax = Sheets("成绩排名").UsedRange.Rows.Count

'格式清洗
'替换--
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="--", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'逐列文本转数值
    For i = 1 To colmax
    Cells(2, i).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Cells(2, i), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Next

'加列？？？？
    '查找并增加总分、总分班次、总分级次3列
    col_a = Rows("1:1").Find(What:="总分").Column
    col_a_Name = Split(Cells(1, col_a).Address, "$")(1) 'col_array(col_a - 1) '总分列名
    For i = 1 To 5 Step 2
        Columns(col_a + i).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    '查找并增加英语加权、英语加权年级排名2列
    col_e = Rows("1:1").Find(What:="英语").Column
    col_e_Name = Split(Cells(1, col_e).Address, "$")(1) 'col_array(col_e - 1) '英语列号
    For i = 1 To 3 Step 2
        Columns(col_e + i).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next