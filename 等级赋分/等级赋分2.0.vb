Sub 等级赋分()

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

'确定总分位置并插列
    '确定查找范围
    Set title_range = Rows("1:1")
    subject_arr = Array("总分", "物理", "化学", "生物", "历史", "地理", "政治", "通用技术", "信息技术")
    
    '确定总分位置
    Rows("1:1").Select
    Selection.Find(What:="总分", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
    col_sa = ActiveCell.Column
    
    '插入总分相关列
    For i = 1 To 5 Step 2
        Columns(col_sa + i).Select
        Selection.Insert Shift:=xlToRight
    Next
    
    '重命名总分相关列
    Cells(1, col_sa) = "原始总分"
    Cells(1, col_sa + 1) = "赋分总分"
    Cells(1, col_sa + 2) = "原始班次"
    Cells(1, col_sa + 3) = "赋分班次"
    Cells(1, col_sa + 4) = "原始级次"
    Cells(1, col_sa + 5) = "赋分级次"
    
'确定各科位置并插列
    '逐个确定小学科位置
    Set title_range = Rows("1:1")
    For i = 1 To 8
        subject_is_exist = Application.WorksheetFunction.CountIf(title_range, subject_arr(i))
        If subject_is_exist > 0 Then
            Rows("1:1").Select
            Selection.Find(What:=subject_arr(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, MatchByte:=False, SearchFormat:=False).Activate
            col_sn = ActiveCell.Column
            name_sn = ActiveCell.Value
            '插入小学科相关列
            For h = 1 To 3 Step 2
                Columns(col_sn + h).Select
                Selection.Insert Shift:=xlToRight
            Next
            Columns(col_sn + 1).Select
            Selection.Insert Shift:=xlToRight
            '重命名总分相关列
            Cells(1, col_sn) = "原始" & name_sn
            Cells(1, col_sn + 1) = name_sn & "等级"
            Cells(1, col_sn + 2) = "赋分" & name_sn
            Cells(1, col_sn + 3) = "原始级次"
            Cells(1, col_sn + 4) = "赋分级次"
        End If
    Next




'备份，这个需要等列都加完了之后，确定相关科目位置，用来计算总分
    Set title_range = Rows("3:3")
    subject_arr = Array("总分", "物理", "化学", "生物", "历史", "地理", "政治", "通用技术", "信息技术")
    subject_score_arr = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    For i = 0 To 8
        subject_score_arr(i) = Application.WorksheetFunction.CountIf(title_range, subject_arr(i))
    Next
    '确定已存在科目列数、列标及标题内容
    subject_score_col_arr = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    subject_score_cname_arr = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    subject_score_tname_arr = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    For i = 0 To 8
        If subject_score_arr(i) <> 0 Then
            Rows("3:3").Select
            Selection.Find(What:=subject_arr(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, MatchByte:=False, SearchFormat:=False).Activate
                subject_score_col_arr(i) = ActiveCell.Column
                subject_score_cname_arr(i) = Split(ActiveCell.Address, "$")(1)
                subject_score_tname_arr(i) = ActiveCell.Value
        End If
    Next
    '清洗科目数组，得到不含0的列数、列标及标题数组
    subject_score_col_isNotNull = Join(subject_score_col_arr, ",")
    subject_score_cname_isNotNull = Join(subject_score_cname_arr, ",")
    subject_score_tname_isNotNull = Join(subject_score_tname_arr, ",")
    subject_score_col_isNotNull = Replace(subject_score_col_isNotNull, ",0", "")
    subject_score_cname_isNotNull = Replace(subject_score_cname_isNotNull, ",0", "")
    subject_score_tname_isNotNull = Replace(subject_score_tname_isNotNull, ",0", "")
    subject_col_arr = Split(subject_score_col_isNotNull, ",")
    subject_colname_arr = Split(subject_score_cname_isNotNull, ",")
    subject_til_arr = Split(subject_score_tname_isNotNull, ",")
    MsgBox Join(subject_col_arr, ",") & Chr(13) & Join(subject_colname_arr, ",") & Chr(13) & Join(subject_til_arr, ",")