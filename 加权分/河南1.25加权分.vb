Sub 河南加权分()

'定义文件处理相关变量
    Dim splfile As Variant, fName As String, fPath As String, file As String
'定义运行时间相关变量
    Dim dat As String, tim As String, tim1 As String, tim2 As String
'定义最大行数相关变量
    Dim rowmax As Integer, colmax As Integer
'定义新增列相关变量
    Dim col_a As Integer, col_e As Integer

'先选择文件，获取路径，若未选择任何文件，终止程序
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "请选择年级全科文件"
        .InitialFileName = "D:\会通\VBA\加权分\"
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
        .InitialFileName = "D:\会通\VBA\加权分\"
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
        ActiveWorkbook.SaveAs Filename:=fPath & "\年级全科1.25加权后-" & dat & "-" & tim & "生成.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows(fName).Activate
        ActiveWorkbook.Close savechanges:=False

'删除顶部1行
    Sheets("成绩排名").Select
    If Range("A3").Value = "学生学号" Then
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
    Else
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
    End If

'文本转数值格式
    colmax = ActiveSheet.UsedRange.Columns.Count
    For i = 1 To colmax
    Cells(2, i).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Cells(2, i), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Next

'列名转换数组
    col_array = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ")

'查找并新增1.25分相关列
'总分、总分班次、总分级次
    col_a = Rows("1:1").Find(What:="总分").Column
    For i = 1 To 5 Step 2
        Columns(col_a + i).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    col_c_Name = col_array(col_a - 1 - 1)
    col_a_Name = col_array(col_a - 1)
    col_a125_Name = col_array(col_a + 1 - 1)
    col_a_cr_Name = col_array(col_a + 2 - 1)
    col_a125_cr_Name = col_array(col_a + 3 - 1)
    col_a_gr_Name = col_array(col_a + 4 - 1)
    col_a125_gr_Name = col_array(col_a + 5 - 1)
    Range(col_a125_Name & "1") = "总分加权"
    Range(col_a125_cr_Name & "1") = "加权班级排名"
    Range(col_a125_gr_Name & "1") = "加权年级排名"


'英语、英语级次
    col_e = Rows("1:1").Find(What:="英语").Column
    For i = 1 To 3 Step 2
        Columns(col_e + i).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    col_e_Name = col_array(col_e - 1)
    col_e125_Name = col_array(col_e + 1 - 1)
    col_e_gr_Name = col_array(col_e + 2 - 1)
    col_e125_gr_Name = col_array(col_e + 3 - 1)
    Range(col_e125_Name & "1") = "英语加权"
    Range(col_e125_gr_Name & "1") = "加权年级排名"


'计算英语学科加权分及排名
    rowmax = ActiveSheet.UsedRange.Rows.Count
    Sheets("成绩排名").Range(col_e125_Name & "2:" & col_e125_Name & rowmax).Formula = "=ROUNDUP(IFERROR(" & col_e_Name & "2*1.25,0),0)"
        Cells(2, col_e125_Name).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.TextToColumns Destination:=Cells(2, col_e125_Name), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Sheets("成绩排名").Range(col_e125_gr_Name & "2:" & col_e125_gr_Name & rowmax).Formula = "=RANK(" & col_e125_Name & "2,$" & col_e125_Name & "$2:$" & col_e125_Name & "$" & rowmax & ",0)"
        Cells(2, col_e125_gr_Name).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.TextToColumns Destination:=Cells(2, col_e125_gr_Name), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

'创建科目数组
    subject_array = Array("语文", "数学", "英语加权", "物理", "化学", "生物", "历史", "地理", "政治")
    subject_col_array = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    Dim col_istrue As Range
    For i = 0 To 8
        Set col_istrue = Rows("1:1").Find(What:=subject_array(i))
        If Not col_istrue Is Nothing Then
            subject_col_array(i) = Split(col_istrue.Address, "$")(1)
        End If
    Next
    
'清洗科目数组，过滤掉不存在的列名，创建纯列名数组
    col_isNotNull = Join(subject_col_array, ",")
    col_isNotNull = Replace(col_isNotNull, "0,", "")
    col_isNotNull = Replace(col_isNotNull, ",0", "")
    col_isNotNull = Replace(col_isNotNull, "0", "")
    col_isNotNull = Replace(col_isNotNull, ",,", ",")
    subject_col_array = Split(col_isNotNull, ",")

'计算加权总分及排名
    For h = 2 To rowmax
        a = 0
        For Each i In subject_col_array
            a = a + Cells(h, i).Value
        Next
        Range(col_a125_Name & h) = a
    Next
    For m = 2 To rowmax
        Sheets("成绩排名").Range(col_a125_cr_Name & m).Formula = "=SUMPRODUCT((" & col_c_Name & ":" & col_c_Name & "=" & col_c_Name & m & ")*(" & col_a125_Name & ":" & col_a125_Name & ">" & col_a125_Name & m & "))+1"
            Range(col_a125_cr_Name & m).Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    Next
    Sheets("成绩排名").Range(col_a125_gr_Name & "2:" & col_a125_gr_Name & rowmax).Formula = "=RANK(" & col_a125_Name & "2,$" & col_a125_Name & "$2:$" & col_a125_Name & "$" & rowmax & ",1)"
        Cells(2, col_a125_gr_Name).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.TextToColumns Destination:=Cells(2, col_a125_gr_Name), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True

'最大列及最大+1列列名
    colmax = ActiveSheet.UsedRange.Columns.Count
    colmaxN = col_array(colmax - 1)
    colmaxN1 = col_array(colmax + 1 - 1)

'成绩排名工作表添加班级辅助列并去重
    Range(colmaxN1 & "2:" & colmaxN1 & rowmax).Value = Range("C2:C" & rowmax).Value
    Range(colmaxN1 & "2:" & colmaxN1 & rowmax).RemoveDuplicates 1
    cn = Application.CountA(Range(colmaxN1 & ":" & colmaxN1))

'逐个取辅助列的值，命名新建sheet，复制内容
    For k = 2 To cn
        clName = Sheets("成绩排名").Range(colmaxN1 & k).Value
        sheetNum = Sheets.Count
        Sheets.Add after:=Sheets(sheetNum)
        Sheets(sheetNum + 1).Name = clName
        Sheets(sheetNum + 1).Range("A1:" & colmaxN & rowmax).Value = Sheets("成绩排名").Range("A1:" & colmaxN & rowmax).Value
    Next
 
'删除辅助列
    Sheets("成绩排名").Select
    Columns(colmaxN1 & ":" & colmaxN1).Select
    Selection.Delete Shift:=xlToLeft
    
'逐个调整行高列宽，设置标题字体样式
    For s = 1 To sheetNum
        Sheets(s).Select
        Columns(col_c_Name & ":" & col_c_Name).ColumnWidth = 14
        Rows("1:" & rowmax).EntireRow.AutoFit
        Range("A1:" & col_array(colmax - 1) & "1").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = True
        '删除非本班学生成绩（这有问题）
            For y = 2 To rowmax
                If Cells(y, 3).Value <> ActiveSheet.Name Then
                    Rows(y & ":" & y).Select
                    Selection.Delete Shift:=xlUp
                End If
            Next
    Next

'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
ActiveWindow.WindowState = xlMaximized
ActiveWorkbook.Save
MsgBox "计算完成，用时" & Format(using_time, "0.0秒")

End Sub

