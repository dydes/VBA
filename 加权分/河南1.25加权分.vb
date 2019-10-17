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
        .Title = "请选择工资文件"
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
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

'文本转数值格式
    colmax = ActiveSheet.UsedRange.Columns.Count
    For i = 1 To 26
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
    Sheets("成绩排名").Range(col_e125_Name & "2:" & col_e125_Name & rowmax).Formula = "=ROUNDUP(IFERROR(" & col_e_Name & "2*1.25,0),1)"
    Sheets("成绩排名").Range(col_e125_gr_Name & "2:" & col_e125_gr_Name & rowmax).Formula = "=RANK(" & col_e125_Name & "2,$" & col_e125_Name & "$2:$" & col_e125_Name & "$" & rowmax & ",1)"

'创建科目数组
    subject_array = Array("语文", "数学", "英语加权", "物理", "化学", "生物", "历史", "地理", "政治")
    subject_col_array = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    For i = 0 To 8
        col_istrue = Rows("1:1").Find(What:=subject_array(i)).Column
        If col_istrue > 0 Then
        subject_col_array(i) = col_array(col_istrue - 1)'这一步得到了带列标的数组
        End If
    Next

'计算加权总分及排名
    
'完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
ActiveWindow.WindowState = xlMaximized
MsgBox "计算完成，用时" & Format(using_time, "0.0秒")

End Sub