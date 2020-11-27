Public ofile, ofName, cfile, cfName, 成绩排名_colmax, 成绩排名_rowmax
Sub 等级赋分()
    
    '读取相关参数

    
    Application.ScreenUpdating = False '暂停刷新
    Application.DisplayAlerts = False '暂停通知
    
    '调用选文件函数，输出ofile是带路径的文件名，ofName是不带路径的文件名
    Call file_open_name("请选择年级全科文件", "D:\会通\VBA\等级赋分\")
    
    '调用保存文件函数，，输出cfile是带路径的文件夹名，cfName是不带路径的文件夹名
    Call file_save_name("请选择要保存的文件夹", "D:\会通\VBA\等级赋分\")
    
    '打开文件
    Workbooks.Open (ofile)
    
    '获取当前时间
    dat = Format(Date, "yyyy年mm月dd日") '当前年月日
    tim = Format(Time, "hh时mm分ss秒") '当前时间
    tim1 = Timer
    
    '选择“成绩排名”工作表，复制并新建文件，保存新文件，关闭源文件
    Windows(ofName).Activate
    Sheets("成绩排名").Select
    Sheets("成绩排名").Copy
    new_file = cfile & "\等级赋分-" & dat & "-" & tim & "生成.xlsx"
    ChDir cfile
    ActiveWorkbook.SaveAs Filename:=new_file, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Windows(ofName).Activate
    
    '关闭源文件
    ActiveWorkbook.Close savechanges:=False

    '删除顶部多余行
    Sheets("成绩排名").Select
    top_rows = Application.Match("学生学号", Range("A1:A100"), 0) - 1
        If top_rows > 0 Then
            Rows("1:" & top_rows).Delete Shift:=xlUp
        ElseIf top_rows = 0 Then
            Range("A1").Select
        Else
            MsgBox "未找到学生学号字段，请确认文件是否正确"
        End If
        
    '求成绩排名工作表最大行数列数
    成绩排名_colmax = Sheets("成绩排名").UsedRange.Columns.Count
    成绩排名_rowmax = Sheets("成绩排名").UsedRange.Rows.Count
    
    '调用替换函数替换--
    Call replace("--", "")
    
    '逐列文本转数值
    For i = 4 To 成绩排名_colmax
        Range(Cells(2, i), Cells(成绩排名_rowmax, i)).TextToColumns FieldInfo:=Array(1, 1)
    Next
    
    '新建工作表
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Name = "参数"
    Sheets("参数").Move Before:=Sheets(1)
    '填写列
    Range("A1") = "科目"
    Range("B1") = "是否存在"
    arr_sub = Array("总分", "语文", "数学", "英语", "物理", "化学", "生物", "历史", "地理", "政治")
    Range("A2").Resize(10, 1) = Application.Transpose(arr_sub)
    arr_insert = Array("原始班次", "原始级次", "等级", "赋分", "赋分班次", "赋分级次")
    Range("C1").Resize(1, 6) = arr_insert
        
    '判断科目是否存在
    For i = 0 To UBound(arr_sub)
        Debug.Print "i=" & i
        Sheets("成绩排名").Select
        col = Application.Match(arr_sub(i), Range(Cells(1, 1), Cells(1, 成绩排名_colmax)), 0)
            If IsNumeric(col) = True Then
                Sheets("参数").Select
                Row = Application.Match(arr_sub(i), Range("A2:A11"), 0)
                Range("B" & Row + 1) = 1
            Else
                Sheets("参数").Select
                Row = Application.Match(arr_sub(i), Range("A2:A11"), 0)
                Range("B" & Row + 1) = 0
            End If
    Next
    
    '调用插列函数
    For i = 2 To 11 '参数这个sheet，循环A列的各个科目
        
        Sheets("参数").Select
        If Range("B" & i) = 1 Then '如果B列是1，那么该科目存在
            j = Range("A" & i) '取出对应的科目名称
            Sheets("成绩排名").Select
            Call insert_subcol(j, 6) '调用插列函数，在对应的列后面插入6列，等后面再把多余的删掉
            Cells(1, (Application.Match(j, Range(Cells(1, 1), Cells(1, Columns.Count)), 0) + 1)).Resize(1, 6) = _
            j & i '插完列之后要选中对应的列头，填充标题
            
    
    '完成时间
    tim2 = Timer
    using_time = tim2 - tim1
    
    ActiveWindow.WindowState = xlMaximized
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True '重启刷新
    MsgBox "计算完成，用时" & Format(using_time, "0.0秒")
    
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

Function file_save_name(til, ifilname) 'til是文件选择器标题，ifilname是默认保存路径
'选择要保存的文件路径，若未选择任何文件夹，终止程序，输出cfile是带路径的文件夹名，cfName是不带路径的文件夹名
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = til
        .InitialFileName = ifilname
        If .Show = -1 Then
            cfile = .SelectedItems(1)
            Else: Exit Function
        End If
    End With
'用斜杠分割文件路径，创建数组，选取数组最后一个元素做为不含路径的文件名
    splfile = Split(cfile, "\")
    cfName = splfile(UBound(splfile))
End Function

'rbef表示替换什么，rlat表示替换后是什么，数字字符均可，字符用双引号'
Function replace(rbef, rlat)
    Range("A1").CurrentRegion.replace What:=rbef, Replacement:=rlat, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Function

Function insert_subcol(a, b) '参数a是需要找的列标题，参数b是需要插入的列数，在a后面插入b列
    For i = 1 To b
        Columns(Application.Match(a, Range(Cells(1, 1), Cells(1, Columns.Count)), 0) + 1).Insert _
        Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
End Function


