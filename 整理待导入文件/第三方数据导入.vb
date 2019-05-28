Sub 第三方数据导入()

'定义变量
    Dim rowmax As Integer, file As String, fPath As String, splfile As Variant, fName As String, shtName As String

'先选择文件，获取路径，若未选择任何文件，终止程序
With Application.FileDialog(msoFileDialogFilePicker)
    .title = "请选择模板文件"
    .InitialFileName = "c:\"
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
    .title = "请选择要保存的文件夹"
    .InitialFileName = "c:\"
    If .Show = -1 Then
        fPath = .SelectedItems(1)
    Else: Exit Sub
    End If
End With

'打开指定工作簿并获取工作表名
Workbooks.Open (file)
shtName=activesheet.name & ".xlsx"

'选择模板工作簿的第一个sheet，复制并新建文件，保存新文件，关闭源文件
Windows(fName).Activate
    Sheets(1).Select
    Sheets(1).Copy
ChDir fPath
    ActiveWorkbook.SaveAs Filename:=fPath & "\" & shtName, FileFormat:= _
    xlOpenXMLWorkbook, CreateBackup:=False
Windows(fName).Activate
    ActiveWorkbook.Close savechanges:=False

'增加班级姓名key列
rowmax = ActiveSheet.UsedRange.Rows.Count
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("B:B").Select
Selection.Copy
Columns("D:D").Select
Selection.Insert Shift:=xlToRight
Columns("D:D").Select
Selection.Replace What:="高二", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
Selection.Replace What:="班", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
Range("E1").Select
ActiveCell.FormulaR1C1 = "key"
Range("E2:E" & rowmax).Formula = "=C2&D2"
Columns("F:F").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("F1").Select
ActiveCell.FormulaR1C1 = "系统有学校没有"
ActiveWorkbook.Save


End sub