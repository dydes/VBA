' 获取当前 VBS 文件路径
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolderPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\"
 
' 判断 ProgressLog.docx 文件是否存在并进行处理
If objFSO.FileExists(strCurrentFolderPath & "ProgressLog.docx") Then
    ' 删除 ProgressLog.docx 文件
    objFSO.DeleteFile strCurrentFolderPath & "ProgressLog.docx"
End If

' 弹出文件夹选择对话框，让用户选择一个文件夹
Dim shell, folder
Set shell = CreateObject("Shell.Application")
Set folder = shell.BrowseForFolder(0, "请选择一个文件夹", 0)

If Not folder Is Nothing Then
    ' 初始化变量
    Dim folderPath, fso, scriptPath, logFilePath, wordApp, logDoc, startTime, endTime, totalCount, currentCount, speed, durTime, successCount, failCount, successRate, failRate
    folderPath = folder.Self.Path
    scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    logFilePath = scriptPath & "\ProgressLog.docx"
    Set fso = CreateObject("Scripting.FileSystemObject")
    totalCount = 0
    currentCount = 0
    successCount = 0
    failCount = 0

    ' 创建 Word.Application 对象用于记录进度，保持可见
    Set wordApp = CreateObject("Word.Application")
    Set logDoc = wordApp.Documents.Add
    wordApp.Visible = True '记录用的文档前端可见
    logDoc.Content.Text = "Go..."
    AdjustMargins

    ' 创建一个新的 Word.Application 对象用于在后台处理文档
    Dim wordAppBackground
    Set wordAppBackground = CreateObject("Word.Application")
    wordAppBackground.Visible = False '被处理的文档全都不可见

    ' 预先计算待处理文件总数
    CalculateTotalCount folderPath

    ' 记录开始时间
    startTime = Now

    ' 遍历选定文件夹并处理文件
    ProcessFolder folderPath

    ' 记录结束时间
    endTime = Now

    ' 计算相关信息
    durTime = DateDiff("s", startTime, endTime)'持续时间
    speed = durTime/totalCount'处理速度（每个文件多少秒）
    successRate = successCount/totalCount
    failRate = failCount/totalCount

    ' 结束时打印相关信息在日志文件顶部
    With logDoc.Content
        .InsertBefore "failCount: " & failCount & ", " & formatPercent(failRate,1) & vbCrLf & vbCrLf
        .InsertBefore "successCount: " & successCount & ", " & formatPercent(successRate,1) & vbCrLf
        .InsertBefore "speed: " & FormatNumber(speed, 2) & " s/file" & vbCrLf
        .InsertBefore "endTime: " & FormatDateTime(endTime, vbLongTime) & vbCrLf
        .InsertBefore "startTime: " & FormatDateTime(startTime, vbLongTime) & vbCrLf
    End With

    '将日志文件保存并提示结束
    logDoc.SaveAs2(logFilePath)
    MsgBox "Done, pls check ProgressLog.docx"

    ' 清理对象
    wordAppBackground.Quit
    Set wordAppBackground = Nothing
    Set logDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing
    Set folder = Nothing
    Set shell = Nothing
Else
    ' 取消选择文件夹时提示
    MsgBox "no folder picked"
End If

' 遍历文件夹和子文件夹以计算总文件数的函数
Sub CalculateTotalCount(path)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(path)
    ' 遍历所有文件
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "doc" Or LCase(fso.GetExtensionName(file.Name)) = "docx" Then
            totalCount = totalCount + 1
        End If
    Next
    ' 遍历子文件夹
    For Each subFolder In folder.SubFolders
        CalculateTotalCount subFolder.Path
    Next
End Sub

' 遍历文件夹和子文件夹的递归函数
Sub ProcessFolder(path)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(path)
    ' 遍历所有文件
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "doc" Or LCase(fso.GetExtensionName(file.Name)) = "docx" Then
            currentCount = currentCount + 1' 更新当前处理的文件编号
            ' 尝试打开并处理 Word 文档
            On Error Resume Next
            Dim doc
            Set doc = wordAppBackground.Documents.Open(file.Path, False, False)'第三个参数如果是True，表示以只读方式打开，会无法静默保存
            If Err.Number = 0 Then
                ClearHeadersAndFooters doc
                doc.Close True  ' False关闭文档但不保存更改
                LogProgress file.Path, True, ""
                successCount = successCount + 1
            Else
                doc.Close True
                ' 记录无法处理的文件信息
                LogProgress file.Path, False, "Error " & Err.Number & ": " & Err.Description
                Err.Clear
                failCount = failCount + 1
            End If
            On Error GoTo 0
        End If
    Next
    ' 遍历子文件夹
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder.Path' 递归调用以处理子文件夹
    Next
End Sub

' 清理 Word 文档中的页眉和页脚函数
Sub ClearHeadersAndFooters(doc)
    If doc Is Nothing Then Exit Sub
    Dim section
    For Each section In doc.Sections
        Dim headerFooter
        ' 清除每个部分的所有页眉
        For Each headerFooter In section.Headers
            headerFooter.Range.Text = ""
        Next
        ' 清除每个部分的所有页脚
        For Each headerFooter In section.Footers
            headerFooter.Range.Text = ""
        Next
    Next
End Sub

' 记录处理进度和错误到 Word 文档的函数
Sub LogProgress(filePath, isSuccess, errorMsg)
    Dim fileName, status, rate
    fileName = fso.GetFileName(filePath)
    status = IIf(isSuccess, "success", "fail: " & errorMsg)
    rate = FormatPercent(currentCount/totalCount, 1)
    ' 向 Word 文档顶部追加进度信息
    With logDoc.Content
        .InsertBefore " (" & currentCount & "/" & totalCount & ", " & rate & "), (" & status & "), " & filePath & vbCrLf
    End With
End Sub

' 提供类似三元运算符的功能的函数
Function IIf(condition, truePart, falsePart)
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function

'调整页边距的函数
Sub AdjustMargins()
    'Dim logDoc As Document
    'Set logDoc = ActiveDocument
    With logDoc.PageSetup
        .TopMargin = 14 ' 上边距，1厘米大约28英寸
        .BottomMargin = 14 ' 下边距，1厘米大约28英寸
        .LeftMargin = 14 ' 左边距，1厘米大约28英寸
        .RightMargin = 14 ' 右边距，1厘米大约28英寸
        .Gutter = 0 ' 装订线，1厘米大约28英寸
        .HeaderDistance = 0 ' 页眉距离，1厘米大约28英寸
        .FooterDistance = 0 ' 页脚距离，1厘米大约28英寸
        .Orientation = wdOrientPortrait ' 纸张方向设置为纵向
    End With
End Sub
