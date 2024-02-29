Public pptApp As PowerPoint.Application
Public fso
Public logDoc
Public startTime, endTime, totalCount, currentCount, speed, durTime, successCount, failCount, successRate, failRate
Public pres

' 获取当前PPT文件路径
Sub ModifyPPTFiles()
    Dim objFSO
    Dim strCurrentFolderPath

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set pptApp = Application ' PowerPoint.Application
    strCurrentFolderPath = objFSO.GetParentFolderName(pptApp.ActivePresentation.FullName) & "\"

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
        Dim folderPath
        Dim scriptPath
        Dim logFilePath
        Dim wordApp

        folderPath = folder.Self.path
        scriptPath = objFSO.GetParentFolderName(pptApp.ActivePresentation.FullName)
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
        AdjustMargins logDoc

        ' 预先计算待处理文件总数
        CalculateTotalCount folderPath

        ' 记录开始时间
        startTime = Now

        ' 遍历选定文件夹并处理文件
        ProcessFolder folderPath

        ' 记录结束时间
        endTime = Now

        ' 计算相关信息
        durTime = DateDiff("s", startTime, endTime) '持续时间
        If totalCount > 0 Then
            speed = durTime / totalCount '处理速度（每个文件多少秒）
        Else
            ' 如果 totalCount 为零，设置一个默认值或者采取其他措施
            speed = 0
        End If
        
        ' 计算成功率
        If totalCount > 0 Then
            successRate = successCount / totalCount
        Else
            ' 如果 totalCount 为零，设置一个默认值或者采取其他措施
            successRate = 0
        End If

        ' 计算失败率
        If totalCount > 0 Then
            failRate = failCount / totalCount
        Else
            ' 如果 totalCount 为零，设置一个默认值或者采取其他措施
            failRate = 0
        End If


        ' 结束时打印相关信息在日志文件顶部
        With logDoc.Content
            .InsertBefore "failCount: " & failCount & ", " & FormatPercent(failRate, 1) & vbCrLf & vbCrLf
            .InsertBefore "successCount: " & successCount & ", " & FormatPercent(successRate, 1) & vbCrLf
            .InsertBefore "speed: " & FormatNumber(speed, 2) & " s/file" & vbCrLf
            .InsertBefore "endTime: " & FormatDateTime(endTime, vbLongTime) & vbCrLf
            .InsertBefore "startTime: " & FormatDateTime(startTime, vbLongTime) & vbCrLf
        End With

        '将日志文件保存并提示结束
        logDoc.SaveAs2 logFilePath
        MsgBox "Done, pls check ProgressLog.docx"

        ' 清理对象
        Set logDoc = Nothing
        wordApp.Quit
        Set wordApp = Nothing
        Set fso = Nothing
        Set folder = Nothing
        Set shell = Nothing
    Else
        ' 取消选择文件夹时提示
        MsgBox "no folder picked"
    End If
End Sub

' 遍历文件夹和子文件夹以计算总文件数的函数
Sub CalculateTotalCount(path)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(path)

    ' 遍历所有文件
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "ppt" Or LCase(fso.GetExtensionName(file.Name)) = "pptx" Then
            totalCount = totalCount + 1
        End If
    Next

    ' 遍历子文件夹
    For Each subFolder In folder.SubFolders
        CalculateTotalCount subFolder.path
    Next
End Sub

' 遍历文件夹和子文件夹的递归函数
Sub ProcessFolder(path)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(path)

' 遍历所有文件
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "ppt" Or LCase(fso.GetExtensionName(file.Name)) = "pptx" Then
        currentCount = currentCount + 1 ' 更新当前处理的文件编号
        ' 尝试打开并处理 PowerPoint 文档
        'On Error Resume Next
        Set pres = pptApp.Presentations.Open(file.path, False, False, False)
        ' 确认 pres 对象是否被成功创建
        Debug.Print "Opened presentation at " & file.path
        If Err.Number = 0 Then
            ResetSlidesAndClearAllDesigns1 pres
            pres.Save
            ' 确认文件是否被保存
            Debug.Print "Saved presentation at " & file.path
            pres.Close
            ' 确认 pres 对象是否被成功关闭
            Debug.Print "Closed presentation at " & file.path
            LogProgress file.path, True, ""
            successCount = successCount + 1
        Else
            pres.Close
            ' 记录无法处理的文件信息
            LogProgress file.path, False, "Error " & Err.Number & ": " & Err.Description
            Err.Clear
            failCount = failCount + 1
        End If
        On Error GoTo 0
    End If
Next
' 遍历子文件夹
For Each subFolder In folder.SubFolders
    ProcessFolder subFolder.path ' 递归调用以处理子文件夹
Next
End Sub
' 清理 PowerPoint 文档中的设计和内容的函数
Sub ResetSlidesAndClearAllDesigns1(pres)
Dim pptDesign
Dim pptSlideMaster
Dim pptSlideLayout
Dim i
Dim j

' 确认 ResetSlidesAndClearAllDesigns1 是否被调用
Debug.Print "ResetSlidesAndClearAllDesigns1 started"

' 删除除第一个设计之外的所有设计
For i = pres.Designs.Count To 2 Step -1
    pres.Designs(i).Delete
Next i

' 遍历所有设计
For Each pptDesign In pres.Designs
    ' 获取当前设计的母版
    Set pptSlideMaster = pptDesign.SlideMaster
    ' 清除母版上的所有形状和图片
    For i = pptSlideMaster.Shapes.Count To 1 Step -1
        pptSlideMaster.Shapes(i).Delete
    Next i
    ' 遍历母版下的所有版式并清除它们上的形状
    For Each pptSlideLayout In pptSlideMaster.CustomLayouts
        ' 清除版式上的所有形状和图片
        For i = pptSlideLayout.Shapes.Count To 1 Step -1
            pptSlideLayout.Shapes(i).Delete
        Next i
    Next pptSlideLayout
Next pptDesign

' 遍历演示文稿中的每个幻灯片并重置
For Each pptSlide In pres.Slides
    pptSlide.CustomLayout = pres.Designs(1).SlideMaster.CustomLayouts(1)
Next pptSlide

' 确认 ResetSlidesAndClearAllDesigns1 是否完成
Debug.Print "ResetSlidesAndClearAllDesigns1 finished"
End Sub

'记录处理进度和错误到 Word 文档的函数
Sub LogProgress(filePath, isSuccess, errorMsg)
    Dim fileName, status
    Dim rate
    fileName = fso.GetFileName(filePath)
    status = IIf(isSuccess, "success", "fail: " & errorMsg)
    ' 计算进度百分比
    If totalCount > 0 Then
        rate = FormatPercent(currentCount / totalCount, 1)
    Else
        rate = "0%"
    End If
    ' 向 Word 文档顶部追加进度信息
    With logDoc.Content
        .InsertBefore " (" & currentCount & "/" & totalCount & ", " & rate & "), (" & status & "), " & filePath & vbCrLf
    End With
End Sub

'提供类似三元运算符的功能的函数
Function IIf(condition, truePart, falsePart)
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function

'调整页边距的函数
Sub AdjustMargins(doc)
    With doc.PageSetup
        .TopMargin = 14 ' 上边距，1厘米大约28英寸
        .BottomMargin = 14 ' 下边距，1厘米大约28英寸
        .LeftMargin = 14 ' 左边距，1厘米大约28英寸
        .RightMargin = 14 ' 右边距，1厘米大约28英寸
        .Gutter = 0 ' 装订线，1厘米大约28英寸
        .HeaderDistance = 0 ' 页眉距离，1厘米大约28英寸
        .FooterDistance = 0 ' 页脚距离，1厘米大约28英寸
        .Orientation = ppOrientationPortrait ' 纸张方向设置为纵向
    End With
End Sub

