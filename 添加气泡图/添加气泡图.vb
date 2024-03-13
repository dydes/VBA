Sub CreateBubbleChartWithName()
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim rowmax As Long
    Dim chartName As String

    Set ws = ActiveSheet
    rowmax = ws.UsedRange.Rows.Count '获取最大行数
    chartName = ws.Range("F2").Value '假设F2单元格含有图表的名称
    
    ' 如果存在同名图表，则删除
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = chartName Then
            ws.ChartObjects(i).Delete
            Exit For ' 找到并删除后退出循环
        End If
    Next i
    

    ' 创建一个气泡图
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .ChartType = xlBubble3DEffect ' 设置图表类型为3D气泡图
    End With
    
    ' 给图表命名，确保名称有效
    If chartName <> "" And Not IsError(chartName) Then
        chartObj.Name = chartName
    Else
        MsgBox "图表名称无效，请在F2单元格中输入有效名称。"
        Exit Sub
    End If
    
    ' 添加气泡系列
    With ws.ChartObjects(chartName).Chart
        For i = 1 To rowmax
            .SeriesCollection.NewSeries
            With .SeriesCollection(i)
                .Name = "=" & ws.Name & "!A" & (i + 1)  '添加系列名
                .XValues = "=" & ws.Name & "!B" & (i + 1)  '添加x轴值
                .Values = "=" & ws.Name & "!C" & (i + 1)  '添加y轴值
                .BubbleSizes = "=" & ws.Name & "!D" & (i + 1)  '添加气泡大小值
            End With
        Next i
    End With
    
    ' 配置图表格式和轴
    With ws.ChartObjects(chartName).Chart
        .Axes(xlValue).MajorUnit = 1 ' 设置y轴的主要单位
        
        '图表添加标题
        If Not .HasTitle Then
            .SetElement msoElementChartTitleAboveChart
        End If
        .ChartTitle.Text = chartName ' 设置图表标题
    End With
    
End Sub