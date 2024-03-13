Sub CreateBubbleChartWithName()
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim rowmax As Long
    Dim chartName As String
    Dim crossAtValue As Double
    Dim xmin, xmax, xstep, ymin, ymax, ystep

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
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=500, Top:=50, Height:=250)
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
    With chartObj.Chart
        For i = 1 To rowmax
            .SeriesCollection.NewSeries
            With .SeriesCollection(i)
                .Name = "=" & ws.Name & "!A" & (i + 1)  '添加系列名
                .XValues = "=" & ws.Name & "!B" & (i + 1)  '添加x轴值
                .Values = "=" & ws.Name & "!C" & (i + 1)  '添加y轴值
                .BubbleSizes = "=" & ws.Name & "!D" & (i + 1)  '添加气泡大小值
    
                ' 应用数据标签
                .ApplyDataLabels
                .DataLabels.ShowValue = False
                .DataLabels.ShowSeriesName = True
                .DataLabels.Position = xlLabelPositionCenter  ' 设置数据标签的位置
                
                ' 设置数据标签字体为微软雅黑
                .DataLabels.Font.Name = "微软雅黑"
    
                ' 设置数据标签的字体颜色为白色
                With .DataLabels.Format.TextFrame2.TextRange.Font.Fill
                    .ForeColor.RGB = RGB(255, 255, 255)  ' 白色
                    .Transparency = 0  ' 确保颜色不透明
                End With
            End With
        Next i
    End With

    ' 配置图表格式和轴
    xmin = range("G5")
    xmax = range("H5")
    xstep = range("I5")
    ymin = range("G6")
    ymax = range("H6")
    ystep = range("I6")
    With chartObj.Chart
        .Axes(xlCategory).MinimumScale = xmin ' 设置x轴的最小值
        .Axes(xlCategory).MaximumScale = xmax ' 设置x轴的最值
        .Axes(xlCategory).MajorUnit = xstep ' 设置x轴的步长
        .Axes(xlValue).MinimumScale = ymin ' 设置y轴的最小值
        .Axes(xlValue).MaximumScale = ymax ' 设置y轴的最大值
        .Axes(xlValue).MajorUnit = ystep ' 设置y轴的步长

        '图表添加标题
        If Not .HasTitle Then
            .SetElement msoElementChartTitleAboveChart
        End If
        .ChartTitle.Text = chartName ' 设置图表标题
        .ChartTitle.Font.Name = "微软雅黑"
        
        ' 设置X轴标题
        With chartObj.Chart.Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "低  ←     " & Range("B2") & "     →  高"
            .AxisTitle.Font.Name = "微软雅黑"
        End With
        
        ' 设置Y轴标题
        With chartObj.Chart.Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "低  ←     " & Range("C2") & "     →  高"
            .AxisTitle.Font.Name = "微软雅黑"
        End With
        
        ' 设置轴交叉点
        crossAtCategory = ws.Range("G2").Value
        crossAtValue = ws.Range("H2").Value
        
        If Not IsError(crossAtCategory) And Not IsEmpty(crossAtCategory) Then
            .Axes(xlCategory).CrossesAt = crossAtCategory
            .Axes(xlCategory).TickLabelPosition = xlNone
        End If
        If Not IsError(crossAtValue) And Not IsEmpty(crossAtValue) Then
            .Axes(xlValue).CrossesAt = crossAtValue
            .Axes(xlValue).TickLabelPosition = xlNone
        End If
        

        ' 关闭网格线
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlValue).HasMinorGridlines = False
    
        ' 关闭图例
        .HasLegend = False

        ' 设置边框线
        With chartObj.Chart.PlotArea.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
        End With
    End With
    
    '改变样式
    With chartObj.Chart
        .ChartType = xlBubble3DEffect
    End With
    
End Sub

