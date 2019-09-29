Sub 工资计算()

'删除现有的无用sheet（需要加判断）
    Sheets(array("每日统计";"工资第一步";"工资第二步")).Select
    ActiveWindow.SelectedSheets.Delete
    
'删除顶部2行
    Sheets("每日统计").Select
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp

'插入“计费方式”、“人名+计费方式”式2列
    Columns("B:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'插入“早来”列
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'插入“晚走”、“时间差”、“实际工时”3列
    Columns("N:P").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'合并单元格并填充标题
    Merge_list = Array("B", "C", "K", "N", "O", "P")
    Title = Array("计费方式", "key", "早来", "晚走", "时间差", "实际工时")
    For Each i In Merge_list
        Range(Cells(1, i), Cells(2, i)).Merge
        
        
    Next
'这需要用一个变量遍历两个数组，可以参考一下继东当时的表如何实现的，通过i先构造range，select选中，然后再通过数组下标取数赋值

    For Each h In Title
        Range(Cells(1, i), Cells(2, i)).Select
        ActiveCell.FormulaR1C1 = h
    Next
    

End Sub
