Sub 新增学校()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        b = Range("B" & i)
        C = Range("C" & i)
        D = Range("D" & i)
        E = Range("E" & i)
        f = Range("F" & i)
        g = Range("G" & i)
        Range("H" & i).Value = "INSERT INTO `htobservation`.`school_info` (`id`, `province_id`, `city_id`, `district_id`, `name`, `full_name`, `address`, `school_level`, `school_type`, `is_show`, `is_deleted`) VALUES ('" & A & "', '" & b & "', '" & C & "', '" & D & "', '" & E & "', '" & f & "', '" & g & "', '3', '1', '1', '0');"
    Next
End Sub

Sub 学校关联科目()

'计算有多少个学校
    Sheets("1新增学校").Select
    rowmax1 = ActiveSheet.UsedRange.Rows.Count

'删除当前表格的内容
    Sheets("1学校关联科目").Select
    rowmax2 = ActiveSheet.UsedRange.Rows.Count
    If rowmax2 > 1 Then
        Range("A2:A" & rowmax2).Select
        Selection.Delete Shift:=xlUp
    Else: GoTo 1
    End If
    
'生成关联关系
1:  For i = 2 To rowmax1
        For H = 1 To 23
            Range("A" & H + 1 + 23 * (i - 2)).Value = "INSERT INTO `htobservation`.`school_to_subject` (`school_id`, `subject`) VALUES ('" & Sheets("1新增学校").Range("A" & i).Value & "', '" & H & "');"
        Next
    Next
End Sub

Sub 新增项目组()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        b = Range("B" & i)
        C = Range("C" & i)
        D = Range("D" & i)
        Range("E" & i).Value = "INSERT INTO `htobservation`.`project_team` (`id`, `project_name`, `project_type`, `parent_id`, `is_deleted`) VALUES ('" & A & "', '" & b & "', '" & C & "',  '" & D & "', '0');"
    Next
End Sub

Sub 新增用户()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        b = Range("B" & i)
        Range("C" & i).Value = "INSERT INTO `htobservation`.`user_info` (`user_name`, `password`, `phone_num`, `is_active`, `is_deleted`) VALUES ('" & A & "', 'e10adc3949ba59abbe56e057f20f883e', '" & b & "', '1', '0');"
    Next
End Sub

Sub 新增教材()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        b = Range("B" & i)
        C = Range("C" & i)
        Range("D" & i).Value = "INSERT INTO `htobservation`.`material_info` (`subject_code`, `name`, `short_name`, `is_deleted`) VALUES ('" & A & "', '" & b & "', '" & C & "', '0');"
    Next
End Sub

Sub 新增课本()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        b = Range("B" & i)
        C = Range("C" & i)
        D = Range("D" & i)
        E = Range("E" & i)
        Range("F" & i).Value = "INSERT INTO `htobservation`.`book_info` (`material_id`, `name`, `order_index`, `education_stage`, `is_deleted`) VALUES ('" & b & "', '" & C & "', '" & D & "', '" & E & "', '0');"
    Next
End Sub

Sub 学校关联教材()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        b = Range("B" & i)
        D = Range("D" & i)
        Range("E" & i).Value = "INSERT INTO `htobservation`.`material_school_relation` (`material_id`, `school_id`, `is_deleted`) VALUES ('" & b & "', '" & D & "', '0');"
    Next
End Sub

Sub 新增班级()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        b = Range("B" & i)
        C = Range("C" & i)
        D = Range("D" & i)
        E = Range("E" & i)
        f = Range("F" & i)
        g = Range("G" & i)
        Range("H" & i).Value = "INSERT INTO `htobservation`.`class_info` (`id`, `school_id`, `enter_year`, `major_id`, `name`, `education_stage`, `student_num`, `class_type`, `is_deleted`, `class_index`) VALUES ('" & A & "', '" & b & "', '" & C & "', '" & D & "', '" & E & "', '" & f & "', '50', '11', '0', '" & g & "');"
    Next
End Sub

Sub 新增教师()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        b = Range("B" & i)
        C = Range("C" & i)
        D = Range("D" & i)
        Range("E" & i).Value = "INSERT INTO `htobservation`.`teacher_info` (`school_id`, `name`, `phone_num`, `is_deleted`) VALUES ('" & b & "', '" & C & "', '" & D & "', '0');"
    Next
End Sub

Sub 用户关联项目组()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        C = Range("C" & i)
        E = Range("E" & i)
        Range("F" & i).Value = "INSERT INTO `htobservation`.`user_project_relation` (`user_id`, `project_id`, `member_type`, `invite_status`, `is_deleted`) VALUES ('" & A & "', '" & C & "', '" & E & "', '2', '0');"
    Next
End Sub

Sub 新增章节()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        A = Range("A" & i)
        b = Range("B" & i)
        D = Range("D" & i)
        f = Range("F" & i)
        g = Range("G" & i)
        H = Range("H" & i)
        Range("I" & i).Value = "INSERT INTO `htobservation`.`book_chapter_info` (`name`, `order_index`, `material_id`, `book_id`, `is_leaf`, `father_id`, `is_deleted`) VALUES ('" & A & "', '" & b & "', '" & D & "', '" & f & "', '" & g & "', '" & H & "', '0');"
    Next
End Sub

Sub 班级关联教师()

'计算最大行数
    rowmax = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To rowmax
        b = Range("B" & i)
        C = Range("C" & i)
        E = Range("E" & i)
        g = Range("G" & i)
        Range("H" & i).Value = "INSERT INTO `htobservation`.`teacher_class_relation` (`school_id`, `subject`, `class_id`, `teacher_id`, `is_deleted`) VALUES ('" & b & "', '" & C & "', '" & E & "', '" & g & "', '0');"
    Next
End Sub

Sub 合并1()

'删除当前表格的内容
    Sheets("1合并").Select
    rowmax = ActiveSheet.UsedRange.Rows.Count
    If rowmax > 1 Then
        Range("A2:A" & rowmax).Select
        Selection.Delete Shift:=xlUp
    Else: GoTo 1
    End If
    
'所有1的sheet把sql都复制过来
1:  sht_arr = Array("1新增学校", "1学校关联科目", "1新增项目组", "1新增用户", "1新增教材")
    col_arr = Array("H", "A", "E", "C", "D")
    
    rr = 2
    For i = 0 To 4
        rowmax1 = Sheets(sht_arr(i)).UsedRange.Rows.Count
        Sheets("1合并").Range("A" & rr & ":A" & rr + rowmax1 - 2).Value = Sheets(sht_arr(i)).Range(col_arr(i) & "2:" & col_arr(i) & rowmax1).Value
        rr = ActiveSheet.UsedRange.Rows.Count + 1
    Next
End Sub

Sub 合并2()

'删除当前表格的内容
    Sheets("2合并").Select
    rowmax = ActiveSheet.UsedRange.Rows.Count
    If rowmax > 1 Then
        Range("A2:A" & rowmax).Select
        Selection.Delete Shift:=xlUp
    Else: GoTo 1
    End If
    
'所有1的sheet把sql都复制过来
1:  sht_arr = Array("2新增课本", "2学校关联教材", "2新增班级", "2新增教师", "2用户关联项目组")
    col_arr = Array("F", "E", "H", "E", "F")
    
    rr = 2
    For i = 0 To 4
        rowmax1 = Sheets(sht_arr(i)).UsedRange.Rows.Count
        Sheets("2合并").Range("A" & rr & ":A" & rr + rowmax1 - 2).Value = Sheets(sht_arr(i)).Range(col_arr(i) & "2:" & col_arr(i) & rowmax1).Value
        rr = ActiveSheet.UsedRange.Rows.Count + 1
    Next

End Sub

Sub 合并3()

'删除当前表格的内容
    Sheets("3合并").Select
    rowmax = ActiveSheet.UsedRange.Rows.Count
    If rowmax > 1 Then
        Range("A2:A" & rowmax).Select
        Selection.Delete Shift:=xlUp
    Else: GoTo 1
    End If
    
'所有1的sheet把sql都复制过来
1:  sht_arr = Array("3新增章节", "3班级关联教师")
    col_arr = Array("I", "H")
    
    rr = 2
    For i = 0 To 1
        rowmax1 = Sheets(sht_arr(i)).UsedRange.Rows.Count
        Sheets("3合并").Range("A" & rr & ":A" & rr + rowmax1 - 2).Value = Sheets(sht_arr(i)).Range(col_arr(i) & "2:" & col_arr(i) & rowmax1).Value
        rr = ActiveSheet.UsedRange.Rows.Count + 1
    Next

End Sub
