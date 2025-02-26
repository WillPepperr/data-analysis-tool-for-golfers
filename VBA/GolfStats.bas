Attribute VB_Name = "GolfStats"
Sub Update_hole_averages()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As Range
    Dim col As ListColumn
    Dim startCol As Integer
    Dim endCol As Integer
    Dim cell As Range
    
    Dim threes As Integer, fours As Integer, fives As Integer
    Dim threes_scores As Integer, fours_scores As Integer, fives_scores As Integer
    Dim threes_average As Variant, fours_average As Variant, fives_average As Variant
    
    Set ws = ThisWorkbook.Worksheets("Score Database")
    Set tbl = ws.ListObjects("scoreDatabase")

    If tbl.DataBodyRange Is Nothing Or tbl.ListRows.Count < 2 Then
        Worksheets("Backend").Cells(21, 6).Value = "NA"
        Worksheets("Backend").Cells(22, 6).Value = "NA"
        Worksheets("Backend").Cells(23, 6).Value = "NA"
        Exit Sub
    End If

    startCol = 21
    endCol = 38

    threes = 0: fours = 0: fives = 0
    threes_scores = 0: fours_scores = 0: fives_scores = 0
    
    Dim i As Integer
    For i = 2 To tbl.ListRows.Count
        Set row = tbl.ListRows(i).Range
        For Each col In tbl.ListColumns
            If col.Index >= startCol And col.Index <= endCol Then
                Set cell = row.Cells(col.Index)
                If Not IsEmpty(cell) Then
                    Select Case cell.Value
                        Case 3: threes = threes + 1: threes_scores = threes_scores + cell.Offset(0, -18).Value
                        Case 4: fours = fours + 1: fours_scores = fours_scores + cell.Offset(0, -18).Value
                        Case 5: fives = fives + 1: fives_scores = fives_scores + cell.Offset(0, -18).Value
                    End Select
                End If
            End If
        Next col
    Next i
    
    If threes > 0 Then threes_average = threes_scores / threes Else threes_average = "NA"
    If fours > 0 Then fours_average = fours_scores / fours Else fours_average = "NA"
    If fives > 0 Then fives_average = fives_scores / fives Else fives_average = "NA"
    
    Worksheets("Backend").Cells(21, 6).Value = threes_average
    Worksheets("Backend").Cells(22, 6).Value = fours_average
    Worksheets("Backend").Cells(23, 6).Value = fives_average
End Sub

Sub Update_greens_hit_percentage()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As Range
    Dim col As ListColumn
    Dim cell As Range
    
    Dim greens_hit As Integer, greens_checked As Integer
    Dim greens_percentage As Variant
    
    Set ws = ThisWorkbook.Worksheets("Score Database")
    Set tbl = ws.ListObjects("scoreDatabase")

    If tbl.DataBodyRange Is Nothing Or tbl.ListRows.Count < 2 Then
        Worksheets("Backend").Cells(9, 6).Value = "NA"
        Exit Sub
    End If
    
    Dim green_col_start As Integer: green_col_start = 57
    Dim green_col_end As Integer: green_col_end = 74

    greens_hit = 0
    greens_checked = 0

    Dim i As Integer
    For i = 2 To tbl.ListRows.Count
        Set row = tbl.ListRows(i).Range
        For Each col In tbl.ListColumns
            If col.Index >= green_col_start And col.Index <= green_col_end Then
                Set cell = row.Cells(col.Index)
                If Not IsEmpty(cell) Then
                    If cell.Value = 1 Then greens_hit = greens_hit + 1
                    If cell.Value = 1 Or cell.Value = 0 Then greens_checked = greens_checked + 1 ' Count both 1s and 0s
                End If
            End If
        Next col
    Next i

    If greens_checked > 0 Then
        greens_percentage = greens_hit / greens_checked
    Else
        greens_percentage = "NA"
    End If

    Worksheets("Backend").Cells(9, 6).Value = greens_percentage
End Sub


Sub Update_Fairways_hit_percentage()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim col As ListColumn
    Dim fairways_col_start As Integer
    Dim fairways_col_end As Integer
    Dim cell As Range
    
    Dim fairways_hit As Integer
    Dim fairways_checked As Integer
    Dim fairways_percentage As Double

    Set ws = ThisWorkbook.Worksheets("Score Database")
    Set tbl = ws.ListObjects("scoreDatabase")
    
    fairways_col_start = 39
    fairways_col_end = 56
    
    fairways_hit = 0
    fairways_checked = 0
    
    For Each row In tbl.ListRows
        For Each col In tbl.ListColumns
            If col.Index >= fairways_col_start And col.Index <= fairways_col_end Then
                Set cell = row.Range.Cells(col.Index)
                If Not IsEmpty(cell) Then
                    If cell.Value = 1 Then
                        fairways_hit = fairways_hit + 1
                        fairways_checked = fairways_checked + 1
                    End If
                    If cell.Value = 0 Then
                        fairways_checked = fairways_checked + 1
                    End If
                End If
            End If
        Next col
    Next row
    
    If fairways_checked > 0 Then
        fairways_percentage = fairways_hit / fairways_checked
    Else
        fairways_percentage = 0
    End If
    
    Worksheets("Backend").Cells(12, 6) = fairways_percentage
End Sub

Sub Update_putt_stats()
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As Range
    Dim col As ListColumn
    Dim putts_col_start As Integer
    Dim putts_col_end As Integer
    Dim cell As Range
    
    Dim one_putts As Integer
    Dim two_putts As Integer
    Dim three_putts_plus As Integer
    
    Dim total_putts As Integer
    Dim putt_average As Variant
    
    Dim putt_value As Integer
    Dim putt_instance_count As Integer
    
    Set ws = ThisWorkbook.Worksheets("Score Database")
    Set tbl = ws.ListObjects("scoreDatabase")
    
    If tbl.DataBodyRange Is Nothing Or tbl.ListRows.Count < 2 Then
        Worksheets("Backend").Cells(15, 6).Value = "NA"
        Worksheets("Backend").Cells(16, 6).Value = "NA"
        Worksheets("Backend").Cells(17, 6).Value = "NA"
        Worksheets("Backend").Cells(18, 6).Value = "NA"
        Exit Sub
    End If
    
    putts_col_start = 75
    putts_col_end = 92
    
    one_putts = 0
    two_putts = 0
    three_putts_plus = 0
    total_putts = 0
    putt_instance_count = 0
    
    Dim i As Integer
    For i = 2 To tbl.ListRows.Count
        Set row = tbl.ListRows(i).Range
        For Each col In tbl.ListColumns
            If col.Index >= putts_col_start And col.Index <= putts_col_end Then
                Set cell = row.Cells(col.Index)
                If Not IsEmpty(cell) Then
                    Select Case cell.Value
                        Case 1
                            one_putts = one_putts + 1
                            total_putts = total_putts + 1
                        Case 2
                            two_putts = two_putts + 1
                            total_putts = total_putts + 2
                        Case Is >= 3
                            putt_value = cell.Value
                            three_putts_plus = three_putts_plus + 1
                            total_putts = total_putts + putt_value
                    End Select
                    putt_instance_count = putt_instance_count + 1
                End If
            End If
        Next col
    Next i
    
    If putt_instance_count > 0 Then
        putt_average = total_putts / putt_instance_count
        Worksheets("Backend").Cells(15, 6).Value = putt_average
        Worksheets("Backend").Cells(16, 6).Value = one_putts / putt_instance_count
        Worksheets("Backend").Cells(17, 6).Value = two_putts / putt_instance_count
        Worksheets("Backend").Cells(18, 6).Value = three_putts_plus / putt_instance_count
    Else
        Worksheets("Backend").Cells(15, 6).Value = "NA"
        Worksheets("Backend").Cells(16, 6).Value = "NA"
        Worksheets("Backend").Cells(17, 6).Value = "NA"
        Worksheets("Backend").Cells(18, 6).Value = "NA"
    End If

End Sub

Sub Update_scores_summary()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim highest_score As Integer
    Dim lowest_score As Integer
    Dim average_score As Double
    Dim total_score As Integer
    Dim rounds_entered As Integer
    Dim cell As Range
    Dim i As Integer

    Dim lowest_round_date As Variant
    Dim highest_round_date As Variant
    Dim lowest_round_course As Variant
    Dim highest_round_course As Variant

    Dim scoreCol As Range, dateCol As Range, courseCol As Range
    Dim scoreIndex As Integer, dateIndex As Integer, courseIndex As Integer

    Set ws = ThisWorkbook.Worksheets("Score Database")
    Set tbl = ws.ListObjects("scoreDatabase")

    scoreIndex = 113
    dateIndex = 1
    courseIndex = 2
    Set scoreCol = tbl.ListColumns(scoreIndex).DataBodyRange
    Set dateCol = tbl.ListColumns(dateIndex).DataBodyRange
    Set courseCol = tbl.ListColumns(courseIndex).DataBodyRange

    highest_score = -100
    lowest_score = 10000
    total_score = 0
    rounds_entered = 0

    For i = 2 To scoreCol.Rows.Count
        Dim currentScore As Integer
        Dim currentDate As Variant
        Dim currentCourse As Variant

        currentScore = scoreCol.Cells(i, 1).Value
        currentDate = dateCol.Cells(i, 1).Value
        currentCourse = courseCol.Cells(i, 1).Value

        If currentScore < lowest_score Then
            lowest_score = currentScore
            lowest_round_date = currentDate
            lowest_round_course = currentCourse
        End If

        If currentScore > highest_score Then
            highest_score = currentScore
            highest_round_date = currentDate
            highest_round_course = currentCourse
        End If

        total_score = total_score + currentScore
        rounds_entered = rounds_entered + 1
    Next i

    If rounds_entered > 0 Then
        average_score = total_score / rounds_entered
    End If
    
    With Worksheets("Backend")
        .Cells(4, 5) = lowest_round_course
        .Cells(5, 5) = highest_round_course
        .Cells(6, 7) = average_score
        .Cells(4, 6) = lowest_round_date
        .Cells(5, 6) = highest_round_date
        .Cells(4, 7) = lowest_score
        .Cells(5, 7) = highest_score
    End With

End Sub


Sub Update_dashboard()
    Dim pt As PivotTable
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    Call Update_hole_averages
    Call Update_greens_hit_percentage
    Call Update_Fairways_hit_percentage
    Call Update_scores_summary
    Call Update_putt_stats
End Sub
