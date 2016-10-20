Attribute VB_Name = "Module1"
Sub GenerateTableForSelectX(SelectYears As Integer)

    Dim sh As Worksheet
    Dim rw As Range
    Dim RowCount As Integer
    Dim table_count As Integer
    
    Set sh = Sheet4
    
    table_count = 0
    
    
    For Each rw In sh.Rows
    
        Increment = table_count * (SelectYears + 2)
        
        If sh.Cells(rw.Row, 1).Value = "" Then
            Exit For
        End If
        
        'Check to see if the duration is 1
        If Sheet4.Cells(rw.Row, 1).Value = 1 Then
        
            'puts the Id number
            Sheet1.Range(Sheet1.Cells(2 + Increment, 4), Sheet1.Cells(SelectYears + 3 + Increment, 4)) = Sheet4.Cells(rw.Row, 7).Value
            
            'puts the name of the pointer
            Sheet1.Range(Sheet1.Cells(2 + Increment, 5), Sheet1.Cells(SelectYears + 3 + Increment, 5)) = Sheet4.Cells(rw.Row, 6).Value
            
            'puts the durations
            For i = 1 To SelectYears + 2
                'input the -2 value
                If i = 1 Then
                    Sheet1.Cells(2 + Increment, 6).Value = -2
                Else
                    'inputs 1 to 21 for durations
                    Sheet1.Cells(i + 1 + Increment, 6) = i - 1
                End If
            Next i
            
            'puts the zeros everywhere
            Sheet1.Range(Sheet1.Cells(2 + Increment, 8), Sheet1.Cells(SelectYears + 3 + Increment, 111)) = 0
            
            'puts the ages
            For i = 1 To 100
                Sheet1.Cells(2 + Increment, i + 7) = i - 1
            Next i
            
            'puts pct value for duration 1
            Sheet1.Range(Sheet1.Cells(3 + Increment, Sheet4.Cells(rw.Row, 3).Value + 8), Sheet1.Cells(3 + Increment, Sheet4.Cells(rw.Row, 4) + 8)) = Sheet4.Cells(rw.Row, 5).Value
            
        'puts the pct value for 21 to 99
        ElseIf Sheet4.Cells(rw.Row, 2).Value = 99 And Sheet4.Cells(rw.Row + 1, 2).Value = 99 Then
            Sheet1.Range(Sheet1.Cells(Sheet4.Cells(rw.Row, 1).Value + 2 + Increment, Sheet4.Cells(rw.Row, 3).Value + 8), Sheet1.Cells(Sheet4.Cells(rw.Row, 1).Value + 2 + Increment, Sheet4.Cells(rw.Row, 4).Value + 8)) = Sheet4.Cells(rw.Row, 5).Value * 100
        
        'puts the pct value for duration 21+ if next row is not a new pointer
        ElseIf Sheet4.Cells(rw.Row, 2).Value = 99 And (Sheet4.Cells(rw.Row + 1, 2).Value = 1 Or Sheet4.Cells(rw.Row + 1, 2).Value = "") Then
            Sheet1.Range(Sheet1.Cells(Sheet4.Cells(rw.Row, 1).Value + 2 + Increment, Sheet4.Cells(rw.Row, 3).Value + 8), Sheet1.Cells(Sheet4.Cells(rw.Row, 1).Value + 2 + Increment, Sheet4.Cells(rw.Row, 4).Value + 8)) = Sheet4.Cells(rw.Row, 5).Value * 100
        
        'increase table_count for next talbe
            table_count = table_count + 1
        'puts the pct value for duration and age
        Else
            Sheet1.Range(Sheet1.Cells(Sheet4.Cells(rw.Row, 1).Value + 2 + Increment, Sheet4.Cells(rw.Row, 3).Value + 8), Sheet1.Cells(Sheet4.Cells(rw.Row, 2).Value + 2 + Increment, Sheet4.Cells(rw.Row, 4).Value + 8)) = Sheet4.Cells(rw.Row, 5).Value * 100
        End If
            
    Next rw
    
    'clear the formula column
    Sheet1.Range(Sheet1.Cells(2, 108), Sheet1.Cells(7000, 108)).Clear
    
End Sub
'for 1 to 1 and x to 99
Sub InsertDuration1to1AndxTo99()

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        If IsEmpty(Cells(i, 1)) Then
                Range(Cells(i, 1), Cells(i + 1, 7)).Value = Range("A2:G3").Value
               Range(Cells(i, 6), Cells(i, 7)).Value = Range(Cells(i - 1, 6), Cells(i - 1, 7)).Value
               Range(Cells(i + 1, 6), Cells(i + 1, 7)).Value = Range(Cells(i + 2, 6), Cells(i + 2, 7)).Value
        End If
    Next i
    
End Sub
'for 1 to 1 pointers only
Sub InsertDuration1to1()

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        If IsEmpty(Cells(i, 1)) Then
               Range(Cells(i, 1), Cells(i, 7)).Value = Range("A1:G1").Value
               Range(Cells(i, 6), Cells(i, 7)).Value = Range(Cells(i + 1, 6), Cells(i + 1, 7)).Value
        End If
    Next i
    
End Sub
Sub Insert2BlankRows(top As Integer, chunkSize As Integer)

     'Select last row in worksheet.
    Selection.End(xlDown).Select
     
    Do Until ActiveCell.Row = top
         'Insert blank row.
        ActiveCell.EntireRow.Insert shift:=xlDown
        ActiveCell.EntireRow.Insert shift:=xlDown
         'Move up one row.
        ActiveCell.Offset(-chunkSize, 0).Select
    Loop

End Sub
Sub Insert1BlankRows(top As Integer, chunkSize As Integer)

     'Select last row in worksheet.
    Selection.End(xlDown).Select
     
    Do Until ActiveCell.Row = top
         'Insert blank row.
        ActiveCell.EntireRow.Insert shift:=xlDown
         'Move up one row.
        ActiveCell.Offset(-chunkSize, 0).Select
    Loop

End Sub
Sub Insert1BlankRows1546()
    Insert1BlankRows 1, 10
End Sub
Sub InsertDuration1to1AndxTo99_1546()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert2BlankRows1548()
    Insert2BlankRows 3, 3
End Sub
Sub InsertDuration1to1AndxTo99_1548()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert2BlankRows1551()
    Insert2BlankRows 5, 14
End Sub
Sub InsertDuration1to1AndxTo99_1551()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert1BlankRows1552()
    Insert1BlankRows 1, 12
End Sub
Sub InsertDuration1to1_1552()
    InsertDuration1to1
End Sub
Sub Insert2BlankRows1553()
    Insert2BlankRows 5, 9
End Sub
Sub InsertDuration1to1AndxTo99_1552()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert1BlankRows1554()
    Insert1BlankRows 1, 12
End Sub
Sub InsertDuration1to1_1554()
    InsertDuration1to1
End Sub
Sub Insert1BlankRows1555()
    Insert1BlankRows 1, 12
End Sub
Sub InsertDuration1to1_1555()
    InsertDuration1to1
End Sub
Sub Insert2BlankRows1556()
    Insert2BlankRows 5, 2
End Sub
Sub InsertDuration1to1AndxTo99_1556()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert2BlankRows1557()
    Insert2BlankRows 4, 6
End Sub
Sub InsertDuration1to1AndxTo99_1557()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert2BlankRows1558()
    Insert2BlankRows 3, 3
End Sub
Sub InsertDuration1to1AndxTo99_1558()
    InsertDuration1to1AndxTo99
End Sub
Sub Insert1BlankRows1559()
    Insert1BlankRows 1, 12
End Sub
Sub InsertDuration1to1_1559()
    InsertDuration1to1
End Sub
Sub SelectYears15()
    GenerateTableForSelectX 15
End Sub
Sub SelectYears10()
    GenerateTableForSelectX 10
End Sub
Sub SelectYears20()
    GenerateTableForSelectX 20
End Sub


