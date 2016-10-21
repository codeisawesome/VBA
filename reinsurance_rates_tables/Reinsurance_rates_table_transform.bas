Attribute VB_Name = "Module1"
Sub GenerateTableForSelectX(SelectYears As Integer)

    Dim input_sheet As Worksheet
    Dim rw As Range
    Dim RowCount As Integer
    Dim table_count As Integer
    Dim WS As Worksheet
    Set WS = Sheets.Add
    
    WS.Name = "Select" & " " & SelectYears
    
    Set input_sheet = Worksheets("input_sheet")
    
    table_count = 0
    
    
    For Each rw In input_sheet.Rows
    
        Increment = table_count * (SelectYears + 2)
        
        If input_sheet.Cells(rw.Row, 1).Value = "" Then
            Exit For
        End If
        
        'Check to see if the duration is 1
        If input_sheet.Cells(rw.Row, 1).Value = 1 Then
        
            'puts the Id number
            WS.Range(WS.Cells(2 + Increment, 4), WS.Cells(SelectYears + 3 + Increment, 4)) = input_sheet.Cells(rw.Row, 7).Value
            
            'puts the name of the pointer
            WS.Range(WS.Cells(2 + Increment, 5), WS.Cells(SelectYears + 3 + Increment, 5)) = input_sheet.Cells(rw.Row, 6).Value
            
            'puts the durations
            For i = 1 To SelectYears + 2
                'input_sheet the -2 value
                If i = 1 Then
                    WS.Cells(2 + Increment, 6).Value = -2
                Else
                    'input_sheets 1 to 21 for durations
                    WS.Cells(i + 1 + Increment, 6) = i - 1
                End If
            Next i
            
            'puts the zeros everywhere
            WS.Range(WS.Cells(2 + Increment, 8), WS.Cells(SelectYears + 3 + Increment, 111)) = 0
            
            'puts the ages
            For i = 1 To 100
                WS.Cells(2 + Increment, i + 7) = i - 1
            Next i
            
            'puts pct value for duration 1
            WS.Range(WS.Cells(3 + Increment, input_sheet.Cells(rw.Row, 3).Value + 8), WS.Cells(3 + Increment, input_sheet.Cells(rw.Row, 4) + 8)) = input_sheet.Cells(rw.Row, 5).Value
            
        'puts the pct value for 21 to 99
        ElseIf input_sheet.Cells(rw.Row, 2).Value = 99 And input_sheet.Cells(rw.Row + 1, 2).Value = 99 Then
            WS.Range(WS.Cells(input_sheet.Cells(rw.Row, 1).Value + 2 + Increment, input_sheet.Cells(rw.Row, 3).Value + 8), WS.Cells(input_sheet.Cells(rw.Row, 1).Value + 2 + Increment, input_sheet.Cells(rw.Row, 4).Value + 8)) = input_sheet.Cells(rw.Row, 5).Value * 100
        
        'puts the pct value for duration 21+ if next row is not a new pointer
        ElseIf input_sheet.Cells(rw.Row, 2).Value = 99 And (input_sheet.Cells(rw.Row + 1, 2).Value = 1 Or input_sheet.Cells(rw.Row + 1, 2).Value = "") Then
            WS.Range(WS.Cells(input_sheet.Cells(rw.Row, 1).Value + 2 + Increment, input_sheet.Cells(rw.Row, 3).Value + 8), WS.Cells(input_sheet.Cells(rw.Row, 1).Value + 2 + Increment, input_sheet.Cells(rw.Row, 4).Value + 8)) = input_sheet.Cells(rw.Row, 5).Value * 100
        
        'increase table_count for next talbe
            table_count = table_count + 1
        'puts the pct value for duration and age
        Else
            WS.Range(WS.Cells(input_sheet.Cells(rw.Row, 1).Value + 2 + Increment, input_sheet.Cells(rw.Row, 3).Value + 8), WS.Cells(input_sheet.Cells(rw.Row, 2).Value + 2 + Increment, input_sheet.Cells(rw.Row, 4).Value + 8)) = input_sheet.Cells(rw.Row, 5).Value * 100
        End If
            
    Next rw
    
    'clear the formula column
    WS.Range(WS.Cells(2, 108), WS.Cells(7000, 108)).Clear
    
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
Sub SelectYears15()
    GenerateTableForSelectX 15
End Sub
Sub SelectYears10()
    GenerateTableForSelectX 10
End Sub
Sub SelectYears20()
    GenerateTableForSelectX 20
End Sub


