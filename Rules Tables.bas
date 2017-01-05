Sub GenerateFormulaV2()
    Dim risk_class As Object
    Dim table_to_use As Object
    Dim Current_text
    Dim Additional_text
    
    Set risk_class = CreateObject("Scripting.Dictionary")
    Set table_to_use = CreateObject("Scripting.Dictionary")

    For i = 0 To 246
    
        For j = 1 To 12
            risk_class(j) = Cells(1 + j + i * 12, 2).Value
        Next j
        
        For k = 1 To 12
            table_to_use(k) = Cells(1 + k + i * 12, 3).Value
        Next k
        
        'First line of the formula
        Cells(2 + i * 12, 4).Value = "If RiskAttribute (" & Chr(34) & "Sex1" & Chr(34) & ", " & Chr(34) & Left(risk_class(1), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Smkr1" & Chr(34) & ", " & Chr(34) & Right(Left(risk_class(1), 2), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Unisex Risk Class" & Chr(34) & ", " & Chr(34) & Right(risk_class(1), 1) & Chr(34) & ") Then" & Chr(10) & Chr(10) _
                                     & "   UseTable (" & Chr(34) & table_to_use(1) & Chr(34) & ")" & Chr(10) & Chr(10)
        
        'Next 11 ElseIf of the formula
        For l = 1 To 11
            Current_text = Cells(2 + i * 12, 4).Value
            
            Additional_text = "ElseIf RiskAttribute (" & Chr(34) & "Sex1" & Chr(34) & ", " & Chr(34) & Left(risk_class(l + 1), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Smkr1" & Chr(34) & ", " & Chr(34) & Right(Left(risk_class(l + 1), 2), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Unisex Risk Class" & Chr(34) & ", " & Chr(34) & Right(risk_class(l + 1), 1) & Chr(34) & ") Then" & Chr(10) & Chr(10) _
                                     & "   UseTable (" & Chr(34) & table_to_use(l + 1) & Chr(34) & ")" & Chr(10) & Chr(10)
            
            Cells(2 + i * 12, 4).Value = Current_text & Additional_text
        Next l
        
        'For X and P
        For m = 1 To 8
            Current_text = Cells(2 + i * 12, 4).Value
            
            If m <= 4 Then
                Additional_text = "ElseIf RiskAttribute (" & Chr(34) & "Sex1" & Chr(34) & ", " & Chr(34) & Left(risk_class(m), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Smkr1" & Chr(34) & ", " & Chr(34) & Right(Left(risk_class(m), 2), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Unisex Risk Class" & Chr(34) & ", " & Chr(34) & "X" & Chr(34) & ") Then" & Chr(10) & Chr(10) _
                     & "   UseTable (" & Chr(34) & table_to_use(m) & Chr(34) & ")" & Chr(10) & Chr(10)
            Else
                Additional_text = "ElseIf RiskAttribute (" & Chr(34) & "Sex1" & Chr(34) & ", " & Chr(34) & Left(risk_class(m - 4), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Smkr1" & Chr(34) & ", " & Chr(34) & Right(Left(risk_class(m - 4), 2), 1) & Chr(34) & ") And RiskAttribute (" & Chr(34) & "Unisex Risk Class" & Chr(34) & ", " & Chr(34) & "P" & Chr(34) & ") Then" & Chr(10) & Chr(10) _
                     & "   UseTable (" & Chr(34) & table_to_use(m - 4) & Chr(34) & ")" & Chr(10) & Chr(10)
            End If
            
            Cells(2 + i * 12, 4).Value = Current_text & Additional_text
        Next m
        
        'Set up End If
        Current_text = Cells(2 + i * 12, 4).Value
        Additional_text = "End if"
        Cells(2 + i * 12, 4).Value = Current_text & Additional_text

        
    Next i
        
End Sub

