Attribute VB_Name = "Aggiungi_prodotto"
Sub Add_Product()

    Application.ScreenUpdating = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Listino prezzi")
  
    Dim lastRow As Object
    Set lastRow = sh.CustomProperties.Item(1)
    
    Dim rowValueadd As Integer
    Dim formulaString As String
    
    
    Dim firstCell_row As Integer, firstCell_col As Integer, secondCell_col As Integer, secondCell_row As Integer
    firstCell_row = Selection.Cells(1, 1).Row
    firstCell_col = Selection.Cells(1, 1).Column
    secondCell_row = Selection.Cells(Selection.Cells.Count).Row
    secondCell_col = Selection.Cells(Selection.Cells.Count).Column
    
    '===============================================================================
    
    isAdded = False
    If firstCell_col = 1 And secondCell_col = 8 Then
        isAdded = True
        rowValueadd = secondCell_row + 1
        sh.Range(sh.Cells(rowValueadd, "A"), sh.Cells(rowValueadd, "P")).Insert Shift:=xlDown
    Else
        rowValueadd = lastRow.Value + 1
    End If
        
    With sh.Range(sh.Cells(rowValueadd, "A"), sh.Cells(rowValueadd, "P"))
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With
    
    
    With sh.Range(sh.Cells(rowValueadd, "A"), sh.Cells(rowValueadd, "C"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = xlNone
        With .Font
            .Name = "Calibri"
            .Size = 11
        End With
    End With
    
    With sh.Range(sh.Cells(rowValueadd, "D"), sh.Cells(rowValueadd, "F"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = xlNone
        With .Font
            .Name = "Calibri"
            .Size = 11
        End With
    End With
    
    With sh.Range(sh.Cells(rowValueadd, "G"), sh.Cells(rowValueadd, "H"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = xlNone
        .NumberFormat = "#,##0.00 $"
        With .Font
            .Name = "Calibri"
            .Size = 11
        End With
    End With
      
    With sh.Range(sh.Cells(rowValueadd, "I"), sh.Cells(rowValueadd, "J"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = xlNone
        With .Font
            .Name = "Calibri"
            .Size = 11
        End With
    End With
    
    With sh.Range(sh.Cells(rowValueadd, "K"), sh.Cells(rowValueadd, "L"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,##0.00 $"
        .Formula = "=" & sh.Cells(rowValueadd, "G").Address & "*" & sh.Cells(rowValueadd, "I").Address
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    With sh.Range(sh.Cells(rowValueadd, "M"), sh.Cells(rowValueadd, "N"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = xlNone
        .NumberFormat = "0%"
        With .Font
            .Name = "Calibri"
            .Size = 11
        End With
    End With
    
    With sh.Range(sh.Cells(rowValueadd, "O"), sh.Cells(rowValueadd, "P"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,##0.00 $"
        .Formula = "=" & sh.Cells(rowValueadd, "K").Address & "*(1-" & sh.Cells(rowValueadd, "M").Address & ")"
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    

    lastRow.Value = lastRow.Value + 1
'    Debug.Print lastRow.Value
    
    
    formulaString = ""
    If isAdded Then
        thisaddress = Cells(firstCell_row, "K").Address
        For i = secondCell_row + 1 To lastRow.Value
            If Cells(i, "K").Font.Size = 14 And formulaString <> "" Then
                If Not IsEmpty(thisaddress) Then
                    sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
                End If
                Exit For
            ElseIf Cells(i, "K").Font.Size = 11 Then
                formulaString = formulaString & Cells(i, "K").Address & ","
            End If
            
            If i = CLng(lastRow.Value) Then
                If Not IsEmpty(thisaddress) And formulaString <> "" Then
                    sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
                End If
            End If
        Next
        
    Else
    
        For i = lastRow.Value To 11 Step -1
        
            If Cells(i, "K").Font.Size = 14 And formulaString <> "" Then
                thisaddress = Cells(i - 1, "K").Address
                sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
                Exit For
            ElseIf Cells(i, "K").Font.Size = 11 Then
                formulaString = formulaString & Cells(i, "K").Address & ","
            End If

        Next
    
    End If
    
    
    

    
    formulaString = ""
    If isAdded Then
        thisaddress = Cells(firstCell_row, "O").Address
        For i = secondCell_row + 1 To lastRow.Value
            If sh.Cells(i, "O").Font.Size = 14 And formulaString <> "" Then
                If Not IsEmpty(thisaddress) Then
                    sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
                End If
                Exit For
            ElseIf Cells(i, "O").Font.Size = 11 Then
                formulaString = formulaString & Cells(i, "O").Address & ","
            End If
            
            If i = CLng(lastRow.Value) Then
                If Not IsEmpty(thisaddress) And formulaString <> "" Then
                    sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
                End If
            End If
        Next
        
    Else
    
        For i = lastRow.Value To 11 Step -1
        
            If sh.Cells(i, "O").Font.Size = 14 And formulaString <> "" Then
                thisaddress = Cells(i - 1, "O").Address
                If Not IsEmpty(thisaddress) Then
                    sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
                End If
                Exit For
            ElseIf Cells(i, "O").Font.Size = 11 Then
                formulaString = formulaString & Cells(i, "O").Address & ","
            End If

        Next
    
    End If
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    

    
    
    








    
    

    
    
    
End Sub

