Attribute VB_Name = "Aggiungi_azienda"
Sub Add_Azienda()

    Application.ScreenUpdating = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Listino prezzi")
  
    Dim lastRow As Object
    Set lastRow = sh.CustomProperties.Item(1)
    
    Dim businessName As String
    
    '===============================================================================
    
    businessName = InputBox("Inserisci il nome dell'Azienda:", "Nome Azienda")
    
    If StrPtr(businessName) = 0 Then
        Exit Sub
    End If
    
    With sh.Range(sh.Cells(lastRow.Value + 1, "A"), sh.Cells(lastRow.Value + 2, "P"))
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
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With
    
    With sh.Range(sh.Cells(lastRow.Value + 1, "A"), sh.Cells(lastRow.Value + 2, "H"))
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Value = businessName
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Size = 18
        End With
    End With
    
    With sh.Range(sh.Cells(lastRow.Value + 1, "K"), sh.Cells(lastRow.Value + 2, "L"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,##0.00 $"
        .Value = 0
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Size = 14
        End With
    End With
     
    With sh.Range(sh.Cells(lastRow.Value + 1, "O"), sh.Cells(lastRow.Value + 2, "P"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,##0.00 $"
        .Value = 0
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Size = 14
        End With
    End With
    
    lastRow.Value = lastRow.Value + 2
    Debug.Print lastRow.Value
    
    
    formulaString = ""
    For i = 11 To lastRow.Value
        If Cells(i, "K").Font.Size = 14 Then
            formulaString = formulaString & Cells(i, "K").Address & ","
            i = i + 1
        End If
    Next
    If formulaString <> "" Then
        sh.Range("A7").Formula = "=SUM(" & formulaString & ")"
    Else
        sh.Range("A7").Value = 0
    End If
    
    
    formulaString = ""
    For i = 11 To lastRow.Value
        If Cells(i, "O").Font.Size = 14 Then
            formulaString = formulaString & Cells(i, "O").Address & ","
            i = i + 1
        End If
    Next
    If formulaString <> "" Then
        sh.Range("G7").Formula = "=SUM(" & formulaString & ")"
    Else
        sh.Range("G7").Value = 0
    End If
    
    
    
    
    
    
    
    
End Sub
