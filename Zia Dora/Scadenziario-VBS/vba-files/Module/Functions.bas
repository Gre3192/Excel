Attribute VB_Name = "Functions"

' ==============================================================================================================================================================
Function formattazione_tabella_ciane(ByRef sh_mese As Worksheet, ByVal stringValue As String, ByVal i As Integer)

    ' Formatta bordi uscita
    With sh_mese.Range(sh_mese.Cells(18 + i, "A"), sh_mese.Cells(18 + 5 + i, "M"))
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ' Formatta nome uscita
    With sh_mese.Range(sh_mese.Cells(18 + i, "A"), sh_mese.Cells(18 + 5 + i, "D"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = stringValue
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    ' Formatta celle input e validazione
    For j = 0 To 5
    
        ' Tipologia pagamento
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "E"), sh_mese.Cells(18 + i + j, "G"))
            .Merge
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
             End With
             With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-,Effetto Bancario,RID Bancario,Bonifico,Assegni,Contanti"
                .IgnoreBlank = True
                .ShowInput = True
                .ShowError = True
             End With
             .Value = "-"
             .HorizontalAlignment = xlCenter
             .Font.Color = RGB(48, 84, 150)
        End With
        
        ' Data scadenza
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "H"), sh_mese.Cells(18 + i + j, "I"))
            .Merge
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
             End With
            With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
                .IgnoreBlank = True
                .ShowInput = True
                .ShowError = True
             End With
             .Value = "-"
             .HorizontalAlignment = xlCenter
             .Font.Color = RGB(48, 84, 150)
        End With
        
        ' Importo
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "J"), sh_mese.Cells(18 + i + j, "K"))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Stato pagamento
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "L"), sh_mese.Cells(18 + i + j, "M"))
            .Merge
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
             End With
            With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-,Pagato"
                .IgnoreBlank = True
                .ShowInput = True
                .ShowError = True
             End With
             .Value = "-"
             .HorizontalAlignment = xlCenter
             .Font.Color = RGB(48, 84, 150)
        End With
        targetValue = "Pagato"
        Set fc = sh_mese.Range(sh_mese.Cells(18 + i + j, "L"), sh_mese.Cells(18 + i + j, "M")).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & targetValue & """")
        fc.Interior.Color = RGB(198, 239, 206)
        fc.Font.Color = RGB(0, 97, 0)
        
    Next

End Function
' ==============================================================================================================================================================
Function formattazione_tabella_fornitori(ByRef sh_mese As Worksheet, ByVal stringValue As String, ByVal i As Integer)

    ' Formatta bordi uscita
    With sh_mese.Range(sh_mese.Cells(18 + i, "O"), sh_mese.Cells(18 + 5 + i, "AA"))
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ' Formatta nome uscita
    With sh_mese.Range(sh_mese.Cells(18 + i, "O"), sh_mese.Cells(18 + 5 + i, "R"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = stringValue
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    ' Formatta celle input e validazione
    For j = 0 To 5
    
        ' Tipologia pagamento
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "S"), sh_mese.Cells(18 + i + j, "U"))
            .Merge
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
             End With
             With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-,Effetto Bancario,RID Bancario,Bonifico,Assegni,Contanti"
                .IgnoreBlank = True
                .ShowInput = True
                .ShowError = True
             End With
             .Value = "-"
             .HorizontalAlignment = xlCenter
             .Font.Color = RGB(48, 84, 150)
        End With
        
        ' Data scadenza
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "V"), sh_mese.Cells(18 + i + j, "W"))
            .Merge
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
             End With
            With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31"
                .IgnoreBlank = True
                .ShowInput = True
                .ShowError = True
             End With
             .Value = "-"
             .HorizontalAlignment = xlCenter
             .Font.Color = RGB(48, 84, 150)
        End With
        
        ' Importo
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "X"), sh_mese.Cells(18 + i + j, "Y"))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Stato pagamento
        With sh_mese.Range(sh_mese.Cells(18 + i + j, "Z"), sh_mese.Cells(18 + i + j, "AA"))
            .Merge
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
             End With
            With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-,Pagato"
                .IgnoreBlank = True
                .ShowInput = True
                .ShowError = True
             End With
             .Value = "-"
             .HorizontalAlignment = xlCenter
             .Font.Color = RGB(48, 84, 150)
        End With
        targetValue = "Pagato"
        Set fc = sh_mese.Range(sh_mese.Cells(18 + i + j, "Z"), sh_mese.Cells(18 + i + j, "AA")).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & targetValue & """")
        fc.Interior.Color = RGB(198, 239, 206)
        fc.Font.Color = RGB(0, 97, 0)
        
    Next

End Function
' ==============================================================================================================================================================
Function formattazione_ciane_riepilogo(ByRef sh_mese As Worksheet, ByRef sh_riepilogo As Worksheet, ByVal stringValue As String, ByVal i As Integer)

    sh_riepilogo.Cells(18 + i, "A").Value = stringValue
    With sh_riepilogo.Range(sh_riepilogo.Cells(18 + i, "A"), sh_riepilogo.Cells(18 + i, "D"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    
    For col = 1 To 24 Step 2
    
        
    
        With sh_riepilogo.Range(sh_riepilogo.Cells(18 + i, col + 4), sh_riepilogo.Cells(18 + i, col + 5))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
    '        .Formula = "=" & startCell & ":" & endCell
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End With
    Next
    
    
End Function
' ==============================================================================================================================================================
Function formattazione_fornitori_riepilogo(ByRef sh_mese As Worksheet, ByRef sh_riepilogo As Worksheet, ByVal stringValue As String, ByVal i As Integer)

    sh_riepilogo.Cells(18 + i, "AD").Value = stringValue
    With sh_riepilogo.Range(sh_riepilogo.Cells(18 + i, "AD"), sh_riepilogo.Cells(18 + i, "AG"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With
    
    
    For col = 1 To 24 Step 2
    
        With sh_riepilogo.Range(sh_riepilogo.Cells(18 + i, col + 33), sh_riepilogo.Cells(18 + i, col + 34))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
    '        .Formula = "=" & startCell & ":" & endCell
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End With
    Next

End Function
' ==============================================================================================================================================================
Function formule_ciane_riepilogo(ByRef sh_riepilogo As Worksheet, ByRef sh_mese As Worksheet, ByRef tot_uscite_elenco As Integer, ByRef index_mese_colonna As Integer)

    Dim j As Integer
    
    j = 1
    For i = 1 To tot_uscite_elenco
    
        startCell = sh_mese.Cells(j + 17, "J").Address
        endCell = sh_mese.Cells(j + 22, "K").Address
        
        With sh_riepilogo.Range(sh_riepilogo.Cells(17 + i, index_mese_colonna + 1), sh_riepilogo.Cells(17 + i, index_mese_colonna + 2))
            .Formula = "=SUM('" & sh_mese.Name & "'!" & startCell & ":'" & sh_mese.Name & "'!" & endCell & ")"
        End With
        j = j + 6

    Next

End Function
' ==============================================================================================================================================================
Function formule_fornitori_riepilogo(ByRef sh_riepilogo As Worksheet, ByRef sh_mese As Worksheet, ByRef tot_fornitori_elenco As Integer, ByRef index_mese_colonna As Integer)

    Dim j As Integer
    
    j = 1
    For i = 1 To tot_fornitori_elenco
    
        startCell = sh_mese.Cells(j + 17, "X").Address
        endCell = sh_mese.Cells(j + 22, "Y").Address
        
        With sh_riepilogo.Range(sh_riepilogo.Cells(17 + i, index_mese_colonna + 1), sh_riepilogo.Cells(17 + i, index_mese_colonna + 2))
            .Formula = "=SUM('" & sh_mese.Name & "'!" & startCell & ":'" & sh_mese.Name & "'!" & endCell & ")"
        End With
        j = j + 6

    Next

End Function
' ==============================================================================================================================================================
Function formule_totali_ciane_riepilogo(ByRef sh_riepilogo As Worksheet, Optional status As String = "No Delete")

    Dim lastrow_riepilogo_uscite As Integer
    lastrow_riepilogo_uscite = sh_riepilogo.Cells(Rows.Count, "A").End(xlUp).Row
    
    If status = "Delete" Then
        sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite - 2, "A"), sh_riepilogo.Cells(lastrow_riepilogo_uscite, "AB")).Delete Shift:=xlUp
        Exit Function
    End If
    
    ' Formula totale mensile tabella riepilogo ciane
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite + 1, "A"), sh_riepilogo.Cells(lastrow_riepilogo_uscite + 1, "D"))
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "TOTALE MENSILE"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
    End With
    
    For col = 1 To 24 Step 2
        startCell = sh_riepilogo.Cells(18, col + 4).Address
        endCell = sh_riepilogo.Cells(lastrow_riepilogo_uscite, col + 5).Address
        With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite + 1, col + 4), sh_riepilogo.Cells(lastrow_riepilogo_uscite + 1, col + 5))
            .Merge
            .Font.Bold = True
            .Formula = "=SUM(" & startCell & ":" & endCell & ")"
            .HorizontalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        End With
    Next
    
    ' Formula totale semestre tabella riepilogo ciane
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite + 2, "A"), sh_riepilogo.Cells(lastrow_riepilogo_uscite + 2, "D"))
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "TOTALE SEMESTRE"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With
    
    For col = 1 To 13 Step 12
        startCell = sh_riepilogo.Cells(lastrow_riepilogo_uscite + 1, col + 4).Address
        endCell = sh_riepilogo.Cells(lastrow_riepilogo_uscite + 1, col + 15).Address
        With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite + 2, col + 4), sh_riepilogo.Cells(lastrow_riepilogo_uscite + 2, col + 15))
            .Merge
            .Formula = "=SUM(" & startCell & ":" & endCell & ")"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    Next
           
    ' Formula totale anno tabella riepilogo ciane
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite + 3, "A"), sh_riepilogo.Cells(lastrow_riepilogo_uscite + 3, "D"))
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "TOTALE ANNO"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With
    
    startCell = sh_riepilogo.Cells(lastrow_riepilogo_uscite + 2, "E").Address
    endCell = sh_riepilogo.Cells(lastrow_riepilogo_uscite + 2, "AB").Address
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_uscite + 3, "E"), sh_riepilogo.Cells(lastrow_riepilogo_uscite + 3, "AB"))
        .Merge
        .Font.Bold = True
        .Formula = "=SUM(" & startCell & ":" & endCell & ")"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With

End Function
' ==============================================================================================================================================================
Function formule_totali_fornitori_riepilogo(ByRef sh_riepilogo As Worksheet, Optional status As String = "No Delete")

    Dim lastrow_riepilogo_fornitori As Integer
    lastrow_riepilogo_fornitori = sh_riepilogo.Cells(Rows.Count, "AD").End(xlUp).Row
    
    If status = "Delete" Then
        sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori - 2, "AD"), sh_riepilogo.Cells(lastrow_riepilogo_fornitori, "BE")).Delete Shift:=xlUp
        Exit Function
    End If
    
    ' Formula totale mensile tabella riepilogo ciane
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 1, "AD"), sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 1, "AG"))
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "TOTALE MENSILE"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With
    
    For col = 1 To 24 Step 2
        startCell = sh_riepilogo.Cells(18, col + 33).Address
        endCell = sh_riepilogo.Cells(lastrow_riepilogo_fornitori, col + 34).Address
        With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 1, col + 33), sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 1, col + 34))
            .Merge
            .Font.Bold = True
            .Formula = "=SUM(" & startCell & ":" & endCell & ")"
            .HorizontalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        End With
    Next
    
    ' Formula totale semestre tabella riepilogo ciane
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 2, "AD"), sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 2, "AG"))
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "TOTALE SEMESTRE"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With
    
    For col = 1 To 13 Step 12
        startCell = sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 1, col + 33).Address
        endCell = sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 1, col + 44).Address
        With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 2, col + 33), sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 2, col + 44))
            .Merge
            .Formula = "=SUM(" & startCell & ":" & endCell & ")"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    Next
            
    ' Formula totale anno tabella riepilogo ciane
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 3, "AD"), sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 3, "AG"))
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "TOTALE ANNO"
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With
    
    startCell = sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 2, "AH").Address
    endCell = sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 2, "BE").Address
    With sh_riepilogo.Range(sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 3, "AH"), sh_riepilogo.Cells(lastrow_riepilogo_fornitori + 3, "BE"))
        .Merge
        .Font.Bold = True
        .Formula = "=SUM(" & startCell & ":" & endCell & ")"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    End With

End Function












