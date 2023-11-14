Attribute VB_Name = "Aggiungi_Button"

Sub Aggiungi()
    
    Application.ScreenUpdating = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Elenco Ditte")
    
    Dim newvalue As Integer
    
    Dim lastrow_ciane_input As Integer, lastrow_fornitori_input As Integer, lastrow As Integer
    lastrow_ciane_input = sh.CustomProperties.Item(1).Value
    lastrow_fornitori_input = sh.CustomProperties.Item(2).Value
    lastrow = IIf(Application.Caller = "Aggiungi_Uscita_Ciane13", lastrow_ciane_input, lastrow_fornitori_input)

    Dim itemSelected As Integer
    itemSelected = IIf(Application.Caller = "Aggiungi_Uscita_Ciane13", 1, 2)

    Dim leftLim_ciane As Integer, leftLim_fornitori As Integer, leftLim As Integer
    leftLim_ciane = 1
    leftLim_fornitori = 8
    leftLim = IIf(Application.Caller = "Aggiungi_Uscita_Ciane13", leftLim_ciane, leftLim_fornitori)
    
    Dim rightLim_ciane As Integer, rightLim_fornitori As Integer, rightLim As Integer
    rightLim_ciane = 5
    rightLim_fornitori = 12
    rightLim = IIf(Application.Caller = "Aggiungi_Uscita_Ciane13", rightLim_ciane, rightLim_fornitori)
    
    ' ========================================================================================================
    
    ' Togli il bordo all'ultima cella
    With sh.Range(sh.Cells(lastrow, leftLim), sh.Cells(lastrow, rightLim)).Borders(xlEdgeBottom)
        .LineStyle = xlLineStyleNone
    End With
    
    ' Calcola valore ultimo numero
    If lastrow = 15 Then
        newvalue = 1
    Else
        newvalue = sh.Cells(lastrow, leftLim).Value + 1
    End If

    ' Celle del numero dell'elenco
    With sh.Cells(lastrow + 1, leftLim)
        .Value = newvalue
        .HorizontalAlignment = xlCenter
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = IIf(Application.Caller = "Aggiungi_Uscita_Ciane13", xlThemeColorAccent4, xlThemeColorAccent2)
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End With

    ' Celle del nome
    With sh.Range(sh.Cells(lastrow + 1, leftLim + 1), sh.Cells(lastrow + 1, rightLim))
        .Merge
        .Locked = False
        .HorizontalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
    End With

    sh.CustomProperties.Item(itemSelected).Value = lastrow + 1
    
'    Debug.Print sh.CustomProperties.Item(itemSelected).Value

End Sub
