Attribute VB_Name = "Elimina_Button"

Sub Elimina()

    Application.ScreenUpdating = False

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Elenco Ditte")
    
    Dim selectionRow1 As Integer, selectionCol1 As Integer, selectionRow2 As Integer, selectioncol2 As Integer
    selectionRow1 = Selection.Cells(1, 1).Row
    selectionCol1 = Selection.Cells(1, 1).Column
    selectionRow2 = Selection.Cells(Selection.Cells.Count).Row
    selectioncol2 = Selection.Cells(Selection.Cells.Count).Column

    Dim itemSelected As Integer, newLastRow As Integer, rowsDeleted As Integer
    itemSelected = IIf(Application.Caller = "Elimina_Uscita_Ciane13", 1, 2)
       
    Dim lastrow_ciane_input As Integer, lastrow_fornitori_input As Integer, lastrow As Integer
    lastrow_ciane_input = sh.CustomProperties.Item(1).Value
    lastrow_fornitori_input = sh.CustomProperties.Item(2).Value
    lastrow = IIf(Application.Caller = "Elimina_Uscita_Ciane13", lastrow_ciane_input, lastrow_fornitori_input)
    
    Dim leftLim_ciane As Integer, leftLim_fornitori As Integer, leftLim As Integer
    leftLim_ciane = 1
    leftLim_fornitori = 8
    leftLim = IIf(Application.Caller = "Elimina_Uscita_Ciane13", leftLim_ciane, leftLim_fornitori)

    Dim rightLim_ciane As Integer, rightLim_fornitori As Integer, rightLim As Integer
    rightLim_ciane = 5
    rightLim_fornitori = 12
    rightLim = IIf(Application.Caller = "Elimina_Uscita_Ciane13", rightLim_ciane, rightLim_fornitori)
    
    ' ========================================================================================================
    
    If lastrow > 15 Then

        If (15 < selectionRow1 And selectionRow1 <= lastrow) And (15 < selectionRow2 And selectionRow2 <= lastrow) And selectionCol1 = leftLim And selectioncol2 = rightLim Then
            
            sh.Range(sh.Cells(selectionRow1, leftLim), sh.Cells(selectionRow2, rightLim)).Delete Shift:=xlUp
            rowsDeleted = selectionRow2 - selectionRow1 + 1
            newLastRow = lastrow - rowsDeleted

            For i = 1 To newLastRow - 15
                sh.Cells(15 + i, leftLim) = i
            Next
            
            Cells(newLastRow, leftLim).Select
            
        Else
            
            sh.Range(sh.Cells(lastrow, leftLim), sh.Cells(lastrow, rightLim)).Delete Shift:=xlUp
            rowsDeleted = 1
            newLastRow = lastrow - rowsDeleted
            
        End If

        With sh.Range(sh.Cells(newLastRow, leftLim), sh.Cells(newLastRow, rightLim))
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With

        sh.CustomProperties.Item(itemSelected).Value = newLastRow

    End If

    Debug.Print sh.CustomProperties.Item(itemSelected).Value
    
End Sub

