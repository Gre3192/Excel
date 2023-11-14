Attribute VB_Name = "Cancella"
Sub Elimina()

    Application.ScreenUpdating = False

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Listino prezzi")
    
    Dim lastRow As Object
    Set lastRow = sh.CustomProperties.Item(1)
    
    Dim deletedRow As Integer, addOne As Integer

    Dim firstCell_row As Integer, firstCell_col As Integer, secondCell_col As Integer, secondCell_row As Integer
    firstCell_row = Selection.Cells(1, 1).Row
    firstCell_col = Selection.Cells(1, 1).Column
    secondCell_row = Selection.Cells(Selection.Cells.Count).Row
    secondCell_col = Selection.Cells(Selection.Cells.Count).Column

    ' ----------------------------------------------------------------------------------------

    If (11 <= firstCell_row And firstCell_row <= lastRow.Value And 11 <= secondCell_row And secondCell_row <= lastRow.Value And firstCell_col = 1 And secondCell_col = 16) Then
    
        deletedRow = Selection.Rows.Count
        Selection.Delete Shift:=xlUp
        lastRow.Value = lastRow.Value - deletedRow
        
        addOne = IIf(sh.Cells(lastRow.Value, "A").Font.Size = 18, 1, 0)

        With sh.Range(sh.Cells(lastRow.Value - addOne, "A"), sh.Cells(lastRow.Value, "P"))
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
        
        ricalcola_formule
        
        Debug.Print lastRow.Value
        
    End If
    
End Sub
