Attribute VB_Name = "EliminaCarico"
'namespace=vba-files\Moduli

Sub elimina_carico()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = Application.ThisWorkbook.ActiveSheet
    
    Dim button_clicked As String
    button_clicked = Application.caller

    Dim tot As Variant
    Dim isDash As Boolean
    Dim incrCol As Integer
    Dim current_cell As Range
    Dim start_row As Long, start_col As Long

    start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
    start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column

    incrCol = IIf(button_clicked = "Elimina Qk", 2, 0)
    isDash = IIf(ws.Cells(start_row + 1, start_col) = "-", True, False)

    
    '-- COLONNA TOT --------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 1, start_col)
        tot = current_cell.Value
        If isDash Then
            Exit Sub
        ElseIf current_cell.Value = 1 Then
            current_cell.Value = "-"
        Else
            current_cell.Value = current_cell.Value - 1
        End If
    '
    '-- COLONNA N° ---------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 3 + tot, start_col)
        If tot = 1 Then
            current_cell.Value = "-"
            cells_style "N°", current_cell
        Else
            current_cell.Value = current_cell.Value - 1
            cells_style "Cancella", current_cell
        End If
    '
    '-- COLONNA DESCRIZIONE ------------------------------------------------------------------------------------
        Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 1), ws.Cells(start_row + 3 + tot, start_col + 3))
        If tot = 1 Then
            cells_style "Descrizione", current_cell
            current_cell.ClearContents
        Else
            cells_style "Cancella", current_cell
        End If
    '
    '-- COLONNA CORRELAZIONE -----------------------------------------------------------------------------------
        If button_clicked = "Elimina Qk" Then
            Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 4), ws.Cells(start_row + 3 + tot, start_col + 5))
            If tot = 1 Then
                current_cell.Value = "-"
                cells_style "Correlazione", current_cell
            Else
                cells_style "Cancella", current_cell
            End If
        End If
    '
    '-- COLONNA INPUT CARICO -----------------------------------------------------------------------------------
        Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 4 + incrCol), ws.Cells(start_row + 3 + tot, start_col + 5 + incrCol))
        cells_style "Cancella", current_cell
    '
    '-- COLONNA CONDIZIONE -------------------------------------------------------------------------------------
        Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 6 + incrCol), ws.Cells(start_row + 3 + tot, start_col + 7 + incrCol))
        If tot = 1 Then
            current_cell.Value = "-"
            cells_style "Condizione", current_cell
            current_cell.Validation.Delete
        Else
            cells_style "Cancella", current_cell
        End If
    '
    '-- COLONNA ANALISI ----------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 3 + tot, start_col + 8 + incrCol)
        If tot = 1 Then
            current_cell.Value = "-"
            cells_style "Analisi", current_cell
            current_cell.Validation.Delete
        Else
            cells_style "Cancella", current_cell
        End If
    '
    '-- COLONNA CATEGORIA --------------------------------------------------------------------------------------
        If button_clicked = "Elimina Qk" Then
            Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 11), ws.Cells(start_row + 3 + tot, start_col + 13))
            If tot = 1 Then
                current_cell.Value = "-"
                cells_style "Categoria", current_cell
                current_cell.Validation.Delete
            Else
                cells_style "Cancella", current_cell
            End If
        End If
    '
    '-- COLONNA DIREZIONE --------------------------------------------------------------------------------------
    '
    '-- COLONNA DIMENSIONE CORRISPONDENTE ----------------------------------------------------------------------
    '
    '-- COLONNA STATO ------------------------------------------------------------------------------------------
        If button_clicked = "Elimina Qk" Then
            Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 14), ws.Cells(start_row + 3 + tot, start_col + 15))      
        Else
            Set current_cell = ws.Range(ws.Cells(start_row + 3 + tot, start_col + 9), ws.Cells(start_row + 3 + tot, start_col + 10))
        End If
        If tot = 1 Then
            current_cell.Value = "-"
            cells_style "Condizione", current_cell
            current_cell.Validation.Delete
        Else
            cells_style "Cancella", current_cell
        End If
    '
End Sub