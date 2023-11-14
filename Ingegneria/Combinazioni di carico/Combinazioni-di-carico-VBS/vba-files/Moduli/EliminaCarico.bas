'namespace=vba-files\Moduli
Attribute VB_Name = "EliminaCarico"

Sub elimina_carico()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Dim current_cell As Range
    Dim isDash As Boolean
    Dim last_row As Long, start_row As Long, start_col As Long
    Dim cell_tot_value As Variant
    Dim button_clicked As String

    Set ws = Application.ThisWorkbook.ActiveSheet
    button_clicked = Application.Caller

    start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
    start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column
    last_row = ws.Cells(Rows.Count, start_col).End(xlUp).Row
    
    isDash = IIf(ws.Cells(start_row + 1, start_col) = "-", True, False)


    '-- COLONNA TOT --------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 1, start_col)
        cell_tot_value = current_cell.Value
        If isDash Then
            Exit Sub
        Elseif current_cell.Value = 1 Then
            current_cell.Value = "-"
        Else
            current_cell.Value = current_cell.Value - 1
        End If
    '
    '-- COLONNA N° ---------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(last_row, start_col)
        If cell_tot_value = 1 Then
            current_cell.Value = "-"
            formato_colore_celle "N°", current_cell
        Else
            current_cell.Value = current_cell.Value - 1
            formato_colore_celle "Cancella", current_cell
        End If
    '
    '-- COLONNA INPUT CARICO -----------------------------------------------------------------------------------
        Set current_cell = ws.Cells(last_row, start_col + 1)
        formato_colore_celle "Input carico", current_cell
    '
    '-- COLONNA CORRELAZIONE -----------------------------------------------------------------------------------
        Set current_cell = ws.Range(Cells(last_row, start_col + 2), Cells(last_row, start_col + 3))
        If cell_tot_value = 1 Then
            current_cell.Value = "-"
            formato_colore_celle "Correlazione", current_cell
        Else
            formato_colore_celle "Cancella", current_cell
        End If
    '
    '-- COLONNA CONDIZIONE -------------------------------------------------------------------------------------
        incr_qk = IIf(button_clicked = "Elimina Qk", 2, 0)
        Set current_cell = ws.Range(Cells(last_row, start_col + 2 + incr_qk), Cells(last_row, start_col + 3 + incr_qk))
        If cell_tot_value = 1 Then
            current_cell.Value = "-"
            formato_colore_celle "Condizione", current_cell
            current_cell.Validation.Delete
        Else
            formato_colore_celle "Cancella", current_cell
        End If
    '      
    '-- COLONNA ANALISI ----------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(last_row, start_col + 4 + incr_qk)
        If cell_tot_value = 1 Then
            current_cell.Value = "-"
            formato_colore_celle "Analisi", current_cell
            current_cell.Validation.Delete
        Else
            formato_colore_celle "Cancella", current_cell
        End If
    '
    '-- COLONNA CATEGORIA --------------------------------------------------------------------------------------
        If button_clicked = "Elimina Qk" Then
            Set current_cell = ws.Range(Cells(last_row, start_col + 7), Cells(last_row, start_col + 9))
            If cell_tot_value = 1 Then
                current_cell.Value = "-"
                formato_colore_celle "Categoria", current_cell
                current_cell.Validation.Delete
            Else
                formato_colore_celle "Cancella", current_cell
            End If
        End If
    '
    '-- COLONNA DIREZIONE --------------------------------------------------------------------------------------
    '
    '-- COLONNA DIMENSIONE CORRISPONDENTE ----------------------------------------------------------------------
    '

End Sub



