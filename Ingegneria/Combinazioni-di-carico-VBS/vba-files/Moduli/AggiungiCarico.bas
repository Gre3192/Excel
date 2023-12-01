Attribute VB_Name = "AggiungiCarico"
'namespace=vba-files\Moduli

Sub aggiungi_carico()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = Application.ThisWorkbook.ActiveSheet

    Dim button_clicked As String
    button_clicked = Application.caller

    Dim isDash As Boolean
    Dim incrRow As Integer, incrCol As Integer
    Dim current_cell As Range, next_cell As Range
    Dim start_row As Long, start_col As Long, tot As Long

    start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
    start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column

    isDash = IIf(ws.Cells(start_row + 1, start_col) = "-", True, False)
    incrRow = IIf(isDash, 0, 1)
    incrCol = IIf(button_clicked = "Aggiungi Qk", 2, 0)

    '-- COLONNA TOT --------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 1, start_col)
        If isDash Then
            tot = 1
            current_cell.Value = tot
        Else
            tot = current_cell.Value
            current_cell.Value = tot + 1
        End If
    '
    '-- COLONNA N° ---------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 3 + tot, start_col)
        Set next_cell = ws.Cells(start_row + 3 + tot + incrRow, start_col)
        If isDash Then
            current_cell.Value = 1
        Else
            next_cell.Value = current_cell.Value + 1
        End If
        cells_style "N°", next_cell
    '
    '-- COLONNA DESCRIZIONE ------------------------------------------------------------------------------------
        Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 1), ws.Cells(start_row + 3 + tot + incrRow, start_col + 3))
        cells_style "Descrizione", next_cell
    '
    '-- COLONNA CORRELAZIONE -----------------------------------------------------------------------------------
        If button_clicked = "Aggiungi Qk" Then
            Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 4), ws.Cells(start_row + 3 + tot + incrRow, start_col + 5))
            cells_style "Correlazione", next_cell
        End If
    '
    '-- COLONNA INPUT CARICO -----------------------------------------------------------------------------------
        Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 4 + incrCol), ws.Cells(start_row + 3 + tot + incrRow, start_col + 5 + incrCol))
        cells_style "Input carico", next_cell
    '
    '-- COLONNA CONDIZIONE -------------------------------------------------------------------------------------
        Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 6 + incrCol), ws.Cells(start_row + 3 + tot + incrRow, start_col + 7 + incrCol))
        cells_style "Condizione", next_cell
        cells_valid "Condizione", next_cell
    '
    '-- COLONNA ANALISI ----------------------------------------------------------------------------------------
        Set next_cell = ws.Cells(start_row + 3 + tot + incrRow, start_col + 8 + incrCol)
        cells_style "Analisi", next_cell
        cells_valid "Analisi", next_cell
    '
    '-- COLONNA CATEGORIA --------------------------------------------------------------------------------------
        If button_clicked = "Aggiungi Qk" Then
            Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 11), ws.Cells(start_row + 3 + tot + incrRow, start_col + 13))
            cells_style "Categoria", next_cell
            cells_valid "Categoria", next_cell
        End If
    '
    '-- COLONNA DIREZIONE --------------------------------------------------------------------------------------
    '
    '-- COLONNA DIMENSIONE CORRISPONDENTE ----------------------------------------------------------------------
    '
    '-- COLONNA STATO ------------------------------------------------------------------------------------------
        If button_clicked = "Aggiungi Qk" Then
            Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 14), ws.Cells(start_row + 3 + tot + incrRow, start_col + 15))
        Else
            Set next_cell = ws.Range(ws.Cells(start_row + 3 + tot + incrRow, start_col + 9), ws.Cells(start_row + 3 + tot + incrRow, start_col + 10))
        End If
        cells_style "Stato", next_cell
        cells_valid "Stato", next_cell
    '
End Sub
