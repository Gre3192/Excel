'namespace=vba-files\Moduli
Attribute VB_Name = "AggiungiCarico"

Sub aggiungi_carico()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Dim current_cell As Range, next_cell As Range
    Dim isDash As Boolean
    Dim incr As Integer
    Dim last_row As Long, start_row As Long, start_col As Long
    Dim button_clicked As String

    Set ws = Application.ThisWorkbook.ActiveSheet
    button_clicked = Application.caller 

    start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
    start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column
    last_row = ws.Cells(Rows.Count, start_col).End(xlUp).Row

    isDash = IIf(ws.Cells(start_row + 1, start_col) = "-", True, False)
    incr = IIf(isDash, 0, 1)


    '-- COLONNA TOT --------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 1, start_col)
        If isDash Then
            current_cell.Value = 1
        Else
            current_cell.Value = current_cell.Value + 1
        End If
    '
    '-- COLONNA N° ---------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(last_row, start_col)
        Set next_cell = ws.Cells(last_row + incr, start_col)       
        If isDash Then
            current_cell.Value = 1
        Else
            next_cell.Value = current_cell.Value + 1
        End If
        formato_colore_celle "N°", next_cell
    '
    '-- COLONNA INPUT CARICO -----------------------------------------------------------------------------------
        Set next_cell = ws.Cells(last_row + incr, start_col + 1)    
        formato_colore_celle "Input carico", next_cell
    '
    '-- COLONNA CORRELAZIONE -----------------------------------------------------------------------------------
        Set next_cell = ws.Range(Cells(last_row + incr, start_col + 2), Cells(last_row + incr, start_col + 3))      
        formato_colore_celle "Correlazione", next_cell
    '
    '-- COLONNA CONDIZIONE -------------------------------------------------------------------------------------
        incr_qk = IIf(button_clicked = "Aggiungi Qk", 2, 0)
        Set next_cell = ws.Range(Cells(last_row + incr, start_col + 2 + incr_qk), Cells(last_row + incr, start_col + 3 + incr_qk))        
        formato_colore_celle "Condizione", next_cell
        validazione_celle "Condizione", next_cell
    '    
    '-- COLONNA ANALISI ----------------------------------------------------------------------------------------
        Set next_cell = ws.Cells(last_row + incr, start_col + 4 + incr_qk) 
        formato_colore_celle "Analisi", next_cell
        validazione_celle "Analisi", next_cell
    '
    '-- COLONNA CATEGORIA --------------------------------------------------------------------------------------
        If button_clicked = "Aggiungi Qk" Then
            Set next_cell = ws.Range(Cells(last_row + incr, start_col + 7), Cells(last_row + incr, start_col + 9))     
            formato_colore_celle "Categoria", next_cell
            validazione_celle "Categoria", next_cell
        End If
    '
    '-- COLONNA DIREZIONE --------------------------------------------------------------------------------------
    '
    '-- COLONNA DIMENSIONE CORRISPONDENTE ----------------------------------------------------------------------
    '
End Sub

