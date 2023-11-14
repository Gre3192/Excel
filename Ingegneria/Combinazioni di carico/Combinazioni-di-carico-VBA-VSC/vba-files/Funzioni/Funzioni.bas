'namespace=vba-files\Funzioni
Attribute VB_Name = "Funzioni"


Function range_pointer(Byval str As String) As String

    Dim start_line As Integer
    Dim start_col AS String, last_col As string 

    start_line = 3

    Select Case str

        Case "G1"
            start_col = "C"
            last_col = "H"
        Case "G2"
            start_col = "I"
            last_col = "N"
        Case "Qk"
            start_col = "O"
            last_col = "Y"
        Case "P"
            start_col = "Z"
            last_col = "AE"
        Case "E"
            start_col = "AF"
            last_col = "AK"
        Case "SLU"
            start_col = "AN"
            last_col = "AV"
        Case "SLE RARA"
            start_col = "AX"
            last_col = "BF"
        Case "SLE FREQUENTE"
            start_col = "BH"
            last_col = "BP"
        Case "SLE QUASI PERMANENTE"
            start_col = "BR"
            last_col = "BZ"
        Case "SISMICA"
            start_col = "CB"
            last_col = "CJ"
    End Select

    range_pointer = start_col & Cstr(start_line) & ":" & last_col & Cstr(start_line)

End Function   
'================================================================================================================================
Function getBlockName(ByVal button_clicked As String) As String

    Select Case button_clicked
        
        Case "Aggiungi G1", "Elimina G1", "Resetta G1"
            getBlockName = "G1"

        Case "Aggiungi G2", "Elimina G2", "Resetta G2"
            getBlockName = "G2"

        Case "Aggiungi Qk", "Elimina Qk", "Resetta Qk"
            getBlockName = "Qk" 

        Case "Aggiungi P", "Elimina P", "Resetta P"
            getBlockName = "P" 

        Case "Aggiungi E", "Elimina E", "Resetta E"
            getBlockName = "E"

        Case "Calcola SLU", "Resetta SLU"
            getBlockName = "SLU"

        Case "Calcola SLE RARA", "Resetta SLE RARA"
            getBlockName = "SLE RARA"

        Case "Calcola SLE FREQUENTE", "Resetta SLE FREQUENTE"
            getBlockName = "SLE FREQUENTE"

        Case "Calcola SLE Q.P.", "Resetta SLE Q.P."
            getBlockName = "SLE QUASI PERMANENTE"

        Case "Calcola SISMICA", "Resetta SISMICA"
            getBlockName = "SISMICA"

    End Select
    
End Function
'================================================================================================================================
Function isInputZone(ByVal button_clicked As String) As Boolean

    Select Case button_clicked
        
        Case "Aggiungi G1", "Elimina G1", "Resetta G1"
            isInputZone = true

        Case "Aggiungi G2", "Elimina G2", "Resetta G2"
            isInputZone = true

        Case "Aggiungi Qk", "Elimina Qk", "Resetta Qk"
            isInputZone = true 

        Case "Aggiungi P", "Elimina P", "Resetta P"
            isInputZone = true

        Case "Aggiungi E", "Elimina E", "Resetta E"
            isInputZone = true

        Case "Calcola SLU", "Resetta SLU"
            isInputZone = false

        Case "Calcola SLE RARA", "Resetta SLE RARA"
            isInputZone = false

        Case "Calcola SLE FREQUENTE", "Resetta SLE FREQUENTE"
            isInputZone = false

        Case "Calcola SLE Q.P.", "Resetta SLE Q.P."
            isInputZone = false

        Case "Calcola SISMICA", "Resetta SISMICA"
            isInputZone = false

    End Select
    
End Function
'================================================================================================================================
Function gamma(ByVal limit_state As String, ByVal load_type As String, ByVal value_condition As String, ByVal value_analysis As String) As Double

    If Not limit_state = "SLU" Then
        gamma = 1
        Exit Function

    Elseif load_type = "G1" Then
        If value_analy "EQU" Then
            gamma = IIf(value_condition = "Favorevole", 0.9, 1.1)
        Elseif value_analy "A1 (STR)" Then
            gamma = IIf(value_condition = "Favorevole", 1, 1.3)
        Elseif value_analy "A2" Then
            gamma = IIf(value_condition = "Favorevole", 1, 1)
        End If

    Elseif load_type = "G2" Then
        If value_analy "EQU" Then
            gamma = IIf(value_condition = "Favorevole", 0.8, 1.5)
        Elseif value_analy "A1 (STR)" Then
            gamma = IIf(value_condition = "Favorevole", 0.8, 1.5)
        Elseif value_analy "A2" Then
            gamma = IIf(value_condition = "Favorevole", 0.8, 1.3)
        End If

    Elseif load_type = "Qk" Then
        If value_analy "EQU" Then
            gamma = IIf(value_condition = "Favorevole", 0, 1.5)
        Elseif value_analy "A1 (STR)" Then
            gamma = IIf(value_condition = "Favorevole", 0, 1.5)
        Elseif value_analy "A2" Then
            gamma = IIf(value_condition = "Favorevole", 0, 1.3)
        End If

    End If

End Function
'================================================================================================================================
Function psi(ByVal norma As String, ByVal limit_state As String, ByVal num_psi As String, ByVal value_category As String) As Double

    Dim Index_cat As Integer
    Dim array_cat As Variant
    Dim psi_0_ntc08 As Variant, psi_1_ntc08 As Variant, psi_2_ntc08 As Variant
    Dim psi_0_ntc18 As Variant, psi_1_ntc18 As Variant, psi_2_ntc18 As Variant

    array_cat = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "Vento", "Neve (As " & ChrW(&H2264) & " 1000 m s.l.m.)", "Neve (As " & ChrW(&H3E) & " 1000 m s.l.m.)", "Variazioni termiche")
    psi_0_ntc08 = Array(0.7, 0.7, 0.7, 0.7, 1, 0.7, 0.7, 0, Null, Null, 0.6, 0.5, 0.7, 0.6)
    psi_0_ntc18 = Array(0.7, 0.7, 0.7, 0.7, 1, 0.7, 0.7, 0, Null, Null, 0.6, 0.5, 0.7, 0.6)
    psi_1_ntc08 = Array(0.5, 0.5, 0.7, 0.7, 0.9, 0.7, 0.5, 0, Null, Null, 0.2, 0.2, 0.5, 0.5)
    psi_1_ntc18 = Array(0.5, 0.5, 0.7, 0.7, 0.9, 0.7, 0.5, 0, Null, Null, 0.2, 0.2, 0.5, 0.5)
    psi_2_ntc08 = Array(0.3, 0.3, 0.6, 0.6, 0.8, 0.6, 0.3, 0, Null, Null, 0, 0, 0.2, 0)
    psi_2_ntc18 = Array(0.3, 0.3, 0.6, 0.6, 0.8, 0.6, 0.3, 0, Null, Null, 0, 0, 0.2, 0)

    For i = 1 To UBound(array_cat)
        If array_cat(i) = value_category Then
            Index_cat = i
        End If
    Next

    If norma = "NTC08" Then
        If num_psi = "Not" Then
            psi = 1

        Elseif num_psi = "0" Then
            psi = IIf(limit_state = "SLE FREQUENTE" Or limit_state = "SLE Q.P.", 1, psi_0_ntc08(Index_cat))

        Elseif num_psi = "1" Then
            psi = IIf(limit_state = "SLU" Or limit_state = "SLE Q.P." Or limit_state = "SLE rara", 1, psi_1_ntc08(Index_cat))

        Elseif num_psi = "2" Then
            psi = IIf(limit_state = "SLU" Or limit_state = "SLE RARA", 1, psi_2_ntc08(Index_cat))

        End If

    Elseif norma = "NTC18" Then
        If num_psi = "Not" Then
            psi = 1

        Elseif num_psi = "0" Then
            psi = IIf(limit_state = "SLE FREQUENTE" Or limit_state = "SLE Q.P.", 1, psi_0_ntc18(Index_cat))

        Elseif num_psi = "1" Then
            psi = IIf(limit_state = "SLU" Or limit_state = "SLE Q.P." Or limit_state = "SLE rara", 1, psi_1_ntc18(Index_cat))

        Elseif num_psi = "2" Then
            psi = IIf(limit_state = "SLU" Or limit_state = "SLE RARA", 1, psi_2_ntc18(Index_cat))

        End If

    End If

End Function
'================================================================================================================================
Function udm()
    void
End Function
'================================================================================================================================
Function formato_colore_celle(ByVal title_cell As String, ByRef cell_range As Range)

    Dim unisci_celle As Boolean

    Select Case title_cell

        Case "Cancella"
            cell_range.Clear
            unisci_celle = False

        Case "N°", "Combo", "Carico principale", "q progetto"
            With cell_range.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
            unisci_celle = False

        Case "Condizione", "Analisi", "Categoria"
            With cell_range
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
                With .Font
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = -0.499984740745262
                End With
            End With
            unisci_celle = True

        Case "Input carico"
            cell_range.Clear
            unisci_celle = False

        Case "Correlazione"
            cell_range.Clear
            unisci_celle = True

        Case "Direzione"
            unisci_celle = False

        Case "Dimensione Corrispondente"
            unisci_celle = True

        Case "Carico variabile principale", "q NTC08", "q NTC18"
            With cell_range.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            unisci_celle = True
            
    End Select

    If unisci_celle Then
        cell_range.Merge
    End If

    cell_range.HorizontalAlignment = xlCenter

End Function
'================================================================================================================================
Function validazione_celle(ByVal title_cell As String, ByRef cell_range As Range)

    Dim elenco_validazione As String, initializzazione As String

    Select Case title_cell
        Case "Condizione"
            elenco_validazione = "Sfavorevole,Favorevole"
            initializzazione = "Sfavorevole"
        Case "Analisi"
            elenco_validazione = "EQU, A1 (STR), A2"
            initializzazione = "A1 (STR)"
        Case "Categoria"
            elenco_validazione = "A,B,C,D,E,F,G,H,I,K,Vento,Neve (As " & ChrW(&H2264) & " 1000 m s.l.m.),Neve (As " & ChrW(&H3E) & " 1000 m s.l.m.),Variazioni termiche"
            initializzazione = "A"
    End Select

    With cell_range.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=elenco_validazione
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Errore"
        .InputMessage = ""
        .ErrorMessage = "Immettere uno dei valori in elenco!"
        .ShowInput = True
        .ShowError = True
    End With

    cell_range.Value = initializzazione

End Function
'================================================================================================================================
Function reset(ByVal button_clicked As String)

    Dim ws As Worksheet
    Dim j As Integer
    Dim current_cell As Range
    Dim isDash As Boolean
    Dim last_row As Long, start_row As Long, start_col As Long, last_col As Long, cell_tot As Long
    Dim button_involved() As Variant

    Set ws = Application.ThisWorkbook.ActiveSheet

    start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
    start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column
    last_row = ws.Cells(Rows.Count, start_col).End(xlUp).Row
    last_col = ws.Range(range_pointer(getBlockName(button_clicked))).Columns(ws.Range(range_pointer(getBlockName(button_clicked))).Columns.Count).Column
    
    isDash = IIf(ws.Cells(start_row + 1, start_col) = "-", True, False)


    '-- COLONNA TOT --------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 1, start_col)
        If isDash Then
            Exit Function
        Else
            cell_tot = current_cell.Value
            current_cell.Value = "-"
        End If
    '
    '-- COLONNA N° o COMBO -------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 4, start_col)
        current_cell.Value = "-"
        formato_colore_celle "N°", current_cell
    '
    if isInputZone(button_clicked) then
        '-- COLONNA INPUT CARICO -----------------------------------------------------------------------------------
            Set current_cell = ws.Cells(start_row + 4, start_col + 1)
            formato_colore_celle "Cancella", current_cell
        '
        '-- COLONNA CORRELAZIONE -----------------------------------------------------------------------------------
            Set current_cell = ws.Range(Cells(start_row + 4, start_col + 2), Cells(start_row + 4, start_col + 3))
            '.......................................................................
            current_cell.Value = "-"
            formato_colore_celle "Correlazione", current_cell
        '
        '-- COLONNA CONDIZIONE -------------------------------------------------------------------------------------
            incr_qk = IIf(button_clicked = "Resetta Qk", 2, 0)
            Set current_cell = ws.Range(Cells(start_row + 4, start_col + 2 + incr_qk), Cells(start_row + 4, start_col + 3 + incr_qk))
            current_cell.Value = "-"
            formato_colore_celle "Condizione", current_cell
            current_cell.Validation.Delete
        '
        '-- COLONNA ANALISI ----------------------------------------------------------------------------------------
            Set current_cell = ws.Cells(start_row + 4, start_col + 4 + incr_qk)
            current_cell.Value = "-"
            formato_colore_celle "Analisi", current_cell
            current_cell.Validation.Delete
        '
        '-- COLONNA CATEGORIA --------------------------------------------------------------------------------------
            If button_clicked = "Resetta Qk" Then
                Set current_cell = ws.Range(Cells(start_row + 4, start_col + 7), Cells(start_row + 4, start_col + 9))
                current_cell.Value = "-"
                formato_colore_celle "Categoria", current_cell
                current_cell.Validation.Delete
            End If
        '
        '-- COLONNA DIREZIONE --------------------------------------------------------------------------------------
        '
        '-- COLONNA DIMENSIONE CORRISPONDENTE ----------------------------------------------------------------------
        '
    Else
        '-- COLONNA CARICO VARIABILE PRINCIPALE --------------------------------------------------------------------
            Set current_cell = ws.Cells(start_row + 4, start_col + 1)
            current_cell.Value = "-"
            formato_colore_celle "Carico variabile principale", current_cell
        '
        '-- COLONNA q NTC08 ----------------------------------------------------------------------------------------
            Set current_cell = ws.Cells(start_row + 4, start_col + 4)
            current_cell.Value = "-"
            formato_colore_celle "q NTC08", current_cell
        '
        '-- COLONNA q NTC18 ----------------------------------------------------------------------------------------
            Set current_cell = ws.Cells(start_row + 4, start_col + 6)
            current_cell.Value = "-"
            formato_colore_celle "q NTC18", current_cell
        '
    end if

    
    If Not cell_tot = 1 Then
        Range(Cells(start_row + 5, start_col), Cells(last_row, last_col - 1)).Clear
    End If
    
End Function
