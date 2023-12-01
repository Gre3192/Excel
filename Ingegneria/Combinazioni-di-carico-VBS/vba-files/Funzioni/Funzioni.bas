Attribute VB_Name = "Funzioni"
'namespace=vba-files\Funzioni

Function range_pointer(ByVal str As String) As String

    Dim start_line As Integer
    Dim start_col As String, last_col As String

    start_line = 3

    Select Case str

        Case "G1"
            start_col = "C"
            last_col = "N"
        Case "G2"
            start_col = "O"
            last_col = "Z"
        Case "Qk"
            start_col = "AA"
            last_col = "AQ"
        Case "P"
            start_col = "AR"
            last_col = "BC"
        Case "E"
            start_col = "BD"
            last_col = "BO"
        Case "SLU"
            start_col = "BR"
            last_col = "CC"
        Case "SLE RARA"
            start_col = "CE"
            last_col = "CP"
        Case "SLE FREQUENTE"
            start_col = "CR"
            last_col = "DC"
        Case "SLE QUASI PERMANENTE"
            start_col = "DE"
            last_col = "DJ"
        Case "SISMICA"
            start_col = "DL"
            last_col = "DT"
    End Select

    range_pointer = start_col & CStr(start_line) & ":" & last_col & CStr(start_line)

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
            isInputZone = True

        Case "Aggiungi G2", "Elimina G2", "Resetta G2"
            isInputZone = True

        Case "Aggiungi Qk", "Elimina Qk", "Resetta Qk"
            isInputZone = True

        Case "Aggiungi P", "Elimina P", "Resetta P"
            isInputZone = True

        Case "Aggiungi E", "Elimina E", "Resetta E"
            isInputZone = True

        Case "Calcola SLU", "Resetta SLU"
            isInputZone = False

        Case "Calcola SLE RARA", "Resetta SLE RARA"
            isInputZone = False

        Case "Calcola SLE FREQUENTE", "Resetta SLE FREQUENTE"
            isInputZone = False

        Case "Calcola SLE Q.P.", "Resetta SLE Q.P."
            isInputZone = False

        Case "Calcola SISMICA", "Resetta SISMICA"
            isInputZone = False

    End Select
    
End Function
'================================================================================================================================
Function getGamma(ByVal norma As String, ByVal stateLimit As String, ByVal load_type As String, ByVal value_condition As String, ByVal value_analysis As String) As Double

    Dim gamma As Double
    
    If norma = "NTC08" Then

        If stateLimit <> "SLU" Then

            gamma = 1
    
        ElseIf load_type = "G1" Then
    
            If value_analysis = "EQU" Then
                gamma = IIf(value_condition = "Favorevole", 0.9, 1.1)
            ElseIf value_analysis = "A1 (STR)" Then
                gamma = IIf(value_condition = "Favorevole", 1, 1.3)
            ElseIf value_analysis = "A2" Then
                gamma = IIf(value_condition = "Favorevole", 1, 1)
            End If
    
        ElseIf load_type = "G2" Then
    
            If value_analysis = "EQU" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.5)
            ElseIf value_analysis = "A1 (STR)" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.5)
            ElseIf value_analysis = "A2" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.3)
            End If
    
        ElseIf load_type = "Qk" Then
    
            If value_analysis = "EQU" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.5)
            ElseIf value_analysis = "A1 (STR)" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.5)
            ElseIf value_analysis = "A2" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.3)
            End If
    
        End If

    ElseIf norma = "NTC18" Then

        If stateLimit <> "SLU" Then

            gamma = 1
    
        ElseIf load_type = "G1" Then
    
            If value_analysis = "EQU" Then
                gamma = IIf(value_condition = "Favorevole", 0.9, 1.1)
            ElseIf value_analysis = "A1 (STR)" Then
                gamma = IIf(value_condition = "Favorevole", 1, 1.3)
            ElseIf value_analysis = "A2" Then
                gamma = IIf(value_condition = "Favorevole", 1, 1)
            End If
    
        ElseIf load_type = "G2" Then
    
            If value_analysis = "EQU" Then
                gamma = IIf(value_condition = "Favorevole", 0.8, 1.5)
            ElseIf value_analysis = "A1 (STR)" Then
                gamma = IIf(value_condition = "Favorevole", 0.8, 1.5)
            ElseIf value_analysis = "A2" Then
                gamma = IIf(value_condition = "Favorevole", 0.8, 1.3)
            End If
    
        ElseIf load_type = "Qk" Then
    
            If value_analysis = "EQU" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.5)
            ElseIf value_analysis = "A1 (STR)" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.5)
            ElseIf value_analysis = "A2" Then
                gamma = IIf(value_condition = "Favorevole", 0, 1.3)
            End If
    
        End If

    End If

    ' debug.print norma, stateLimit, value_analysis, value_condition, "gamma = " & Cstr(gamma)
    getGamma = gamma

End Function
'================================================================================================================================
Function getPsi(ByVal norma As String, ByVal stateLimit As String, ByVal num_psi As String, ByVal value_category As String) As Double

    Dim Index_cat As Integer
    Dim array_cat() As Variant
    Dim psi As Double
    Dim psi_0_ntc08 As Variant, psi_1_ntc08 As Variant, psi_2_ntc08 As Variant
    Dim psi_0_ntc18 As Variant, psi_1_ntc18 As Variant, psi_2_ntc18 As Variant

    array_cat = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "Vento", "Neve (as " & ChrW(&H2264) & " 1000 m.s.l.m.)", "Neve (as " & ChrW(&H3E) & " 1000 m.s.l.m.)", "Variazioni termiche")
    psi_0_ntc08 = Array(0.7, 0.7, 0.7, 0.7, 1, 0.7, 0.7, 0, Null, Null, 0.6, 0.5, 0.7, 0.6)
    psi_0_ntc18 = Array(0.7, 0.7, 0.7, 0.7, 1, 0.7, 0.7, 0, Null, Null, 0.6, 0.5, 0.7, 0.6)
    psi_1_ntc08 = Array(0.5, 0.5, 0.7, 0.7, 0.9, 0.7, 0.5, 0, Null, Null, 0.2, 0.2, 0.5, 0.5)
    psi_1_ntc18 = Array(0.5, 0.5, 0.7, 0.7, 0.9, 0.7, 0.5, 0, Null, Null, 0.2, 0.2, 0.5, 0.5)
    psi_2_ntc08 = Array(0#, 0#, 0.6, 0.6, 0.8, 0.6, 0#, 0, Null, Null, 0, 0, 0.2, 0)
    psi_2_ntc18 = Array(0#, 0#, 0.6, 0.6, 0.8, 0.6, 0#, 0, Null, Null, 0, 0, 0.2, 0)

    For i = 0 To UBound(array_cat)

        If array_cat(i) = value_category Then
            Index_cat = i
        End If

    Next

    If norma = "NTC08" Then

        If num_psi = "NotNum" Then
            psi = 1
        ElseIf num_psi = "0" Then
            psi = IIf(stateLimit = "SLE FREQUENTE" Or stateLimit = "SLE QUASI PERMANENTE", 1, psi_0_ntc08(Index_cat))
        ElseIf num_psi = "1" Then
            psi = IIf(stateLimit = "SLU" Or stateLimit = "SLE QUASI PERMANENTE" Or stateLimit = "SLE RARA", 1, psi_1_ntc08(Index_cat))
        ElseIf num_psi = "2" Then
            psi = IIf(stateLimit = "SLU" Or stateLimit = "SLE RARA", 1, psi_2_ntc08(Index_cat))
        End If

    ElseIf norma = "NTC18" Then

        If num_psi = "NotNum" Then
            psi = 1
        ElseIf num_psi = "0" Then
            psi = IIf(stateLimit = "SLE FREQUENTE" Or stateLimit = "SLE QUASI PERMANENTE", 1, psi_0_ntc18(Index_cat))
        ElseIf num_psi = "1" Then
            psi = IIf(stateLimit = "SLU" Or stateLimit = "SLE QUASI PERMANENTE" Or stateLimit = "SLE RARA", 1, psi_1_ntc18(Index_cat))
        ElseIf num_psi = "2" Then
            psi = IIf(stateLimit = "SLU" Or stateLimit = "SLE RARA", 1, psi_2_ntc18(Index_cat))
        End If

    End If

    ' debug.print norma, stateLimit, "Cat. " & value_category, "psi_" & num_psi & " = " & Cstr(psi)
    getPsi = psi

End Function
'================================================================================================================================
Function udm_force(ByVal unit As String) As Double

    Select Case unit

        Case "QN", "anti-qN"
            udm_force = 10 ^ 30
        Case "RN", "anti-rN"
            udm_force = 10 ^ 27
        Case "YN", "anti-yN"
            udm_force = 10 ^ 24
        Case "ZN", "anti-zN"
            udm_force = 10 ^ 21
        Case "EN", "anti-aN"
            udm_force = 10 ^ 18
        Case "PN", "anti-fN"
            udm_force = 10 ^ 15
        Case "TN", "anti-pN"
            udm_force = 10 ^ 12
        Case "GN", "anti-nN"
            udm_force = 10 ^ 9
        Case "MN", "anti-muN"
            udm_force = 10 ^ 6
        Case "kN", "anti-mN"
            udm_force = 10 ^ 3
        Case "hN", "anti-cN"
            udm_force = 10 ^ 2
        Case "daN", "anti-dN"
            udm_force = 10 ^ 1
        Case "-", "anti--", "N", "anti-N"
            udm_force = 1
        Case "dN", "anti-daN"
            udm_force = 10 ^ -1
        Case "cN", "anti-hN"
            udm_force = 10 ^ -2
        Case "mN", "anti-kN"
            udm_force = 10 ^ -3
        Case "muN", "anti-MN"
            udm_force = 10 ^ -6
        Case "nN", "anti-GN"
            udm_force = 10 ^ -9
        Case "pN", "anti-TN"
            udm_force = 10 ^ -12
        Case "fN", "anti-PN"
            udm_force = 10 ^ -15
        Case "aN", "anti-EN"
            udm_force = 10 ^ -18
        Case "zN", "anti-ZN"
            udm_force = 10 ^ -21
        Case "yN", "anti-YN"
            udm_force = 10 ^ -24
        Case "rN", "anti-RN"
            udm_force = 10 ^ -27
        Case "qN", "anti-QN"
            udm_force = 10 ^ -30

    End Select

End Function
'================================================================================================================================
Function udm_meter(ByVal unit As String) As Double

    Dim exponent As Integer
    Dim exponentString As String
    exponent = IIf(IsNumeric(Right(unit, 1)), Right(unit, 1), 1)
    exponentString = IIf(IsNumeric(Right(unit, 1)), Right(unit, 1), "")
    
    Select Case unit

        Case "Qm" & exponentString, "anti-qm" & exponentString
            udm_meter = (10 ^ 30) ^ exponent
        Case "Rm" & exponentString, "anti-rm" & exponentString
            udm_meter = (10 ^ 27) ^ exponent
        Case "Ym" & exponentString, "anti-ym" & exponentString
            udm_meter = (10 ^ 24) ^ exponent
        Case "Zm" & exponentString, "anti-zm" & exponentString
            udm_meter = (10 ^ 21) ^ exponent
        Case "Em" & exponentString, "anti-am" & exponentString
            udm_meter = (10 ^ 18) ^ exponent
        Case "Pm" & exponentString, "anti-fm" & exponentString
            udm_meter = (10 ^ 15) ^ exponent
        Case "Tm" & exponentString, "anti-pm" & exponentString
            udm_meter = (10 ^ 12) ^ exponent
        Case "Gm" & exponentString, "anti-nm" & exponentString
            udm_meter = (10 ^ 9) ^ exponent
        Case "Mm" & exponentString, "anti-mmN" & exponentString
            udm_meter = (10 ^ 6) ^ exponent
        Case "km" & exponentString, "anti-mm" & exponentString
            udm_meter = (10 ^ 3) ^ exponent
        Case "hm" & exponentString, "anti-cm" & exponentString
            udm_meter = (10 ^ 2) ^ exponent
        Case "dam" & exponentString, "anti-dm" & exponentString
            udm_meter = (10 ^ 1) ^ exponent
        Case "-", "anti--", "m" & exponentString, "anti-m" & exponentString
            udm_meter = 1
        Case "dm" & exponentString, "anti-dmN" & exponentString
            udm_meter = (10 ^ -1) ^ exponent
        Case "cm" & exponentString, "anti-hm" & exponentString
            udm_meter = (10 ^ -2) ^ exponent
        Case "mm" & exponentString, "anti-km" & exponentString
            udm_meter = (10 ^ -3) ^ exponent
        Case "mum" & exponentString, "anti-Mm" & exponentString
            udm_meter = (10 ^ -6) ^ exponent
        Case "nm" & exponentString, "anti-Gm" & exponentString
            udm_meter = (10 ^ -9) ^ exponent
        Case "pm" & exponentString, "anti-Tm" & exponentString
            udm_meter = (10 ^ -12) ^ exponent
        Case "fm" & exponentString, "anti-Pm" & exponentString
            udm_meter = (10 ^ -15) ^ exponent
        Case "am" & exponentString, "anti-Em" & exponentString
            udm_meter = (10 ^ -18) ^ exponent
        Case "zm" & exponentString, "anti-Zm" & exponentString
            udm_meter = (10 ^ -21) ^ exponent
        Case "ym" & exponentString, "anti-Ym" & exponentString
            udm_meter = (10 ^ -24) ^ exponent
        Case "rm" & exponentString, "anti-Rm" & exponentString
            udm_meter = (10 ^ -27) ^ exponent
        Case "qm" & exponentString, "anti-Qm" & exponentString
            udm_meter = (10 ^ -30) ^ exponent

    End Select

End Function
'================================================================================================================================
Function udm() As Double

    Dim ws As Worksheet
    Set ws = Application.ThisWorkbook.ActiveSheet

    Dim UDM1 As Double, UDM2 As Double, UDM3 As Double, UDMValue As Double
    Dim antiUDM1 As Double, antiUDM2 As Double, antiUDM3 As Double, antiUDMValue As Double

    UDM1 = udm_force(ws.Range("A6").Value)
    UDM2 = udm_meter(ws.Range("B6").Value)
    UDM3 = udm_meter(ws.Range("A7").Value)
    UDMValue = UDM1 * UDM2 / UDM3

    antiUDM1 = udm_force("anti-" & ws.Range("A9").Value)
    antiUDM2 = udm_meter("anti-" & ws.Range("B9").Value)
    antiUDM3 = udm_meter("anti-" & ws.Range("A10").Value)
    antiUDMValue = antiUDM1 * antiUDM2 / antiUDM3

    udm = UDMValue * antiUDMValue

End Function
'================================================================================================================================
Function cells_style(ByVal col_Name As String, ByRef col_Range As Range)

    Dim ws As Worksheet
    Set ws = Application.ThisWorkbook.ActiveSheet
    Dim mergeCells As Boolean
    Dim decNum As Integer

    Select Case col_Name

        Case "MaxColor"
             col_Range.Font.Color = RGB(255, 0, 0)

        Case "Cancella"
            col_Range.Clear
            mergeCells = False

        Case "N°", "Combo"
            With col_Range
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.149998474074526
                    .PatternTintAndShade = 0
                End With
            End With
            mergeCells = False

        Case "Descrizione"
            With col_Range
                With .Font
                    .Italic = True
                End With
            End With
            mergeCells = True

        Case "Condizione", "Analisi", "Categoria"
            With col_Range
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
            mergeCells = True

        Case "Input carico"
            col_Range.Clear
            mergeCells = True

        Case "Correlazione"
            col_Range.Clear
            mergeCells = True

        Case "Direzione"
            mergeCells = False

        Case "Dimensione Corrispondente"
            mergeCells = True

        Case "Carichi variabili principali", "Annesse categorie principali", "q - NTC08", "q - NTC18"
            With col_Range
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
                If col_Name = "q - NTC08" Or col_Name = "q - NTC18" Then
                    decNum = ws.Range("A12").Value
                    .NumberFormat = "0" & IIf(decNum <> 0, ".", "") & String(decNum, "0")
                End If
            End With
            mergeCells = True

        Case "Stato"
            col_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
            Formula1:="=""Attivo"""
            col_Range.FormatConditions(col_Range.FormatConditions.Count).SetFirstPriority
            With col_Range.FormatConditions(1).Font
                .Color = -16752384
                .TintAndShade = 0
            End With
            With col_Range.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13561798
                .TintAndShade = 0
            End With
            col_Range.FormatConditions(1).StopIfTrue = False
            col_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=""Non Attivo"""
            col_Range.FormatConditions(col_Range.FormatConditions.Count).SetFirstPriority
            With col_Range.FormatConditions(1).Font
                .Color = -16383844
                .TintAndShade = 0
            End With
            With col_Range.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
            End With
            col_Range.FormatConditions(1).StopIfTrue = False
            mergeCells = True
            
    End Select

    If mergeCells Then
        col_Range.Merge
    End If

    col_Range.HorizontalAlignment = xlCenter
    col_Range.VerticalAlignment = xlCenter

End Function
'================================================================================================================================
Function cells_valid(ByVal col_Name As String, ByRef col_Range As Range)

    Dim wsUtils As Worksheet
    Set wsUtils = Application.ThisWorkbook.Worksheets("Utils")
    
    Dim start_valid As String
    Dim array_valid As Range

    Select Case col_Name
        Case "Condizione"
            Set array_valid = wsUtils.Range("C4:C17")
            start_valid = "Sfavorevole"
        Case "Analisi"
            Set array_valid = wsUtils.Range("E4:E17")
            start_valid = Application.ThisWorkbook.Worksheets("Combinazioni").Range("A15").Value
        Case "Categoria"
            Set array_valid = wsUtils.Range("F4:F17")
            start_valid = "A"
        Case "Stato"
            Set array_valid = wsUtils.Range("I4:I17")
            start_valid = "Attivo"
    End Select

    With col_Range
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(Application.Transpose(array_valid), ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .errorTitle = "Errore"
            .InputMessage = ""
            .ErrorMessage = "Immettere uno dei valori in elenco!"
            .ShowInput = True
            .ShowError = True
        End With
        .Value = start_valid
    End With

End Function
'================================================================================================================================
Function reset(ByVal button_clicked As String)

    Dim ws As Worksheet
    Dim j As Integer, incrCol As Integer
    Dim current_cell As Range
    Dim isDash As Boolean
    Dim start_row As Long, start_col As Long, last_col As Long, tot As Long
    Dim Button_involved() As Variant

    Set ws = Application.ThisWorkbook.ActiveSheet

    start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
    start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column
    last_col = ws.Range(range_pointer(getBlockName(button_clicked))).Columns(ws.Range(range_pointer(getBlockName(button_clicked))).Columns.Count).Column

    '-- COLONNA TOT --------------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 1, start_col)
        isDash = IIf(current_cell.Value = "-", True, False)
        If isDash Then
            Exit Function
        Else
            tot = current_cell.Value
            current_cell.Value = "-"
        End If
    '
    If Not tot = 1 Then
        Range(ws.Cells(start_row + 4, start_col), ws.Cells(start_row + 4 + tot, last_col - 1)).Clear
    End If
    '-- COLONNA N° o COMBO -------------------------------------------------------------------------------------
        Set current_cell = ws.Cells(start_row + 4, start_col)
        cells_style "N°", current_cell
        current_cell.Value = "-"
    '
    If isInputZone(button_clicked) Then

        incrCol = IIf(button_clicked = "Resetta Qk", 2, 0)

        '-- COLONNA DESCRIZIONE ------------------------------------------------------------------------------------
            Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 1), ws.Cells(start_row + 4, start_col + 3))
            cells_style "Descrizione", current_cell
        '
        '-- COLONNA CORRELAZIONE -----------------------------------------------------------------------------------
            If button_clicked = "Resetta Qk" Then
                Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 4), ws.Cells(start_row + 4, start_col + 5))
                cells_style "Correlazione", current_cell
            End If
        '
        '-- COLONNA INPUT CARICO -----------------------------------------------------------------------------------
            Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 4 + incrCol), ws.Cells(start_row + 4, start_col + 5 + incrCol))
            cells_style "Cancella", current_cell
        '
        '-- COLONNA CONDIZIONE -------------------------------------------------------------------------------------
            Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 6 + incrCol), ws.Cells(start_row + 4, start_col + 7 + incrCol))
            cells_style "Condizione", current_cell
            current_cell.Value = "-"
            'current_cell.Validation.Delete
        '
        '-- COLONNA ANALISI ----------------------------------------------------------------------------------------
            Set current_cell = ws.Cells(start_row + 4, start_col + 8 + incrCol)
            cells_style "Analisi", current_cell
            current_cell.Value = "-"
            'current_cell.Validation.Delete
        '
        '-- COLONNA CATEGORIA --------------------------------------------------------------------------------------
            If button_clicked = "Resetta Qk" Then
                Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 11), ws.Cells(start_row + 4, start_col + 13))
                cells_style "Categoria", current_cell
                current_cell.Value = "-"
                'current_cell.Validation.Delete
            End If
        '
        '-- COLONNA DIREZIONE --------------------------------------------------------------------------------------
        '
        '-- COLONNA DIMENSIONE CORRISPONDENTE ----------------------------------------------------------------------
        '
        '-- COLONNA STATO ------------------------------------------------------------------------------------------
            If button_clicked = "Resetta Qk" Then
                Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 14), ws.Cells(start_row + 4, start_col + 15))
            Else
                Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 9), ws.Cells(start_row + 4, start_col + 10))
            End If
            cells_style "Condizione", current_cell
            current_cell.Value = "-"
            'current_cell.Validation.Delete
        '
    Else

        incrCol = IIf(button_clicked = "Resetta SLE Q.P.", -6, 0)

        '-- COLONNA CARICHI VARIABILI PRINCIPALI -------------------------------------------------------------------
            If button_clicked <> "Resetta SLE Q.P." Then
                Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 1), ws.Cells(start_row + 4, start_col + 3))
                cells_style "Carichi variabili principali", current_cell
                current_cell.Value = "-"
            End If
        '
        '-- COLONNA ANNESSE CATEGORIE PRINCIPALI -------------------------------------------------------------------
            If button_clicked <> "Resetta SLE Q.P." Then
                Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 4), ws.Cells(start_row + 4, start_col + 6))
                cells_style "Annesse categorie principali", current_cell
                current_cell.Value = "-"
            End If
        '
        '-- COLONNA q - NTC08 --------------------------------------------------------------------------------------
            Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 7 + incrCol), ws.Cells(start_row + 4, start_col + 8 + incrCol))
            cells_style "q - NTC08", current_cell
            current_cell.Value = "-"
        '
        '-- COLONNA q - NTC18 --------------------------------------------------------------------------------------
            Set current_cell = ws.Range(ws.Cells(start_row + 4, start_col + 9 + incrCol), ws.Cells(start_row + 4, start_col + 10 + incrCol))
            cells_style "q - NTC18", current_cell
            current_cell.Value = "-"
        '
    End If
    
End Function
'================================================================================================================================
Function getQkPrincArray(ByVal norma As String, ByVal numPsi As Variant, ByVal stateLimit As String, ByVal blockName As String, ByVal tot As Variant, ByRef rangeNum As Range, ByRef rangeCorr As Range, ByRef rangeQinput As Range, ByRef rangeCondition As Range, ByRef rangeAnalysis As Range, ByRef rangeCategory As Range, ByRef rangeState As Range, Optional QkOrCatOrNum As String = "getQkPrinc") As Variant

    Dim j As Long
    Dim Qd As Double
    Dim dictQkPrinc As Object, dictQkPrincCategory As Object
    Set dictQkPrinc = CreateObject("Scripting.Dictionary")
    Set dictQkPrincCategory = CreateObject("Scripting.Dictionary")
    Set dictQkPrincNum = CreateObject("Scripting.Dictionary")

    j = 1
    For i = 1 To tot

        state = rangeState(i, 1).Value
        If state = "Attivo" Then

            num = rangeNum(i, 1).Value
            correlation = rangeCorr(i, 1).Value
            Qinput = IIf(IsEmpty(rangeQinput(i, 1).Value) Or Not IsNumeric(rangeQinput(i, 1).Value), 0, rangeQinput(i, 1).Value)
            condition = rangeCondition(i, 1).Value
            analysis = rangeAnalysis(i, 1).Value
            category = rangeCategory(i, 1).Value

            Qd = Qinput * getGamma(norma, stateLimit, blockName, condition, analysis) * getPsi(norma, stateLimit, numPsi, category)
            ' Debug.Print "Qd = Qinput * gamma * Psi = " & Qinput & " * " & getGamma(norma, stateLimit, blockName, condition, analysis) & " * " getPsi(norma, stateLimit, numPsi, category) & " = " & Qd

            If IsEmpty(correlation) Then
                dictQkPrinc.Add "CorrVoid" & CStr(j), Qd
                dictQkPrincCategory.Add "CorrVoid" & CStr(j), category
                dictQkPrincNum.Add "CorrVoid" & CStr(j), num
                j = j + 1
            ElseIf Not dictQkPrinc.Exists(correlation) Then
                dictQkPrinc.Add correlation, Qd
                dictQkPrincCategory.Add correlation, category
                dictQkPrincNum.Add correlation, num
            Else
                dictQkPrinc(correlation) = dictQkPrinc(correlation) + Qd
                dictQkPrincCategory(correlation) = dictQkPrincCategory(correlation) & ", " & category
                dictQkPrincNum(correlation) = dictQkPrincNum(correlation) & ", " & num
            End If

        End If
    Next

    If QkOrCatOrNum = "getQkPrinc" Then
        getQkPrincArray = IIf(dictQkPrinc.Count = 0, Array(0), dictQkPrinc.Items)
    ElseIf QkOrCatOrNum = "getQkPrincCategory" Then
        getQkPrincArray = IIf(dictQkPrincCategory.Count = 0, Array("-"), dictQkPrincCategory.Items)
    ElseIf QkOrCatOrNum = "getQkPrincNum" Then
        getQkPrincArray = IIf(dictQkPrincNum.Count = 0, Array("-"), dictQkPrincNum.Items)
    End If

End Function
'================================================================================================================================
Function getQkSeconArray(ByVal norma As String, ByVal numPsi As Variant, ByVal stateLimit As String, ByVal blockName As String, ByVal tot As Variant, ByRef rangeCorr As Range, ByRef rangeQinput As Range, ByRef rangeCondition As Range, ByRef rangeAnalysis As Range, ByRef rangeCategory As Range, ByRef rangeState As Range) As Variant

    Dim j As Long
    Dim Qd As Double
    Dim dictQkSecon As Object
    Set dictQkSecon = CreateObject("Scripting.Dictionary")

    j = 1
    For i = 1 To tot

        state = rangeState(i, 1).Value
        If state = "Attivo" Then

            correlation = rangeCorr(i, 1).Value
            Qinput = IIf(IsEmpty(rangeQinput(i, 1).Value) Or Not IsNumeric(rangeQinput(i, 1).Value), 0, rangeQinput(i, 1).Value)
            condition = rangeCondition(i, 1).Value
            analysis = rangeAnalysis(i, 1).Value
            category = rangeCategory(i, 1).Value

            Qd = Qinput * getGamma(norma, stateLimit, blockName, condition, analysis) * getPsi(norma, stateLimit, numPsi, category)
            ' Debug.Print "Qd = Qinput * gamma * Psi = " & Qinput & " * " & getGamma(norma, stateLimit, blockName, condition, analysis) & " * " getPsi(norma, stateLimit, numPsi, category) & " = " & Qd

            If IsEmpty(correlation) Then
                dictQkSecon.Add "CorrVoid" & CStr(j), Qd
                j = j + 1
            ElseIf Not dictQkSecon.Exists(correlation) Then
                dictQkSecon.Add correlation, Qd
            Else
                dictQkSecon(correlation) = dictQkSecon(correlation) + Qd
            End If

        End If
    Next

    getQkSeconArray = IIf(dictQkSecon.Count = 0, Array(0), dictQkSecon.Items)

End Function
