'namespace=vba-files\Moduli
Attribute VB_Name = "Calcola_SLU_SLE"

Sub calcola_SLUSLE()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Dim qk_princ_ntc08() As Variant, qk_princ_ntc18() As Variant, array_correlation() As Variant
    Dim qk_secon_ntc08() As Variant, qk_secon_ntc18() As Variant, array_string_correlation As Variant
    Dim qkd() As Variant, matrix_qk_princ() As Variant, matrix_qk_secon() As Variant
    Dim isThereG1 As Boolean, isThereG2 As Boolean, isThereQk As Boolean
    Dim start_row As Long, start_col As Long, next_row As Long
    Dim g1 As Double, qsl_g1 As Double, g2 As Double, qsl_g2 As Double, gam As Double, psi_princ As Double, psi_secon As Double
    Dim psi_princ_ntc18 As Double, psi_secon_ntc18 As Double, psi_princ_ntc08 As Double, psi_secon_ntc08 As Double
    Dim cell_num As Range, cell_tot As Range, cell_input As Range, cell_condition As Range, cell_analysis As Range
    Dim button_clicked As String, state_limit_selected As String, num_psi_princ As String, num_psi_secon As String
    Dim sum_qk_secon_ntc08 As Double, sum_qk_secon_ntc18 As Double

    Set ws = Application.ThisWorkbook.ActiveSheet
    button_clicked = Application.caller

    
    If button_clicked = "Calcola SLU" Then
        state_limit_selected = "SLU"
        num_psi_princ = "Not"
        num_psi_secon = "0"
    Elseif button_clicked = "Calcola SLE RARA" Then
        state_limit_selected = "SLE RARA"
        num_psi_princ = "Not"
        num_psi_secon = "0"
    Elseif button_clicked = "Calcola SLE FREQUENTE" Then
        state_limit_selected = "SLE FREQUENTE"
        num_psi_princ = "1"
        num_psi_secon = "2"
    Elseif button_clicked = "Calcola SLE Q.P." Then
        state_limit_selected = "SLE Q.P."
        num_psi_princ = "2"
        num_psi_secon = "2"
    Elseif button_clicked = "Calcola SISMICA" Then
        state_limit_selected = "SISMICA"
        num_psi_princ = "Not"
        num_psi_secon = "0"
    End If


    '-- CARICO G1 ----------------------------------------------------------------------------------------
        qsl_g1 = 0
        isThereG1 = True
        start_row = ws.Range(range_pointer(getBlockName(button_clicked))).row
        start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column

        Set cell_tot = ws.Cells(start_row + 1, start_col)

        If cell_tot.Value = "-" Then
            isThereG1 = False
            Goto jump_g1
        End If

        For i = 1 To cell_tot.Value
            next_row = start_row + 3 + i
            Set cell_input = ws.Cells(next_row, start_col + 1)
            Set cell_condition = ws.Cells(next_row, start_col + 2)
            Set cell_analysis = ws.Cells(next_row, start_col + 4)
            If IsEmpty(cell_input.Value) Or Not IsNumeric(cell_input.Value) Then
                message_box = True
                g1 = 0
                gam = 0
            Else
                g1 = cell_input.Value
                gam = gamma(state_limit_selected, "G1", cell_condition.Value, cell_analysis.Value)
            End If
            qsl_g1 = qsl_g1 + g1 * gam
            ' Debug.Print "(" & i & ") Qsl_G1 = G1 * gamma = " & g1 & " * " & gam & " = " & g1 * gam
        Next 
        ' Debug.Print "SOMMA Qsl_G1 = " & qsl_g1
        jump_g1:
    '
    '-- CARICO G2 ----------------------------------------------------------------------------------------       
        qsl_g2 = 0
        isThereG2 = True
        start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
        start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column
        
        Set cell_tot = ws.Cells(start_row + 1, start_col)
        
        If cell_tot.Value = "-" Then
            isThereG2 = False
            Goto jump_g2
        End If

        For i = 1 To cell_tot.Value
            next_row = start_row + 3 + i
            Set cell_input = ws.Cells(next_row, start_col + 1)
            Set cell_condition = ws.Cells(next_row, start_col + 2)
            Set cell_analysis = ws.Cells(next_row, start_col + 4)
            If IsEmpty(cell_input.Value) Or Not IsNumeric(cell_input.Value) Then
                message_box = True
                g2 = 0
                gam = 0
            Else
                g2 = cell_input.Value
                gam = gamma(state_limit_selected, "G2", cell_condition.Value, cell_analysis.Value)
            End If
            qsl_g2 = qsl_g2 + g2 * gam
            ' Debug.Print "(" & i & ") Qsl_G2 = G2 * gamma = " & g2 & " * " & gam & " = " & g2 * gam
        Next
        ' Debug.Print "SOMMA Qsl_G2 = " & qsl_g2
        jump_g2:
    '
    '-- CARICO Qk ----------------------------------------------------------------------------------------
        start_row = ws.Range(range_pointer(getBlockName(button_clicked))).Row
        start_col = ws.Range(range_pointer(getBlockName(button_clicked))).Column
        
        Set cell_tot = ws.Cells(start_row + 1, start_col)
        ReDim array_correlation(cell_tot.Value)
        ReDim array_string_correlation(cell_tot.Value)
        ReDim qk_princ_ntc08(cell_tot.Value)
        ReDim qk_princ_ntc18(cell_tot.Value)
        ReDim qk_secon_ntc08(cell_tot.Value)
        ReDim qk_secon_ntc18(cell_tot.Value)
        
        isThereQk = True
        If cell_tot.Value = "-" Then
            isThereQk = False
            Goto jump_qk
        End If

        '... Array Correlazioni ...................................................................................
        For i = 1 To cell_tot.Value
            next_row = start_row + 3 + i
            Set cell_correlation = ws.Cells(next_row, start_col + 2)
            If IsEmpty(cell_correlation.Value) Then
                array_correlation(i - 1) = "Empty Correlation" & CStr(i)
            Else
                array_correlation(i - 1) = cell_correlation.Value
            End If
        Next
        '..........................................................................................................
        For i = 1 To cell_tot.Value
            next_row = start_row + 3 + i
            Set cell_num = ws.Cells(next_row, start_col)
            Set cell_input = ws.Cells(next_row, start_col + 1)
            Set cell_correlation = ws.Cells(next_row, start_col + 2)
            Set cell_condition = ws.Cells(next_row, start_col + 4)
            Set cell_analysis = ws.Cells(next_row, start_col + 6)
            Set cell_category = ws.Cells(next_row, start_col + 7)
            gam = gamma(state_limit_selected, "Qk", cell_condition.Value, cell_analysis.Value)
            psi_princ_ntc18 = psi("NTC18", state_limit_selected, num_psi_princ, cell_category.Value)
            psi_secon_ntc18 = psi("NTC18", state_limit_selected, num_psi_secon, cell_category.Value)
            psi_princ_ntc08 = psi("NTC08", state_limit_selected, num_psi_princ, cell_category.Value)
            psi_secon_ntc08 = psi("NTC08", state_limit_selected, num_psi_secon, cell_category.Value)
            '... Carichi principali e secondari .......................................................................
            For j = 1 To cell_tot.Value
                If array_correlation(j - 1) = "Empty Correlation" & CStr(i) Then
                    qk_princ_ntc08(j - 1) = cell_input.Value * gam * psi_princ_ntc08
                    qk_princ_ntc18(j - 1) = cell_input.Value * gam * psi_princ_ntc18
                    qk_secon_ntc08(j - 1) = cell_input.Value * gam * psi_secon_ntc08
                    qk_secon_ntc18(j - 1) = cell_input.Value * gam * psi_secon_ntc18
                    array_string_correlation(j - 1) = CStr(cell_num.Value)
                Elseif array_correlation(j - 1) = cell_correlation.Value Then
                    qk_princ_ntc08(j - 1) = qk_princ_ntc08(j - 1) + cell_input.Value * gam * psi_princ_ntc08
                    qk_princ_ntc18(j - 1) = qk_princ_ntc18(j - 1) + cell_input.Value * gam * psi_princ_ntc18
                    qk_secon_ntc08(j - 1) = qk_secon_ntc08(j - 1) + cell_input.Value * gam * psi_secon_ntc08
                    qk_secon_ntc18(j - 1) = qk_secon_ntc18(j - 1) + cell_input.Value * gam * psi_secon_ntc18
                    array_string_correlation(j - 1) = array_string_correlation(j - 1) & IIf(IsEmpty(array_string_correlation(j - 1)), "", ",") & CStr(cell_num.Value)
                    Exit For
                End If
            Next
            sum_qk_secon_ntc08 = Application.WorksheetFunction.Sum(qk_secon_ntc08)
            sum_qk_secon_ntc18 = Application.WorksheetFunction.Sum(qk_secon_ntc18)
        Next

        ' Debug.Print "---------------------------"
        ' Debug.Print "ARRAY CORRELAZIONI"
        ' Debug.Print "---------------------------"
        ' For j = 1 To cell_tot.Value
        '     Debug.Print array_correlation(j - 1)
        ' Next
        ' Debug.Print "---------------------------"
        ' Debug.Print "ARRAY STRINGA CORRELAZIONI"
        ' Debug.Print "---------------------------"
        ' For j = 1 To cell_tot.Value
        '     Debug.Print array_string_correlation(j - 1)
        ' Next
        ' Debug.Print "---------------------------"
        ' Debug.Print "CARICHI PRINCIPALI"
        ' Debug.Print "---------------------------"
        ' For j = 1 To cell_tot.Value
        '     Debug.Print qk_princ_ntc18(j - 1)
        ' Next
        ' Debug.Print "---------------------------"
        ' Debug.Print "CARICHI SECONDARI"
        ' Debug.Print "---------------------------"
        ' For j = 1 To cell_tot.Value
        '     Debug.Print qk_secon_ntc18(j - 1)
        ' Next
        ' Debug.Print "..........................."
        ' Debug.Print "Sum sec:", sum_qk_secon_ntc18
        jump_qk:
    '

    
    If Not isThereG1 And Not isThereG2 And Not isThereQk Then
        Exit Sub
    End If








    ' '-- CARICHI STATO LIMITE --------------------------------------------------------------------------------
    '     For i = 1 To cell_tot.Value
            





    '     Next
     

    '     For i = 1 To cell_tot.Value
    '         next_row = start_row + 3 + i
    '         Set cell_input = ws.Cells(next_row, start_col + 1)


    '         psi_princ = psi(num_psi_princ, cell_category.Value)
    '         psi_secon = psi(num_psi_secon, cell_category.Value)

    '         If IsEmpty(cell_input.Value) Or Not IsNumeric(cell_input.Value) Then
    '             message_box = True
    '             qk(i) = 0
    '             gam = 0
    '         Else
    '             qk(i) = cell_input
    '             gam = gamma(state_limit_selected, "Qk", cell_condition.Value, cell_analysis.Value)
    '         End If
    '         qsl_qk(i) = qk(i) * gam * psi_0
    '     '    Debug.Print "(" & i & ") Qsl_qk = qk * gamma * psi_0 = " & qk(i) & " * " & gam & " * " & psi_0 & " = " & qsl_qk(i)
    '     Next
    '    sum_qsl_qk = Application.WorksheetFunction.Sum(qslu_qk)
    ' '    Debug.Print "SOMMA Qslu_qk = " & sum_qslu_qk
    ' jump_qk:







    
    ' '----------------------------------------------------------------------------------------------------------
    ' start_row = Range(start_cell).Row
    ' start_col = Range(start_cell).Column
    ' last_row = Cells(Rows.Count, start_col - 1).End(xlUp).Row
    ' '--- QSLU -------------------------------------------------------------------------------------------------
    ' qslu_G1G2 = qslu_g1 + qslu_g2
    ' For i = 0 To cell_tot - 1
    '     '--- COMBO --------------------------------------------------------------------------------------------
    '     Set cell_combo = Cells(last_row + i, start_col - 1)
    '     number_validation cell_combo
    '     cell_combo.HorizontalAlignment = xlCenter
    '     cell_combo.Value = i + 1

    '     '--- CARICO PRINCIPALE --------------------------------------------------------------------------------
    '     Set cell_main_load = Range(Cells(last_row + i, start_col), Cells(last_row + i, start_col + 2))
    '     cell_main_load.Merge
    '     number_validation cell_main_load
    '     cell_main_load.HorizontalAlignment = xlCenter
    '     Cells(last_row + i, start_col) = Cells(start_row_load + 2 + i, start_col_load + 4)

    '     '--- CARICO SLU --------------------------------------------------------------------------------
    '     Set cell_condition = Cells(start_row_load + 2 + i, start_col_load + 1)
    '     Set cell_state = Cells(start_row_load + 2 + i, start_col_load + 3)
    '     Set cell_load_slu = Range(Cells(last_row + i, start_col + 3), Cells(last_row + i, start_col + 4))
    '     cell_load_slu.Merge
    '     number_validation cell_load_slu
    '     cell_load_slu.HorizontalAlignment = xlCenter
    '     qslu = qslu_G1G2 + sum_qslu_qk - qslu_qk(i + 1) + qk(i + 1) * gamma("SLU", "Qk", cell_condition, cell_state)
    '     Cells(last_row + i, start_col + 3) = qslu
    ' '    Debug.Print "(" & i + 1 & ") CARICO PRINCIPALE TOLTO: Qslu_k = " & qslu_qk(i + 1)
    ' '    Debug.Print "(" & i + 1 & ") CARICO PRINCIPALE AGGIUNTO: qk * gamma = " & qk(i + 1) & " * " & gamma("SLU", "Qk", cell_condition, cell_state) & " = " & qk(i + 1) * gamma("SLU", "Qk", cell_condition, cell_state)
    ' Next
    ' If cell_tot = 0 Then
    '     '--- COMBO --------------------------------------------------------------------------------------------
    '     Set cell_combo = Cells(last_row, start_col - 1)
    '     number_validation cell_combo
    '     cell_combo.HorizontalAlignment = xlCenter
    '     cell_combo.Value = 1

    '     '--- CARICO SLU --------------------------------------------------------------------------------
    '     Cells(start_row + 4, start_col + 3) = qslu_G1G2
    ' End If

End Sub


