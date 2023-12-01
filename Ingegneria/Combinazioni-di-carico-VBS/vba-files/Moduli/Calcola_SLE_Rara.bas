Attribute VB_Name = "Calcola_SLE_Rara"
'namespace=vba-files\Moduli

Sub calcola_combinazione_SLE_Rara()

    Application.ScreenUpdating = False

    reset "Resetta SLE RARA"

    Dim ws As Worksheet
    Set ws = Application.ThisWorkbook.ActiveSheet

    Dim tot As Variant, Qinput As Variant
    Dim condition As String, analysis As String, state As String
    Dim isThereG1 As Boolean, isThereG2 As Boolean, isThereQk As Boolean
    Dim start_row As Long, start_col As Long, next_row As Long, num As Long
    Dim g1 As Double, tot_qsl_g1_NTC08 As Double, tot_qsl_g1_NTC18 As Double, g2 As Double, tot_qsl_g2_NTC08 As Double, tot_qsl_g2_NTC18 As Double
    Dim rangeNum As Range, rangeCorr As Range, rangeQinput As Range, rangeCondition As Range, rangeAnalysis As Range, rangeCategory As Range, rangeState As Range
    Dim qsl_qkPrinc_NTC08() As Variant, qsl_qkPrinc_NTC18() As Variant, qsl_qkSecon_NTC08() As Variant, qsl_qkSecon_NTC18() As Variant, qkPrinc_Category() As Variant, qkPrinc_Num() As Variant

    '-- CARICO G1 ----------------------------------------------------------------------------------------
        start_row = ws.Range(range_pointer("G1")).Row
        start_col = ws.Range(range_pointer("G1")).Column

        tot_qsl_g1_NTC08 = 0
        tot_qsl_g1_NTC18 = 0
        tot = ws.Cells(start_row + 1, start_col).Value
        If tot = "-" Then
            isThereG1 = False
        Else
            isThereG1 = True
            For i = 1 To tot
                next_row = start_row + 3 + i
                Qinput = ws.Cells(next_row, start_col + 4).Value
                state = ws.Cells(next_row, start_col + 9).Value
                g1 = IIf(IsEmpty(Qinput) Or Not IsNumeric(Qinput) Or state <> "Attivo", 0, Qinput)
                tot_qsl_g1_NTC08 = tot_qsl_g1_NTC08 + g1
                tot_qsl_g1_NTC18 = tot_qsl_g1_NTC18 + g1
                ' Debug.Print "(NTC08) (" & i & ") QsleG1 = G1 = " & g1
                ' Debug.Print "(NTC18) (" & i & ") QsleG1 = G1 = " & g1
            Next
        End If
        ' Debug.Print "(NTC08) SOMMA Qsle_G1 = " & tot_qsl_g1_NTC08 & vbCrLf & "--------------------------------------------"
        ' Debug.Print "(NTC18) SOMMA Qsle_G1 = " & tot_qsl_g1_NTC18 & vbCrLf & "--------------------------------------------"
    '
    '-- CARICO G2 ----------------------------------------------------------------------------------------
        start_row = ws.Range(range_pointer("G2")).Row
        start_col = ws.Range(range_pointer("G2")).Column
        
        tot_qsl_g2_NTC08 = 0
        tot_qsl_g2_NTC18 = 0
        tot = ws.Cells(start_row + 1, start_col).Value
        If tot = "-" Then
            isThereG2 = False
        Else
            isThereG2 = True
            For i = 1 To tot
                next_row = start_row + 3 + i
                Qinput = ws.Cells(next_row, start_col + 4)
                state = ws.Cells(next_row, start_col + 9).Value
                g2 = IIf(IsEmpty(Qinput) Or Not IsNumeric(Qinput) Or state <> "Attivo", 0, Qinput)
                tot_qsl_g2_NTC08 = tot_qsl_g2_NTC08 + g2
                tot_qsl_g2_NTC18 = tot_qsl_g2_NTC18 + g2
                ' Debug.Print "(NTC08) (" & i & ") QsleG2 = G2 = " & g2
                ' Debug.Print "(NTC18) (" & i & ") QsleG2 = G2 = " & g2
            Next
        End If
        ' Debug.Print "(NTC08) SOMMA Qsle_G2 = " & tot_qsl_g2_NTC08 & vbCrLf & "--------------------------------------------"
        ' Debug.Print "(NTC18) SOMMA Qsle_G2 = " & tot_qsl_g2_NTC18 & vbCrLf & "--------------------------------------------"
    '
    '-- CARICO Qk ----------------------------------------------------------------------------------------
        start_row = ws.Range(range_pointer("Qk")).Row
        start_col = ws.Range(range_pointer("Qk")).Column

        qkPrinc_Num = Array("-")
        qkPrinc_Category = Array("-")
        qsl_qkPrinc_NTC08 = Array(0)
        qsl_qkPrinc_NTC18 = Array(0)
        qsl_qkSecon_NTC08 = Array(0)
        qsl_qkSecon_NTC18 = Array(0)
        tot = ws.Cells(start_row + 1, start_col).Value
        If tot = "-" Then
            isThereQk = False
        Else
            isThereQk = True

            Set rangeNum = ws.Range(ws.Cells(start_row + 4, start_col), ws.Cells(start_row + 3 + tot, start_col))
            Set rangeCorr = ws.Range(ws.Cells(start_row + 4, start_col + 4), ws.Cells(start_row + 3 + tot, start_col + 4))
            Set rangeQinput = ws.Range(ws.Cells(start_row + 4, start_col + 6), ws.Cells(start_row + 3 + tot, start_col + 6))
            Set rangeCondition = ws.Range(ws.Cells(start_row + 4, start_col + 8), ws.Cells(start_row + 3 + tot, start_col + 8))
            Set rangeAnalysis = ws.Range(ws.Cells(start_row + 4, start_col + 10), ws.Cells(start_row + 3 + tot, start_col + 10))
            Set rangeCategory = ws.Range(ws.Cells(start_row + 4, start_col + 11), ws.Cells(start_row + 3 + tot, start_col + 11))
            Set rangeState = ws.Range(ws.Cells(start_row + 4, start_col + 14), ws.Cells(start_row + 3 + tot, start_col + 14))

            qkPrinc_Num = getQkPrincArray("NTC18", "NotNum", "SLE RARA", "Qk", tot, rangeNum, rangeCorr, rangeQinput, rangeCondition, rangeAnalysis, rangeCategory, rangeState, "getQkPrincNum")
            qkPrinc_Category = getQkPrincArray("NTC18", "NotNum", "SLE RARA", "Qk", tot, rangeNum, rangeCorr, rangeQinput, rangeCondition, rangeAnalysis, rangeCategory, rangeState, "getQkPrincCategory")
            qsl_qkPrinc_NTC08 = getQkPrincArray("NTC08", "NotNum", "SLE RARA", "Qk", tot, rangeNum, rangeCorr, rangeQinput, rangeCondition, rangeAnalysis, rangeCategory, rangeState)
            qsl_qkPrinc_NTC18 = getQkPrincArray("NTC18", "NotNum", "SLE RARA", "Qk", tot, rangeNum, rangeCorr, rangeQinput, rangeCondition, rangeAnalysis, rangeCategory, rangeState)
            
            qsl_qkSecon_NTC08 = getQkSeconArray("NTC08", 0, "SLE RARA", "Qk", tot, rangeCorr, rangeQinput, rangeCondition, rangeAnalysis, rangeCategory, rangeState)
            qsl_qkSecon_NTC18 = getQkSeconArray("NTC18", 0, "SLE RARA", "Qk", tot, rangeCorr, rangeQinput, rangeCondition, rangeAnalysis, rangeCategory, rangeState)
            
            ' for i = 0 to Ubound(qsl_qkPrinc_NTC18)
            '     debug.print "Carichi principali: " & qkPrinc_Num(i), qkPrinc_Category(i), qsl_qkPrinc_NTC18(i)
            ' next
            ' for i = 0 to Ubound(qsl_qkSecon_NTC18)
            '     debug.print qsl_qkSecon_NTC18(i)
            ' next
            ' Debug.print "----------------------------------------"
            ' for i = 0 to Ubound(qsl_qkPrinc_NTC08)
            '     debug.print "Carichi principali: " & qkPrinc_Num(i), qkPrinc_Category(i), qsl_qkPrinc_NTC08(i)
            ' next
            ' for i = 0 to Ubound(qsl_qkSecon_NTC08)
            '     debug.print qsl_qkSecon_NTC08(i)
            ' next
            ' Debug.print "----------------------------------------"

        End If
    '
    '-- COMBINAZIONE SLE RARA ----------------------------------------------------------------------------
        start_row = ws.Range(range_pointer("SLE RARA")).Row
        start_col = ws.Range(range_pointer("SLE RARA")).Column
        
        For i = 0 To UBound(qsl_qkPrinc_NTC18)

            cells_style "Combo", ws.Cells(start_row + 4 + i, start_col)
            cells_style "Carichi variabili principali", ws.Range(Cells(start_row + 4 + i, start_col + 1), Cells(start_row + 4 + i, start_col + 3))
            cells_style "Annesse categorie principali", ws.Range(Cells(start_row + 4 + i, start_col + 4), Cells(start_row + 4 + i, start_col + 6))
            cells_style "q - NTC08", ws.Range(Cells(start_row + 4 + i, start_col + 7), Cells(start_row + 4 + i, start_col + 8))
            cells_style "q - NTC18", ws.Range(Cells(start_row + 4 + i, start_col + 9), Cells(start_row + 4 + i, start_col + 10))

            ws.Cells(start_row + 4 + i, start_col).Value = i + 1
            ws.Cells(start_row + 4 + i, start_col + 1).Value = qkPrinc_Num(i)
            ws.Cells(start_row + 4 + i, start_col + 4).Value = qkPrinc_Category(i)
            ws.Cells(start_row + 4 + i, start_col + 7).Value = (tot_qsl_g1_NTC08 + tot_qsl_g2_NTC08 + qsl_qkPrinc_NTC08(i) + Application.WorksheetFunction.Sum(qsl_qkSecon_NTC08) - qsl_qkSecon_NTC08(i)) * udm
            ws.Cells(start_row + 4 + i, start_col + 9).Value = (tot_qsl_g1_NTC18 + tot_qsl_g2_NTC18 + qsl_qkPrinc_NTC18(i) + Application.WorksheetFunction.Sum(qsl_qkSecon_NTC18) - qsl_qkSecon_NTC18(i)) * udm

            If i <> 0 And ws.Cells(start_row + 4 + i, start_col + 7).Value > ws.Cells(start_row + 4 + i - 1, start_col + 7).Value Then
                maxRowNTC08 = i
            End If
            If i <> 0 And ws.Cells(start_row + 4 + i, start_col + 9).Value > ws.Cells(start_row + 4 + i - 1, start_col + 9).Value Then
                maxRowNTC18 = i
            End If
        Next

        If UBound(qsl_qkPrinc_NTC18) <> 0 Then
            cells_style "MaxColor", ws.Range(Cells(start_row + 4 + maxRowNTC08, start_col + 7), Cells(start_row + 4 + maxRowNTC08, start_col + 7))
            cells_style "MaxColor", ws.Range(Cells(start_row + 4 + maxRowNTC18, start_col + 9), Cells(start_row + 4 + maxRowNTC18, start_col + 9))
        End If
    '
    '-- CELLA TOT ----------------------------------------------------------------------------------------
        ws.Cells(start_row + 1, start_col).Value = UBound(qsl_qkPrinc_NTC18) + 1
    '
End Sub
