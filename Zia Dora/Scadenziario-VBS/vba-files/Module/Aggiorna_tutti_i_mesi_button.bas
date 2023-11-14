Attribute VB_Name = "Aggiorna_tutti_i_mesi_button"

Sub Aggiorna_tutti_i_mesi()
    
    Application.ScreenUpdating = False
    ' Application.StatusBar = "Aggiornamento in tutti i fogli in corso..."

    Dim sh_riepilogo As Worksheet
    Set sh_riepilogo = ThisWorkbook.Worksheets("Uscite")

    ' Cancella formule dinamiche nei totali ciane
    formule_totali_ciane_riepilogo sh_riepilogo, "Delete"
        
    ' Cancella formule dinamiche nei totali fornitori
    formule_totali_fornitori_riepilogo sh_riepilogo, "Delete"


    ' ========================================================================================================

    Dim sh As Worksheet, sh_mese As Worksheet
    Set sh = ThisWorkbook.Worksheets("Elenco Ditte")

    Dim Array_mesi() As Variant
    Array_mesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")
    Array_mesi_riepilogo = Array("Settembre", "Ottobre", "Novembre", "Dicembre", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto")
    
    Dim incr As Integer, index_mese_colonna_ciane As Integer, index_mese_colonna_fornitori As Integer
    
    Dim lastrow_uscite As Integer, lastrow_fornitori As Integer, doAddedNum As Integer, doDeletedNum As Integer
    
    Dim Array_doAdded() As Variant
    Dim Array_doDeleted() As Variant

    Dim Array_ciane_tab() As Variant
    Dim Array_fornitori_tab() As Variant

    Dim lastrow_ciane_input As Integer, lastrow_fornitori_input As Integer
    lastrow_ciane_input = sh.CustomProperties.Item(1).Value
    lastrow_fornitori_input = sh.CustomProperties.Item(2).Value
    
    Dim tot_ciane_input As Integer, tot_ciane_tab As Integer
    tot_ciane_input = sh.Cells(lastrow_ciane_input, "A").Value
    
    Dim tot_fornitori_input As Integer, tot_fornitori_tab As Integer
    tot_fornitori_input = sh.Cells(lastrow_fornitori_input, "H").Value
    
    Dim Array_ciane_input() As Variant
    ReDim Array_ciane_input(0 To lastrow_ciane_input - 15)
    For i = 0 To lastrow_ciane_input - 15
        Array_ciane_input(i) = sh.Cells(16 + i, "B").Value
    Next
       
    Dim Array_fornitori_input() As Variant
    ReDim Array_fornitori_input(0 To lastrow_fornitori_input - 15)
    For i = 0 To lastrow_fornitori_input - 15
        Array_fornitori_input(i) = sh.Cells(16 + i, "I").Value
    Next
    




    
    ' ========================================================================================================
    
    For Each mese In Array_mesi

        Set sh_mese = ThisWorkbook.Worksheets(mese)
        lastRow_ciane_tab = sh_mese.CustomProperties.Item(1).Value

        ' Popola array ciane tab
        cianeTabNum = 0
        redim Array_ciane_tab(0 to 0)
        For i = 0 To lastRow_ciane_tab
            If Not isEmpty(nameCell) Then
                Array_ciane_tab(cianeTabNum) = sh_mese.Cells(18 + i, "A").Value
                if i <> Array_ciane_tab then 
                    redim Preserve Array_ciane_tab(0 To cianeTabNum + 1)
                    cianeTabNum = cianeTabNum + 1
                end if
            end if
        Next




        
        ' doAddedNum = 0
        ' doDeletedNum = 0
        ' redim Array_doAdded(0 to 0)
        ' redim Array_doDeleted(0 to 0)

        ' For i = 0 To lastRow_ciane_tab

        '     nameCell = sh_mese.Cells(i + 18,"A").Value

        '     If Not isEmpty(nameCell) Then

        '         For j = 0 to UBound(Array_ciane_input)

        '             If Array_ciane_input(j) = nameCell Then

        '                 Exit For

        '             ElseIf j = UBound(Array_ciane_input) then
                        
        '                 redim Preserve Array_doDeleted(0 To doDeletedNum + 1)
        '                 Array_doDeleted(doDeletedNum) = nameCell
        '                 doDeletedNum = doDeletedNum + 1

        '             End If

        '         Next



        '     End If

        ' Next

        ' debug.print "--------------------------------------- do Added"
        ' for i = 0 to UBound(Array_doAdded)
        '    debug.print Array_doAdded(i)
        ' next
        ' debug.print "--------------------------------------- do Deleted"
        ' for i = 0 to UBound(Array_doDeleted)
        '     debug.print Array_doDeleted(i)
        ' next




        
        ' ' Confronta e aggiorna ciane_input e ciane_tab
        ' tot_ciane_tab = sh_mese.Cells(14, "B").Value
        ' incr = 1
        ' If tot_ciane_tab > tot_ciane_input Then
        '     ReDim Preserve Array_ciane_input(tot_ciane_tab)
        '     incr = 0
        ' End If
        ' i = 0
        ' For el = 1 To UBound(Array_ciane_input) + incr
            
        '     If Array_ciane_input(el - 1) <> sh_mese.Cells(18 + i, "A").Value Then
        '         If IsEmpty(Array_ciane_input(el - 1)) Or Array_ciane_input(el - 1) = "" Then
        '             sh_mese.Range(sh_mese.Cells(18 + i, "A"), sh_mese.Cells(18 + 5 + i, "M")).Delete Shift:=xlUp
        '         Else
        '             formattazione_tabella_ciane sh_mese, Array_ciane_input(el - 1), i
        '             i = i + 6
        '         End If
        '     Else
        '         i = i + 6
        '     End If
            
        ' Next
        ' sh_mese.Cells(14, "B").Value = tot_ciane_input
        ' lastrow_uscite = sh_mese.Cells(Rows.Count, "A").End(xlUp).Row
        ' With sh_mese.Range(sh_mese.Cells(lastrow_uscite + 5, "A"), sh_mese.Cells(lastrow_uscite + 5, "M")).Borders(xlEdgeBottom)
        '     .LineStyle = xlContinuous
        '     .Color = vbBlack
        '     .Weight = xlThin
        ' End With
        
        
        ' ' Confronta e aggiorna fornitori_elenco e fornitori_mese
        ' tot_fornitori_tab = sh_mese.Cells(14, "P").Value
        ' incr = 1
        ' If tot_fornitori_tab > tot_fornitori_input Then
        '     ReDim Preserve Array_fornitori_input(tot_fornitori_tab)
        '     incr = 0
        ' End If
        ' i = 0
        ' For el = 1 To UBound(Array_fornitori_input) + incr
            
        '     If Array_fornitori_input(el - 1) <> sh_mese.Cells(18 + i, "O").Value Then
        '         If IsEmpty(Array_fornitori_input(el - 1)) Or Array_fornitori_input(el - 1) = "" Then
        '             sh_mese.Range(sh_mese.Cells(18 + i, "O"), sh_mese.Cells(18 + 5 + i, "AA")).Delete Shift:=xlUp
        '         Else
        '             formattazione_tabella_fornitori sh_mese, Array_fornitori_input(el - 1), i
        '             i = i + 6
        '         End If
        '     Else
        '         i = i + 6
        '     End If
            
        ' Next
        ' sh_mese.Cells(14, "P").Value = tot_fornitori_input
        ' lastrow_fornitori = sh_mese.Cells(Rows.Count, "O").End(xlUp).Row
        ' With sh_mese.Range(sh_mese.Cells(lastrow_fornitori + 5, "O"), sh_mese.Cells(lastrow_fornitori + 5, "AA")).Borders(xlEdgeBottom)
        '     .LineStyle = xlContinuous
        '     .Color = vbBlack
        '     .Weight = xlThin
        ' End With
    
    Next mese
    
    

    ' ========================================================================================================
    
    

    ' For Each mese In Array_mesi_riepilogo
    
    '     Set sh_mese = ThisWorkbook.Worksheets(mese)
    
    '     ' Confronta e aggiorna riepilogo uscite
    '     i = 0
    '     For el = 1 To UBound(Array_ciane_input) + incr
            
    '         If Array_ciane_input(el - 1) <> sh_riepilogo.Cells(18 + i, "A").Value Then
                    
    '             If IsEmpty(Array_ciane_input(el - 1)) Or Array_ciane_input(el - 1) = "" Then
    '                 sh_riepilogo.Range(sh_riepilogo.Cells(18 + i, "A"), sh_riepilogo.Cells(18 + i, "AB")).Delete Shift:=xlUp
    '             Else
    '                 formattazione_ciane_riepilogo sh_mese, sh_riepilogo, Array_ciane_input(el - 1), i
    '                 i = i + 1
    '             End If
    '         Else
    '             i = i + 1
    '         End If
            
    '     Next
        
        
    '     ' Confronta e aggiorna riepilogo fornitori
    '     i = 0
    '     For el = 1 To UBound(Array_fornitori_input) + incr
            
    '         If Array_fornitori_input(el - 1) <> sh_riepilogo.Cells(18 + i, "AD").Value Then
         
    '             If IsEmpty(Array_fornitori_input(el - 1)) Or Array_fornitori_input(el - 1) = "" Then
    '                 sh_riepilogo.Range(sh_riepilogo.Cells(18 + i, "AD"), sh_riepilogo.Cells(18 + 5 + i, "BE")).Delete Shift:=xlUp
    '             Else
    '                 formattazione_fornitori_riepilogo sh_mese, sh_riepilogo, Array_fornitori_input(el - 1), i
    '                 i = i + 1
    '             End If
    '         Else
    '             i = i + 1
    '         End If
            
    '     Next
              
    ' Next mese
    
    
    ' ' Aggiorna formule dinamiche totali
    ' formule_totali_ciane_riepilogo sh_riepilogo
    ' formule_totali_fornitori_riepilogo sh_riepilogo
    
    ' ' Aggiorna formule dinamiche
    ' index_mese_colonna_ciane = 4
    ' index_mese_colonna_fornitori = 33
    ' For Each mese In Array_mesi_riepilogo
    '     Set sh_mese = ThisWorkbook.Worksheets(mese)
    '     formule_ciane_riepilogo sh_riepilogo, sh_mese, tot_ciane_input, index_mese_colonna_ciane
    '     formule_fornitori_riepilogo sh_riepilogo, sh_mese, tot_fornitori_input, index_mese_colonna_fornitori
    '     index_mese_colonna_ciane = index_mese_colonna_ciane + 2
    '     index_mese_colonna_fornitori = index_mese_colonna_fornitori + 2
    ' Next mese
    
    ' Application.StatusBar = False
    
End Sub
