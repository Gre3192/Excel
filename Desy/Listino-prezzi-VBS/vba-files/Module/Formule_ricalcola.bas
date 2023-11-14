Attribute VB_Name = "Formule_ricalcola"
Function ricalcola_formule()

    Application.ScreenUpdating = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Listino prezzi")
  
    Dim lastRow As Object
    Set lastRow = sh.CustomProperties.Item(1)

    Dim formulaString As String
    
    formulaString = ""
    For i = 11 To lastRow.Value
'        sh.Cells(i, "K").Select
        If sh.Cells(i, "K").Font.Size = 14 And sh.Cells(i + 1, "K").Font.Size = 14 Then
            thisaddress = sh.Cells(i, "K").Address
            i = i + 1
        Else
            formulaString = formulaString & sh.Cells(i, "K").Address & ","
        End If

        If sh.Cells(i, "K").Font.Size = 14 And i + 1 >= CLng(lastRow.Value) Then
            Exit For
        End If

        If (sh.Cells(i + 1, "K").Font.Size = 14 And sh.Cells(i + 2, "K").Font.Size = 14 And formulaString <> "") Or i = CLng(lastRow.Value) Then
            If Not IsEmpty(thisaddress) Then
                sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
            End If
            formulaString = ""
        End If

    Next


    formulaString = ""
    For i = 11 To lastRow.Value

        If sh.Cells(i, "O").Font.Size = 14 And sh.Cells(i + 1, "O").Font.Size = 14 Then
            thisaddress = sh.Cells(i, "O").Address
            i = i + 1
        Else
            formulaString = formulaString & sh.Cells(i, "O").Address & ","
        End If

        If sh.Cells(i, "O").Font.Size = 14 And i + 1 >= CLng(lastRow.Value) Then
            Exit For
        End If

        If (sh.Cells(i + 1, "O").Font.Size = 14 And sh.Cells(i + 2, "O").Font.Size = 14 And formulaString <> "") Or i = CLng(lastRow.Value) Then
            If Not IsEmpty(thisaddress) Then
                sh.Range(thisaddress).Formula = "=SUM(" & formulaString & ")"
            End If
            formulaString = ""
        End If

    Next


    formulaString = ""
    For i = 11 To lastRow.Value
        If Cells(i, "K").Font.Size = 14 Then
            formulaString = formulaString & Cells(i, "K").Address & ","
            i = i + 1
        End If
    Next
    If formulaString <> "" Then
        sh.Range("A7").Formula = "=SUM(" & formulaString & ")"
    Else
        sh.Range("A7").Value = 0
    End If


    formulaString = ""
    For i = 11 To lastRow.Value
        If Cells(i, "O").Font.Size = 14 Then
            formulaString = formulaString & Cells(i, "O").Address & ","
            i = i + 1
        End If
    Next
    If formulaString <> "" Then
        sh.Range("G7").Formula = "=SUM(" & formulaString & ")"
    Else
        sh.Range("G7").Value = 0
    End If



End Function
