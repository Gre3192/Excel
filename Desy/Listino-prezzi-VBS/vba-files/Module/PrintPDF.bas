Attribute VB_Name = "PrintPDF"
Sub ExportPDF()

    Dim ws As Worksheet
    Dim rngToExport As Range
    Dim pdfPath As String
    Dim originalSettings As Variant
    
    ' Imposta il foglio di lavoro
    Set ws = ThisWorkbook.Worksheets("Listino prezzi")
    
    Dim lastRow As Object
    Set lastRow = ws.CustomProperties.Item(1)
    
    ' Definisci il range da esportare
    Set rngToExport = ws.Range("A1:P" & CStr(lastRow.Value))

    Dim fileNameWithoutExtension As String
    fileNameWithoutExtension = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    
    ' Salva le impostazioni di stampa originali per ripristinarle in seguito
    With ws.PageSetup
        originalSettings = Array(.Zoom, .FitToPagesWide, .FitToPagesTall, .CenterFooter)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Permette al foglio di espandersi su più pagine
        .CenterFooter = "Pagina &P di &N" ' Aggiunge la numerazione delle pagine nel piè di pagina
    End With

    ' Chiedi all'utente dove salvare il PDF
    pdfPath = Application.GetSaveAsFilename( _
        InitialFileName:=fileNameWithoutExtension, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Salva la tabella come")

    ' Controlla se l'utente ha premuto "Annulla"
    If pdfPath = "Falso" Then
        With ws.PageSetup
            .Zoom = originalSettings(0)
            .FitToPagesWide = originalSettings(1)
            .FitToPagesTall = originalSettings(2)
            .CenterFooter = originalSettings(3)
        End With
        Exit Sub
    End If

    ' Controlla se il file esiste e chiedi all'utente se desidera sovrascriverlo
    If Dir(pdfPath) <> "" Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Il file " & pdfPath & " esiste già. Vuoi sovrascriverlo?", vbYesNo + vbExclamation, "Conferma Sovrascrittura")
        If response = vbNo Then
            With ws.PageSetup
                .Zoom = originalSettings(0)
                .FitToPagesWide = originalSettings(1)
                .FitToPagesTall = originalSettings(2)
                .CenterFooter = originalSettings(3)
            End With
            Exit Sub
        End If
    End If

    ' Esporta il range in PDF
    On Error GoTo PDFExportError
    rngToExport.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Tabella esportata con successo in PDF!", vbInformation
    With ws.PageSetup
        .Zoom = originalSettings(0)
        .FitToPagesWide = originalSettings(1)
        .FitToPagesTall = originalSettings(2)
        .CenterFooter = originalSettings(3)
    End With
    Exit Sub

PDFExportError:
    MsgBox "Si è verificato un errore durante l'esportazione in PDF. Assicurati che il file PDF non sia aperto e riprova.", vbCritical, "Errore Export PDF"
    With ws.PageSetup
        .Zoom = originalSettings(0)
        .FitToPagesWide = originalSettings(1)
        .FitToPagesTall = originalSettings(2)
        .CenterFooter = originalSettings(3)
    End With
    
End Sub



