Attribute VB_Name = "WordRelation"
'namespace=vba-files\Moduli

Sub CreaRelazioneWord()
    ' Dichiarazione delle variabili
    Dim wrdApp As Object
    Dim wrdDoc As Object
    Dim rng As Range
    Dim ws As Worksheet
    Dim dialogSave As FileDialog
    
    ' Imposta il foglio di lavoro attivo
    Set ws = ActiveSheet
    
    ' Definisci il riferimento alla gamma di dati da includere nella relazione
    ' Modifica "A1:C10" in base alla tua gamma di dati effettiva
    Set rng = ws.Range("A1:C10")
    
    ' Crea un'applicazione Word
    On Error Resume Next
    ' Prova a ottenere un'istanza aperta di Word
    Set wrdApp = GetObject(, "Word.Application")
    On Error GoTo 0
    
    ' Se Word non è aperto, crea una nuova istanza di Word
    If wrdApp Is Nothing Then
        Set wrdApp = CreateObject("Word.Application")
    End If
    
    ' Crea un nuovo documento Word
    Set wrdDoc = wrdApp.Documents.Add
    
    ' Copia i dati dalla gamma di Excel nel documento Word
    rng.Copy
    wrdDoc.Range.PasteExcelTable False, False, False
    
    ' Chiudi l'Editor del Registro di Windows per evitare problemi
    wrdApp.Quit False
    
    ' Mostra una finestra di dialogo per selezionare il percorso e il nome del file
    Set dialogSave = Application.FileDialog(msoFileDialogSaveAs)
    With dialogSave
        .Title = "Salva il file Word"
        .InitialFileName = "Relazione"
        If .Show = -1 Then
            ' Salva il documento Word nel percorso specificato
            wrdDoc.SaveAs2 .SelectedItems(1), FileFormat:=wdFormatDocumentDefault
        End If
    End With
    
    ' Chiudi il documento Word
    wrdDoc.Close False
    
    ' Rilascia le variabili
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    Set dialogSave = Nothing
End Sub
