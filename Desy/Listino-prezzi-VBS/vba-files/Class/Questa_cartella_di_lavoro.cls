VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Questa_cartella_di_lavoro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Dim ws As Worksheet
    Dim fileNameWithoutExtension As String
    Dim inputCell As String
    
    Set ws = Application.ThisWorkbook.Sheets("Listino prezzi")
    
    inputCell = "B1"
    
    ' Rimuove l'estensione dal nome del file
    fileNameWithoutExtension = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    
    ' Inserisce il nome del file nella cella "A1" di ogni foglio
    ws.Range(inputCell).Value = fileNameWithoutExtension
    
    ' Imposta il libro come già salvato
    ThisWorkbook.Saved = True

End Sub

