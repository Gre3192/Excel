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

  Application.ScreenUpdating = False

  Dim ws As Worksheet
  Set ws = Application.ThisWorkbook.Sheets("Combinazioni")

  Dim fileNameWithoutExtension As String
  fileNameWithoutExtension = UCase(Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1))

  With ws.Range("A1:EI2")
    .Value = fileNameWithoutExtension
    .Font.Bold = True
  End With
  
  ThisWorkbook.Saved = True

End Sub
'================================================================================================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

  Application.ScreenUpdating = False

  Dim ws As Worksheet
  Set ws = Application.ThisWorkbook.ActiveSheet
  Dim inputCells1 As Range, inputCells2 As Range, inputCells3 As Range, outputCells1 As Range, outputCells2 As Range, outputCells3 As Range

  Set inputCells1 = ws.Range("A6")
  Set inputCells2 = ws.Range("B6")
  Set inputCells3 = ws.Range("A7")

  Set outputCells1 = ws.Range("A9")
  Set outputCells2 = ws.Range("B9")
  Set outputCells3 = ws.Range("A10")

  If Not Intersect(Target, ws.Range("A6:B7")) Is Nothing Then
    
    ws.Range("G6").Value = "[" & inputCells1.Value & IIf(inputCells2 <> "-", ChrW(183) & inputCells2.Value, "") & IIf(inputCells3 <> "-", ChrW(47) & inputCells3.Value, "") & "]"
    ws.Range("S6").Value = "[" & inputCells1.Value & IIf(inputCells2 <> "-", ChrW(183) & inputCells2.Value, "") & IIf(inputCells3 <> "-", ChrW(47) & inputCells3.Value, "") & "]"
    ws.Range("AG6").Value = "[" & inputCells1.Value & IIf(inputCells2 <> "-", ChrW(183) & inputCells2.Value, "") & IIf(inputCells3 <> "-", ChrW(47) & inputCells3.Value, "") & "]"
    ws.Range("AV6").Value = "[" & inputCells1.Value & IIf(inputCells2 <> "-", ChrW(183) & inputCells2.Value, "") & IIf(inputCells3 <> "-", ChrW(47) & inputCells3.Value, "") & "]"
    ws.Range("BH6").Value = "[" & inputCells1.Value & IIf(inputCells2 <> "-", ChrW(183) & inputCells2.Value, "") & IIf(inputCells3 <> "-", ChrW(47) & inputCells3.Value, "") & "]"

  End If

  If Not Intersect(Target, ws.Range("A9:B10")) Is Nothing Then

    ws.Range("BY6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("CA6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("CL6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("CN6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("CY6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("DA6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("DF6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("DH6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("DP6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("DR6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("DZ6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("EB6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("EJ6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"
    ws.Range("EL6").Value = "[" & outputCells1.Value & IIf(outputCells2 <> "-", ChrW(183) & outputCells2.Value, "") & IIf(outputCells3 <> "-", ChrW(47) & outputCells3.Value, "") & "]"

  End If

End Sub
