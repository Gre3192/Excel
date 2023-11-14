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
    
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim macAddress As String, winAddress As String, MACcode As String
    
    winAddress = "winmgmts:\\.\root\cimv2"
    
    Set objWMIService = GetObject(winAddress)
    Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration")
    
    For Each objItem In colItems
        If Not IsNull(objItem.macAddress) Then
            macAddress = objItem.macAddress
            Exit For
        End If
    Next objItem
    
    Set objItem = Nothing
    Set colItems = Nothing
    Set objWMIService = Nothing
    
    ' ----------------------------------------------------------------------------------------- Verifica

    If ThisWorkbook.CustomDocumentProperties("isFirstStart").Value Then
        ThisWorkbook.CustomDocumentProperties("isFirstStart").Value = False
        ThisWorkbook.CustomDocumentProperties("StartLicense").Value = macAddress
    End If
    
    If ThisWorkbook.CustomDocumentProperties("StartLicense").Value <> macAddress Then
        MsgBox "Licensa non attiva!", vbExclamation
        ThisWorkbook.Close SaveChanges:=False
    End If
    
    ThisWorkbook.Save
    
End Sub
