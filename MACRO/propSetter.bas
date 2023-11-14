Attribute VB_Name = "propSetter"
Sub setCustomProperty()

    Dim customPropertyName As String, deleteCustomProperties As String
    Dim customPropertyValue As Variant
    Dim prop As DocumentProperty
    
'   ======================================================================================== AGGIUNGI PROP
    customPropertyName = "isFirstStart"
    customPropertyValue = True
    On Error Resume Next
        ThisWorkbook.CustomDocumentProperties.Add _
        Name:=customPropertyName, LinkToContent:=False, Type:=msoPropertyTypeString, Value:=customPropertyValue
    On Error GoTo 0
    
'   ======================================================================================== SETTA PROP INIZIALI LICENSA
'    ThisWorkbook.CustomDocumentProperties("StartLicense").Value = "Arshes19@Xenoth92"
'    ThisWorkbook.CustomDocumentProperties("isFirstStart").Value = True
    
'   ======================================================================================== CANCELLA PROP
'    deleteCustomProperties = "_"
'    Debug.Print "Proprietà '" & deleteCustomProperties & "' eliminata con successo."
'    ThisWorkbook.CustomDocumentProperties(deleteCustomProperties).Delete
    
'   ======================================================================================== VISUALIZZA PROP
    For Each prop In ThisWorkbook.CustomDocumentProperties
        Debug.Print prop.Name & ": " & prop.Value
    Next prop

End Sub


