Attribute VB_Name = "CustomProp"
Sub CustomProperties()

    Dim sh As Worksheet
    Dim prop As Object

    Set sh = ThisWorkbook.Sheets("Listino prezzi")

    Debug.Print "------------------ Prima:"
    For Each prop In sh.CustomProperties
        Debug.Print prop.Name & ": " & prop.Value
    Next prop
' ==============================================================
    sh.CustomProperties.Item(1).Value = 10
'   sh.CustomProperties.Item(2).Delete
' ==============================================================
    Debug.Print "------------------- Dopo:"
    For Each prop In sh.CustomProperties
        Debug.Print prop.Name & ": " & prop.Value
    Next prop
    
End Sub
