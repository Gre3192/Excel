Attribute VB_Name = "Redirect_button"

Sub Redirect()
    
    Application.ScreenUpdating = False

    Dim nome As String
    nome_foglio = Application.Caller
    Application.Sheets(nome_foglio).Activate
    Application.Range("A13").Select
    
End Sub
