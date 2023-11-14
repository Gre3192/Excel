'namespace=vba-files\Moduli
Attribute VB_Name = "Resetta"

Sub resetta_valori()

    Application.ScreenUpdating = False

    Dim button_clicked As String

    button_clicked = Application.caller

    if button_clicked="Resetta tutto" then
        Dim Button_involved As Variant
        Button_involved = Array("Resetta G1", "Resetta G2", "Resetta Qk", "Resetta P")
        For i = 0 To UBound(Button_involved)
            reset Button_involved(i) 
        Next
    else
        reset button_clicked 
    end if


End Sub

