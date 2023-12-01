Attribute VB_Name = "Resetta"
'namespace=vba-files\Moduli

Sub resetta_valori()

    Application.ScreenUpdating = False

    Dim button_clicked As String

    button_clicked = Application.caller

    If button_clicked = "Resetta tutto" Then

        Dim buttonInvolved() As Variant
        buttonInvolved = Array("Resetta G1", "Resetta G2", "Resetta Qk", "Resetta P", "Resetta SLU", "Resetta SLE RARA", "Resetta SLE FREQUENTE", "Resetta SLE Q.P.")
        
        For Each ResetButton In buttonInvolved
            reset ResetButton
        Next ResetButton

    Else

        reset button_clicked

    End If

End Sub
