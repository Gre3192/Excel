Attribute VB_Name = "FUNC_getMACcode"
Function getMACcode() As String

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim macAddress As String, winAddress As String
    
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
    
'    Debug.Print macAddress, IIf(VarType(macAddress) = vbString, "String", "Not String")

    getMACcode = macAddress

End Function

