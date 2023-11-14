Attribute VB_Name = "sheetProp"

Function sheetProprieties()

    Dim prop As Object
    Dim Arrasheet() As Variant
    Arrasheet = Array("Elenco Ditte", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre", "Uscite")
    

    ' Debug.Print "------------------ Prima:"
    ' For Each sheet In Arrasheet
    '     Debug.Print "---" & sheet & "--------------------------"
    '     For Each prop In ThisWorkbook.Sheets(sheet).CustomProperties
    '         Debug.Print prop.Name & ": " & prop.Value
    '     Next prop
    ' Next prop
    

    ' AGGIUNGI PROP ================================================

        ' ThisWorkbook.Sheets("Elenco Ditte").CustomProperties.Add Name:="lastRowCiane", Value:=44
        ' ThisWorkbook.Sheets("Elenco Ditte").CustomProperties.Add Name:="lastRowFornitori", Value:=35

        ' ThisWorkbook.Sheets("Gennaio").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Gennaio").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Febbraio").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Febbraio").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Marzo").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Marzo").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Aprile").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Aprile").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Maggio").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Maggio").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Giugno").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Giugno").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Luglio").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Luglio").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Agosto").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Agosto").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Settembre").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Settembre").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Ottobre").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Ottobre").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Novembre").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Novembre").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Dicembre").CustomProperties.Add Name:="lastRowCiane", Value:=191
        ' ThisWorkbook.Sheets("Dicembre").CustomProperties.Add Name:="lastRowFornitori", Value:=137

        ' ThisWorkbook.Sheets("Uscite").CustomProperties.Add Name:="lastRowCiane", Value:=49
        ' ThisWorkbook.Sheets("Uscite").CustomProperties.Add Name:="lastRowFornitori", Value:=40

    ' ' EDITA PROP ================================================
        ThisWorkbook.Sheets("Elenco Ditte").CustomProperties.Item(1).Value = 44
        ThisWorkbook.Sheets("Elenco Ditte").CustomProperties.Item(2).Value = 35

    ' ==============================================================
    ' Debug.Print "------------------ Dopo:"
    ' For Each sheet In Arrasheet
    '     Debug.Print "---" & sheet & "--------------------------"
    '     For Each prop In ThisWorkbook.Sheets(sheet).CustomProperties
    '         Debug.Print prop.Name & ": " & prop.Value
    '     Next prop
    ' Next prop


End Function
