Attribute VB_Name = "Capitalise First Word"
Sub CapitalizeFirstLetter()
    Dim cell As Range
    Dim inputText As String
    
    For Each cell In Selection
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) = False Then
            inputText = cell.Value
            cell.Value = StrConv(inputText, vbProperCase)
        End If
    Next cell
End Sub

