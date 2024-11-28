Attribute VB_Name = "4-digit brackets"
Sub AddBracketsToFourDigitNumbers()
    Dim cell As Range
    Dim regex As Object
    Dim match As Object
    Dim inputText As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b(\d{4})\b" ' Matches a 4-digit number as a standalone word
    regex.Global = True ' Find all matches in the cell

    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            inputText = cell.Value
            If regex.test(inputText) Then
                cell.Value = regex.Replace(inputText, "($1)")
            End If
        End If
    Next cell

    Set regex = Nothing
End Sub

