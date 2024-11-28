Attribute VB_Name = "Remove exact duplicates"
Sub RemoveExactDuplicates()
    Dim cell As Range
    Dim compareRange As Range
    Dim cellText As String

    Set compareRange = Selection

    For Each cell In compareRange
        If Not IsEmpty(cell.Value) Then
            cellText = cell.Value
            If WorksheetFunction.CountIf(compareRange, cellText) > 1 Then
                cell.ClearContents
            End If
        End If
    Next cell

    MsgBox "Duplicates removed from the selected column!", vbInformation, "Done"
End Sub

