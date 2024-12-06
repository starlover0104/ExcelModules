Attribute VB_Name = "Module1"
Sub SortSelectedColumns()
    Dim ws As Worksheet
    Dim rng As Range
    Dim sortCol As Range
    Dim colIndex As Integer
    Dim sortDirection As XlSortOrder
    Dim lastRow As Long
    Dim response As Integer
    Dim cell As Range
    Dim isAlphabetical As Boolean
    Dim colToSort As Range

    Set ws = ActiveSheet

    response = MsgBox("Sort in Ascending order? Click 'No' for Descending.", vbYesNoCancel + vbQuestion, "Sort Direction")
    If response = vbCancel Then Exit Sub
    sortDirection = IIf(response = vbYes, xlAscending, xlDescending)

    If Selection.Columns.Count < 1 Then
        MsgBox "Please select the columns you want to sort.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    lastRow = ws.Cells(ws.Rows.Count, rng.Column).End(xlUp).Row

    For colIndex = 1 To rng.Columns.Count
        Set colToSort = ws.Range(rng.Cells(1, colIndex), ws.Cells(lastRow, rng.Cells(1, colIndex).Column))
        isAlphabetical = False
        For Each cell In colToSort
            If cell.Value Like "*[A-Za-z]*" Then
                isAlphabetical = True
                Exit For
            End If
        Next cell

        If isAlphabetical Then
            ws.Sort.SortFields.Clear
            ws.Sort.SortFields.Add Key:=colToSort, Order:=sortDirection

            With ws.Sort
                .SetRange ws.Range(rng.Address)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    Next colIndex

    MsgBox "Columns with alphabetical data have been sorted.", vbInformation
End Sub

