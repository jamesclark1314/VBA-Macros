Attribute VB_Name = "Module1"
Sub copy()
    ' Create a copy of sheet and rename
    Sheets("Sheet1").copy After:=Sheets(1)
    Sheets("Sheet1 (2)").Name = "Copy"
End Sub

Sub delete_columns()
    ' Delete unnecessary columns
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String

    currentColumn = 1
        While currentColumn <= ActiveSheet.UsedRange.Columns.Count
            columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

            ' Columns to keep
            keepColumn = False
            If columnHeading = "issuer_name" Then keepColumn = True
            If columnHeading = "seniority" Then keepColumn = True
            If columnHeading = "sp" Then keepColumn = True
            If columnHeading = "S&P Flag" Then keepColumn = True
            If columnHeading = "Prev sp" Then keepColumn = True
            If columnHeading = "moodys" Then keepColumn = True
            If columnHeading = "Moody's Flag" Then keepColumn = True
            If columnHeading = "Prev moodys" Then keepColumn = True
            If columnHeading = "Fac Size" Then keepColumn = True

            ' If keepColumn = True then skip
            If keepColumn Then
                currentColumn = currentColumn + 1
            Else
            ' If keepColumn = False then delete
                ActiveSheet.Columns(currentColumn).Delete
            End If

        Wend
End Sub

Sub senority_filter()
    ' Filter out 2nd/3rd Lien
    With Worksheets("Copy").[A1].CurrentRegion
        .AutoFilter 2, "2ND/3RD LIEN SECURED"
        If .Rows.Count > 1 Then
            .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        .AutoFilter ' Remove the filter
    End With
End Sub

Sub rating_chg_filter()
    ' Filter by ratings change - leave only newly B-/CCC+
    Dim data As Range
    Dim spColumn As Range
    Dim prevspColumn As Range
    Dim moodysColumn As Range
    Dim prevMoodysColumn As Range

    Set data = Worksheets("Copy").Range("A1").CurrentRegion

    Set spColumn = data.Columns(3)
    Set prevspColumn = data.Columns(5)
    Set moodysColumn = data.Columns(6)
    Set prevMoodysColumn = data.Columns(8)

    Dim row As Long
    For row = data.Rows.Count To 2 Step -1
        If (spColumn.Cells(row).Value = "B-" And prevspColumn.Cells(row).Value <> "B-") Or _
           (spColumn.Cells(row).Value = "CCC+" And prevspColumn.Cells(row).Value <> "CCC+") Or _
           (moodysColumn.Cells(row).Value = "B3" And prevMoodysColumn.Cells(row).Value <> "B3") Or _
           (moodysColumn.Cells(row).Value = "Caa1" And prevMoodysColumn.Cells(row).Value <> "Caa1") Then
        Else
            data.Rows(row).Delete
        End If
    Next row
End Sub

Sub rename_columns()
    ' Rename columns
    [A1].Value = "Issuer"
    [C1].Value = "S&P Curr"
    [E1].Value = "Prev S&P"
    [F1].Value = "Moody's Curr"
    [H1].Value = "Prev Moody's"
    
    ' Delete senority and extra S&P flag column
    Columns(2).Delete
    Columns(8).Delete
End Sub

Sub sort()
    With Worksheets("Copy").[A1].CurrentRegion
        .sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlYes
    End With
End Sub

Sub all_macros()
    Call copy
    Call delete_columns
    Call senority_filter
    Call rating_chg_filter
    Call rename_columns
    Call sort
End Sub

