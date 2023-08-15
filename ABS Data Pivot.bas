Attribute VB_Name = "Module1"
Dim eff_date As Date

Sub set_date()
    ' User must update the effective date
    eff_date = "07/31/2023"
End Sub

Sub copy()
    ' Create a copy of sheet and rename
    Sheets("Sheet1").copy After:=Sheets(1)
    Sheets("Sheet1 (2)").Name = "Copy"
End Sub

Sub transpose()
    ' Transpose the data
    Range("A1:P5").copy
    Range("A6").PasteSpecial transpose:=True
    ' Delete top 5 rows
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete
    ' Unselect
    Range("A1").Select
    Application.CutCopyMode = False
End Sub

Sub format()
    ' Remove the cell coloring
    Range("A1:E16").Interior.Color = xlNone
    ' Remove cell borders
    Range("A1:E16").Borders.LineStyle = xlNone
    ' Remove bolding from column A
    Range("A1:A16").Font.Bold = False
    ' Set font style
    Range("A1:E16").Font.Name = "Calibri"
    Range("A1:E16").Font.Size = 11
    ' Bold the headers
    Range("A1:F1").Font.Bold = True
End Sub

Sub effective_date()
    ' Add a column for the effective date
    Range("B1").EntireColumn.Insert
    ' Insert effective date into column
    Range("B2:B16").Value = eff_date
    ' Title the effective date column
    Range("B1").Value = "effective_date"
    Range("A1").Value = "subindex"
    Range("E1").Value = "excess_return_1m"
End Sub

Sub rearrange_columns()
    ' Insert columns then copy/paste to move - effective date
    Range("A1").EntireColumn.Insert
    Range("C1:C16").copy
    ActiveSheet.Paste Destination:=Worksheets("Copy").Range("A1")
    ActiveSheet.Columns("C").Delete
    ' Same for excess rtn
    Range("C1").EntireColumn.Insert
    Range("F1:F16").copy
    ActiveSheet.Paste Destination:=Worksheets("Copy").Range("C1")
    ActiveSheet.Columns("F").Delete
End Sub

Sub add_prefix()
    ' Add the BB_ prefix to each subindex cell
    Dim prefix As String
    Dim cell As Range
    
    prefix = "BB_"
    
    For Each cell In Range("B2:B16").Cells
        cell.Value = prefix & cell.Value
    Next cell
End Sub

Sub replacement()
    ' Replace spaces/slashes/dashes in subindex column with underscore
    Dim replacement As String
    Dim cell As Range
    
    replacement = "_"
    
    For Each cell In Range("B2:B16").Cells
        cell.Value = replace(cell.Value, "/", replacement)
        cell.Value = replace(cell.Value, " ", replacement)
        cell.Value = replace(cell.Value, "-", replacement)
        ' Delete last character in each cell
        cell.Value = Left(cell.Value, Len(cell.Value) - 1)
    Next cell
    
    ' Rename the LUABTRUU index
    Range("B2").Value = "BB_US_ABS_Index"
End Sub

Sub alphabetize()
    ' Sort the data alphabetically based on subindex
    With Worksheets("Copy").[A1].CurrentRegion
    .Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
    End With
End Sub

Sub all_macros()
    Call set_date
    Call copy
    Call transpose
    Call format
    Call effective_date
    Call rearrange_columns
    Call add_prefix
    Call replacement
    Call alphabetize
End Sub
