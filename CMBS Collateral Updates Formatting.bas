Attribute VB_Name = "Module1"
Sub copy()
    ' Create a copy of sheet and rename
    Sheets("Sheet1").copy After:=Sheets(1)
    Sheets("Sheet1 (2)").name = "Copy"
End Sub

Sub spacers()
    Dim data As Range
    Dim update_column As Range
    
    Set data = Worksheets("Copy").Range("A1").CurrentRegion
    Set update_column = data.Columns(1)
    
    ' Check to see if cell is bold and if so, add a space
    Dim current_row As Long
    For current_row = data.Rows.Count To 3 Step -1
        If update_column.Cells(current_row).Font.Bold = False Then
        Else
            update_column.Cells(current_row).EntireRow.Insert
        End If
    Next current_row
End Sub

Sub fill_analyst()
    Dim last_row As Long
    
    ' Define the last used row in column A
    last_row = Cells(Rows.Count, "A").End(xlUp).row
    
    ' Fill blank cells in column B with value above
    With ThisWorkbook.Worksheets("Copy").Range("B1:B" & last_row)
        .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Value = .Value
    End With
End Sub

Sub all_macros()
    Call copy
    Call spacers
    Call fill_analyst
End Sub
