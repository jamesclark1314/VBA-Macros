Attribute VB_Name = "Module11"
' Set global variables
Dim summary_workbook As Workbook
Dim history_workbook As Workbook

Sub variable_definitions()
    Dim summary_workbook_name As String
    Dim history_workbook_name As String
    
    ' USER MUST UPDATE THE DESIRED WORKBOOK NAMES
    summary_workbook_name = ".06 Securitized AA Historical Monthly Summary - 10.18-9.19"
    history_workbook_name = "Securitized Attribution Performance History"
    
    Dim summary_file_path As String
    Dim history_file_path As String
    summary_file_path = "C:\Users\trpjs86\OneDrive - TRowePrice\My Resources\Other\Attribution Performance History\" & summary_workbook_name & ".xlsm"
    history_file_path = "C:\Users\trpjs86\OneDrive - TRowePrice\My Resources\Other\Attribution Performance History\" & history_workbook_name & ".xlsm"
    
    Set summary_workbook = Workbooks.Open(summary_file_path)
    Set history_workbook = Workbooks.Open(history_file_path)
End Sub

'Sub copy()
'    ' Create a copy of sheet and rename
'    summary_workbook.Worksheets("Oct 2018").copy After:=summary_workbook.Worksheets(1)
'    summary_workbook.Worksheets("Oct 2018 (2)").Name = "Copy"
'End Sub

Sub transfer_abs_data()
    Dim ws As Worksheet
    For Each ws In summary_workbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            
            'TTF
            
            ' Select returns from the port column
            ws.Range("D5:D37").copy
            
            ' Define last used row in the history column
            last_row_b = history_workbook.Worksheets("ABS Performance").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Row
            last_row_a = history_workbook.Worksheets("ABS Performance").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row
            
            ' Transpose / paste to history workbook
            history_workbook.Worksheets("ABS Performance").Range("B" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
            
            ' Label the row with the corresponding worksheet name
            history_workbook.Worksheets("ABS Performance").Range("A" & last_row_a).Value = ws.Name
    
            'GMS
            
            ' Select returns from the port column
            ws.Range("J5:J37").copy
            
            ' Transpose / paste to history workbook
            history_workbook.Worksheets("ABS Performance").Range("AJ" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
    
            'NIF
            
            ' Select returns from the port column
            ws.Range("P5:P37").copy
            
            ' Transpose / paste to history workbook
            history_workbook.Worksheets("ABS Performance").Range("BR" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
    
            'STB
            ' Select returns from the port column
            ws.Range("V5:V37").copy
            
            ' Transpose / paste to history workbook
            history_workbook.Worksheets("ABS Performance").Range("CZ" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
        End If
    Next ws
End Sub

Sub transfer_cmbs_data()
    Dim ws As Worksheet
    For Each ws In summary_workbook.Worksheets
        If ws.Visible = xlSheetVisible Then

            'TTF

            ' Select returns from the port column
            ws.Range("D42:D64").copy

            ' Define last used row in the history column
            last_row_b = history_workbook.Worksheets("CMBS Performance").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Row
            last_row_a = history_workbook.Worksheets("CMBS Performance").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CMBS Performance").Range("B" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            ' Label the row with the corresponding worksheet name
            history_workbook.Worksheets("CMBS Performance").Range("A" & last_row_a).Value = ws.Name

            'GMS

            ' Select returns from the port column
            ws.Range("J42:J64").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CMBS Performance").Range("Z" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            'NIF

            ' Select returns from the port column
            ws.Range("P42:P64").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CMBS Performance").Range("AX" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            'STB
            ' Select returns from the port column
            ws.Range("V42:V64").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CMBS Performance").Range("BV" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
        End If
    Next ws
End Sub

Sub transfer_rmbs_data()
    Dim ws As Worksheet
    For Each ws In summary_workbook.Worksheets
        If ws.Visible = xlSheetVisible Then

            'TTF

            ' Select returns from the port column
            ws.Range("D69:D110").copy

            ' Define last used row in the history column
            last_row_b = history_workbook.Worksheets("RMBS Performance").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Row
            last_row_a = history_workbook.Worksheets("RMBS Performance").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("RMBS Performance").Range("B" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            ' Label the row with the corresponding worksheet name
            history_workbook.Worksheets("RMBS Performance").Range("A" & last_row_a).Value = ws.Name

            'GMS

            ' Select returns from the port column
            ws.Range("J69:J110").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("RMBS Performance").Range("AS" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            'NIF

            ' Select returns from the port column
            ws.Range("P69:P110").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("RMBS Performance").Range("CJ" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            'STB
            ' Select returns from the port column
            ws.Range("V69:V110").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("RMBS Performance").Range("EA" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
        End If
    Next ws
End Sub

Sub transfer_clo_data()
    Dim ws As Worksheet
    For Each ws In summary_workbook.Worksheets
        If ws.Visible = xlSheetVisible Then

            'TTF

            ' Select returns from the port column
            ws.Range("D115:D121").copy

            ' Define last used row in the history column
            last_row_b = history_workbook.Worksheets("CLO Performance").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Row
            last_row_a = history_workbook.Worksheets("CLO Performance").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CLO Performance").Range("B" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            ' Label the row with the corresponding worksheet name
            history_workbook.Worksheets("CLO Performance").Range("A" & last_row_a).Value = ws.Name

            'GMS

            ' Select returns from the port column
            ws.Range("J115:J121").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CLO Performance").Range("J" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            'NIF

            ' Select returns from the port column
            ws.Range("P115:P121").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CLO Performance").Range("R" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

            'STB
            ' Select returns from the port column
            ws.Range("V115:V121").copy

            ' Transpose / paste to history workbook
            history_workbook.Worksheets("CLO Performance").Range("Z" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
        End If
    Next ws
End Sub

Sub all_macros()
    Call variable_definitions
'    Call copy
    Call transfer_abs_data
    Call transfer_cmbs_data
    Call transfer_rmbs_data
    Call transfer_clo_data
End Sub
