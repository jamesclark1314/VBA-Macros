Attribute VB_Name = "Module5"
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
        ' Select returns from the port column and transpose / paste to history workbook
        'TTF
        ws.Range("D7:D37").copy
        history_workbook.Worksheets("ABS Performance").Range("B4").PasteSpecial Transpose:=True, Paste:=xlPasteValues
        history_workbook.Worksheets("ABS Performance").Range("A4").Value = ws.Name
'        'GMS
'        ws.Range("J7:J37").copy
'        'NIF
'        ws.Range("P7:P37").copy
'        'STB
'        ws.Range("V7:V37").copy
    Next ws
End Sub

Sub all_macros()
    Call variable_definitions
'    Call copy
    Call transfer_abs_data
End Sub
