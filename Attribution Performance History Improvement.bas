Attribute VB_Name = "Module6"
' Set global variables
Dim history_workbook As Workbook

' USER MUST DEFINE THE NUMBER OF WORKBOOKS
Dim summary_workbook_name(1 To 6) As String

Sub variable_definitions()
    Dim history_workbook_name As String
    
    ' USER MUST UPDATE THE DESIRED WORKBOOK NAMES IN ORDER OF EARLIEST TO LATEST
    summary_workbook_name(1) = ".06 Securitized AA Historical Monthly Summary - 10.18-9.19"
    summary_workbook_name(2) = ".05 Securitized AA Historical Monthly Summary - 10.19-9.20"
    summary_workbook_name(3) = ".04 Securitized AA Historical Monthly Summary - 10.20-12.21"
    summary_workbook_name(4) = ".03 Securitized AA Historical Monthly Summary - 1.22 - 9.22"
    summary_workbook_name(5) = ".02 Securitized AA Historical Monthly Summary 10.22-6.23"
    summary_workbook_name(6) = ".01 Securitized AA Historical Monthly Summary 7.23-9.23"
    
    history_workbook_name = "Securitized Attribution Performance History"

    history_file_path = "C:\Users\trpjs86\OneDrive - TRowePrice\My Resources\Other\Attribution Performance History\" & history_workbook_name & ".xlsm"

    Set history_workbook = Workbooks.Open(history_file_path)
End Sub

Sub transfer_data()
    Dim ws As Worksheet

    ' Declare variant to hold the array element
    Dim item As Variant

    ' Loop through each item in the array of workbooks
    For Each item In summary_workbook_name
        summary_file_path = "C:\Users\trpjs86\OneDrive - TRowePrice\My Resources\Other\Attribution Performance History\" & item & ".xlsm"
        Debug.Print summary_file_path
        Set summary_workbook = Workbooks.Open(summary_file_path)

        For Each ws In summary_workbook.Worksheets
            If ws.Visible = xlSheetVisible Then
            
                ' Define last used row in the history column
                last_row_b = history_workbook.Worksheets("ABS Performance").Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Row
                last_row_a = history_workbook.Worksheets("ABS Performance").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row
         
                ' Label the row with the corresponding worksheet name
                history_workbook.Worksheets("ABS Performance").Range("A" & last_row_a).Value = ws.Name
         
                'TTF
                ' Select returns from the port column
                ws.Range("D5:D37").Copy
                ' Transpose / paste to history workbook
                history_workbook.Worksheets("ABS Performance").Range("B" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

                'GMS
                ' Select returns from the port column
                ws.Range("J5:J37").Copy
                ' Transpose / paste to history workbook
                history_workbook.Worksheets("ABS Performance").Range("AJ" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

                'NIF
                ' Select returns from the port column
                ws.Range("P5:P37").Copy
                ' Transpose / paste to history workbook
                history_workbook.Worksheets("ABS Performance").Range("BR" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues

                'STB
                ' Select returns from the port column
                ws.Range("V5:V37").Copy
                ' Transpose / paste to history workbook
                history_workbook.Worksheets("ABS Performance").Range("CZ" & last_row_b).PasteSpecial Transpose:=True, Paste:=xlPasteValues
                
            End If
        Next ws
        
        ' Keep first workbook in loop open but close all subsequent workbooks
        If item = summary_workbook_name(1) Then
        Else
            summary_workbook.Close False
        End If
        
    Next item
End Sub

Sub all_macros()
    Call variable_definitions
    Call transfer_data
End Sub
