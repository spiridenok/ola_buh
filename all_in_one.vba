Const TURKEY = 1
Const ITALY = 2
Const GREECE = 3

Const RES_PK = 1
Const RES_ACCOUNT = 2
Const RES_AMOUNT = 3
Const RES_TAX_CODE = 4
Const RES_COST_CENTER = 6
Const RES_DESC = 11

Sub process(country As Integer)
    Dim ws As Worksheet
    Dim rangeNom As String
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets(1)

    Dim f As Workbook
    ' 2nd parameter == 0 suppresses the "Update links" message
    Set f = Workbooks.Open(pick_file(get_prompt(country)), 0)

    If country = ITALY Then
        Dim vendors_file_path As String
        vendors_file_path = Application.ActiveWorkbook.Path + "\Vendors Italy.xlsx"
        
        Dim vendors_file As Workbook
        If Dir(vendors_file_path) = "" Then
            MsgBox "File with Italy vendors is not found, please select a file (press 'Cancel' in the next file open dialog to continue without vendors list)"
            file_path = pick_file("Select Italy Vendors")
    
            If file_path <> "" Then
                Set vendors_file = Workbooks.Open(file_path)
            End If
        Else
            Set vendors_file = Workbooks.Open(vendors_file_path)
        End If
        
        If Not vendors_file Is Nothing Then
            Dim vendors_list As Worksheet
            Set vendors_list = vendors_file.Worksheets(1)
        End If
    End If
        
    Dim a As Long
    'a = ws.Range("A65000").End(xlUp).Row + 1
    a = 13
    
    Dim num_of_empty_rows As Integer
    num_of_empty_rows = 0
    
    Dim nums As New CBaseCountry
    nums.get_col_numbers country, f.Worksheets(1).rows
    
    For Each rw In f.Worksheets(1).rows
        If IsEmpty(rw.Cells(nums.get_debit)) And IsEmpty(rw.Cells(nums.get_credit)) Then
            ' Max of 2 empty rows are allowed after each other
            If num_of_empty_rows > 1 Then
                Exit For
            Else
                num_of_empty_rows = num_of_empty_rows + 1
            End If
        Else
            num_of_empty_rows = 0
        End If
        
        If Not country = GREECE Or (rw.Row > 2 And country = GREECE) Then  'For Greece we skip first 2 rows, will be fixed when dynamic row detection is implemented
            If IsNumeric(rw.Cells(nums.get_credit)) And rw.Cells(nums.get_credit) <> 0 Then
    '            MsgBox "Credit!"
                Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(nums.get_desc).Font.ColorIndex
                ws.Cells(a, RES_DESC) = rw.Cells(nums.get_desc)
                ws.Cells(a, RES_AMOUNT) = rw.Cells(nums.get_credit)
                ws.Cells(a, RES_ACCOUNT) = rw.Cells(nums.get_account)
                ws.Cells(a, RES_PK) = 50
                If ws.Cells(a, RES_ACCOUNT) = 113300 Or ws.Cells(a, RES_ACCOUNT) = 212230 Then ws.Cells(a, RES_ACCOUNT).Font.ColorIndex = 3
                Select Case country
                    Case TURKEY
                        If special_account_turkey(ws.Cells(a, RES_ACCOUNT)) Then
                            ws.Cells(a, RES_PK) = 31
                            Dim split_account() As String
                            split_account = Split(ws.Cells(a, RES_ACCOUNT), ".")
                            ws.Cells(a, RES_ACCOUNT) = split_account(UBound(split_account))
                        End If
                        If ws.Cells(a, RES_ACCOUNT) Like "5*" Then
                            ws.Cells(a, RES_TAX_CODE) = "V0"
                            ws.Cells(a, RES_COST_CENTER) = rw.Cells(nums.get_cost_center)
                        End If
                    Case GREECE
                        ws.Cells(a, RES_COST_CENTER) = rw.Cells(nums.get_cost_center)
                        If special_account(ws.Cells(a, RES_ACCOUNT)) Then ws.Cells(a, RES_PK).Value = 31
                        If ws.Cells(a, RES_ACCOUNT) = 212100 Or ws.Cells(a, RES_ACCOUNT) = 212110 Then ws.Cells(a, RES_ACCOUNT) = 8809
                        If ws.Cells(a, RES_ACCOUNT) = 214401 Then ws.Cells(a, RES_ACCOUNT) = 2413
                        If (ws.Cells(a, RES_ACCOUNT) Like "5*" Or ws.Cells(a, RES_ACCOUNT) Like "4*") And IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                        If Not ws.Cells(a, RES_ACCOUNT) Like "5*" And Not ws.Cells(a, RES_ACCOUNT) Like "4*" And Not IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                    Case ITALY
                        ws.Cells(a, RES_COST_CENTER) = rw.Cells(nums.get_cost_center)
                        If special_account(ws.Cells(a, RES_ACCOUNT)) Then
                            ws.Cells(a, RES_PK).Value = 31
                            If Not vendors_file Is Nothing Then
                                fill_italy_vendor ws.Cells(a, RES_ACCOUNT), rw.Cells(nums.get_alt_desc), rw.Cells(nums.get_desc), vendors_list
                            Else
                               ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                            End If
                        End If
                        If (ws.Cells(a, RES_ACCOUNT) Like "6*" Or ws.Cells(a, RES_ACCOUNT) Like "5*" Or ws.Cells(a, RES_ACCOUNT) Like "4*") And IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                        If Not ws.Cells(a, RES_ACCOUNT) Like "6*" And Not ws.Cells(a, RES_ACCOUNT) Like "5*" And Not ws.Cells(a, RES_ACCOUNT) Like "4*" And Not IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                End Select
                a = a + 1
            ElseIf IsNumeric(rw.Cells(nums.get_debit)) And rw.Cells(nums.get_debit) <> 0 Then
    '            MsgBox "Debit!"
                Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(nums.get_desc).Font.ColorIndex
                ws.Cells(a, RES_DESC) = rw.Cells(nums.get_desc)
                ws.Cells(a, RES_AMOUNT) = rw.Cells(nums.get_debit)
                ws.Cells(a, RES_PK) = 40
                ws.Cells(a, RES_ACCOUNT) = rw.Cells(nums.get_account)
                If ws.Cells(a, RES_ACCOUNT) = 113300 Or ws.Cells(a, RES_ACCOUNT) = 212230 Then ws.Cells(a, RES_ACCOUNT).Font.ColorIndex = 3
                Select Case country
                    Case TURKEY
                        If special_account_turkey(ws.Cells(a, RES_ACCOUNT)) Then
                            ws.Cells(a, RES_PK) = 21
                            ws.Cells(a, RES_TAX_CODE) = "**"
                            split_account = Split(ws.Cells(a, RES_ACCOUNT), ".")
                            ws.Cells(a, RES_ACCOUNT) = split_account(UBound(split_account))
                        End If
                        If ws.Cells(a, RES_ACCOUNT) Like "5*" Then
                            ws.Cells(a, RES_TAX_CODE) = "V0"
                            ws.Cells(a, RES_COST_CENTER) = rw.Cells(nums.get_cost_center)
                        End If
                    Case GREECE
                        ws.Cells(a, RES_COST_CENTER) = rw.Cells(nums.get_cost_center)
                        If special_account(ws.Cells(a, RES_ACCOUNT)) Then ws.Cells(a, RES_PK).Value = 21
                        If ws.Cells(a, RES_ACCOUNT) = 212100 Or ws.Cells(a, RES_ACCOUNT) = 212110 Then ws.Cells(a, RES_ACCOUNT) = 8809
                        If ws.Cells(a, RES_ACCOUNT) = 214401 Then ws.Cells(a, RES_ACCOUNT) = 2413
                        If (ws.Cells(a, RES_ACCOUNT) Like "5*" Or ws.Cells(a, RES_ACCOUNT) Like "4*") And IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                        If Not ws.Cells(a, RES_ACCOUNT) Like "5*" And Not ws.Cells(a, RES_ACCOUNT) Like "4*" And Not IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                    Case ITALY
                        ws.Cells(a, RES_COST_CENTER) = rw.Cells(nums.get_cost_center)
                        If special_account(ws.Cells(a, RES_ACCOUNT)) Then
                            ws.Cells(a, RES_PK).Value = 21
                            If Not vendors_file Is Nothing Then
                                fill_italy_vendor ws.Cells(a, RES_ACCOUNT), rw.Cells(nums.get_alt_desc), rw.Cells(nums.get_desc), vendors_list
                            Else
                               ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                            End If
                        End If
                        If (ws.Cells(a, RES_ACCOUNT) Like "6*" Or ws.Cells(a, RES_ACCOUNT) Like "5*" Or ws.Cells(a, RES_ACCOUNT) Like "4*") And IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                        If Not ws.Cells(a, RES_ACCOUNT) Like "6*" And Not ws.Cells(a, RES_ACCOUNT) Like "5*" And Not ws.Cells(a, RES_ACCOUNT) Like "4*" And Not IsEmpty(ws.Cells(a, RES_COST_CENTER)) Then
                            ws.Cells(a, RES_COST_CENTER).Interior.ColorIndex = 3
                        End If
                End Select
                a = a + 1
            End If
        End If
    Next rw
    
    ActiveWorkbook.Close
    If country = ITALY Then ActiveWorkbook.Close
    
End Sub

Function get_prompt(country As Integer) As String
    Select Case country
        Case TURKEY
            get_prompt = "Select Turkey"
        Case ITALY
            get_prompt = "Select Italy"
        Case GREECE
            get_prompt = "Select Greece"
    End Select
End Function

Sub fill_italy_vendor(ByRef account_cell As Range, ByVal desc_1 As Range, ByVal desc_2 As Range, ByRef vendors_list As Worksheet)
    For Each vendor_row In vendors_list.rows
        If IsEmpty(vendor_row.Cells(2)) Then
            account_cell.Interior.ColorIndex = 6
            Exit For
        End If
        If InStr(1, desc_1, vendor_row.Cells(2), vbTextCompare) > 0 Then
            account_cell = vendor_row.Cells(1)
            Exit For
        ElseIf InStr(1, desc_2, vendor_row.Cells(2), vbTextCompare) > 0 Then
            account_cell = vendor_row.Cells(1)
            Exit For
        End If
    Next vendor_row
End Sub

Function special_account(account_number As String) As Boolean
    If account_number = 212100 Or account_number = 212110 Or account_number = 214401 Or account_number = 212230 Or account_num = 113300 Then
        special_account = True
    Else
        special_account = False
    End If
End Function

Function special_account_turkey(account_number As String) As Boolean
    If InStr(account_number, ".") Then
        special_account_turkey = True
    Else
        special_account_turkey = False
    End If
End Function

Function pick_file(prompt As String) As String
    Dim filePicker As FileDialog

    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)

    With filePicker

        'setup File Dialog'
        .AllowMultiSelect = False
        .ButtonName = prompt
        .InitialView = msoFileDialogViewList
        .title = prompt
'        .InitialFileName = ""

        'add filter for all files'
        With .Filters
        .Clear
        .Add "All Files", "*.xls*"
        End With
        .FilterIndex = 1

        'display file dialog box'
        .Show

     End With
    If filePicker.SelectedItems.Count > 0 Then
        pick_file = filePicker.SelectedItems(1)
    Else
        pick_file = ""
    End If

End Function

Sub ClearStatementData()
    ' To speed up the process only used columns are cleared
    For Each cell In Range("A13:F1000", "K13:K1000").Cells
        cell.ClearContents
        cell.Interior.ColorIndex = 2
        cell.Borders.ColorIndex = 15
        cell.Font.ColorIndex = 1
    Next cell
    ActiveWorkbook.Save
    Range("A13").Select
End Sub


Sub dspi_test()
    'process (TURKEY)
    process (GREECE)
    'process (ITALY)
End Sub

