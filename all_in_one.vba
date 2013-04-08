Const TURKEY = 1
Const ITALY = 2
Const GREECE = 3

Const RES_PK = 1
Const RES_ACCOUNT = 2
Const RES_AMOUNT = 3
Const RES_TAX_CODE = 4
Const RES_COST_CENTER = 6
Const RES_DESC = 11

Function get_column_numbers(country As Integer) As CBaseCountry
    Set get_column_numbers = New CTurkey 'CBaseCountry
End Function

Sub process(country As Integer)
    Dim ws As Worksheet
    Dim rangeNom As String
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets(1)

    Dim f As Workbook
    ' 2nd parameter == 0 suppresses the "Update links" message
    Set f = Workbooks.Open(pick_file(get_prompt(country)), 0)

    Dim a As Long
    'a = ws.Range("A65000").End(xlUp).Row + 1
    a = 13
    
    Dim num_of_empty_rows As Integer
    num_of_empty_rows = 0
    
    Dim nums As CBaseCountry
    Set nums = get_column_numbers(country)
    
    For Each rw In f.Worksheets(1).Rows
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
        
        If IsNumeric(rw.Cells(nums.get_credit)) And rw.Cells(nums.get_credit) <> 0 Then
'            MsgBox "Credit!"
            Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(nums.get_desc).Font.ColorIndex
            ws.Cells(a, RES_DESC) = rw.Cells(nums.get_desc)
            ws.Cells(a, RES_AMOUNT) = rw.Cells(nums.get_credit)
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(nums.get_accout)
            ws.Cells(a, RES_PK) = 50
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
            a = a + 1
        ElseIf IsNumeric(rw.Cells(nums.get_debit)) And rw.Cells(nums.get_debit) <> 0 Then
'            MsgBox "Debit!"
            Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(nums.get_desc).Font.ColorIndex
            ws.Cells(a, RES_DESC) = rw.Cells(nums.get_desc)
            ws.Cells(a, RES_AMOUNT) = rw.Cells(nums.get_debit)
            ws.Cells(a, RES_PK) = 40
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(nums.get_accout)
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
            a = a + 1
        End If
    Next rw
    
    ActiveWorkbook.Close
    
End Sub

Sub dspi_test()
    process (TURKEY)
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
    For Each vendor_row In vendors_list.Rows
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
    If account_number = 212100 Or account_number = 212110 Or account_number = 214401 Or account_number = 212230 Then
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


Sub GREECE_SUB()
    
    Set ws = ActiveSheet
'    rangeNom = "PK"
    
    Dim a As Long
    a = 13
    
    Dim f As Workbook
    
    ' 2nd parameter == 0 suppresses the "Update links" message
    Set f = Workbooks.Open(pick_file("Select Greece"), 0)
    
    Dim num_of_empty_rows As Integer
    num_of_empty_rows = 0
    
    For Each rw In f.Worksheets(1).Rows
        Const DESC = 7 ' Description
        Const ACCOUNT = 5 'GL Account
        Const COST_CENTER = 10
        Const DEBIT = 8
        Const CREDIT = 9
    
        If IsEmpty(rw.Cells(DEBIT)) And IsEmpty(rw.Cells(CREDIT)) Then
            ' Max of 2 empty rows are allowed after each other
            If num_of_empty_rows > 1 Then
                Exit For
            Else
                num_of_empty_rows = num_of_empty_rows + 1
            End If
        Else
            num_of_empty_rows = 0
        End If
        
        If rw.Row > 2 Then 'For Greece we skip first 2 rows, will be fixed when dynamic row detection is implemented
            If Not IsEmpty(rw.Cells(CREDIT)) And IsNumeric(rw.Cells(CREDIT)) And rw.Cells(CREDIT) <> 0 Then
    '            MsgBox "Credit!"
                ws.Cells(a, RES_DESC) = rw.Cells(DESC).Value
                Range(ws.Cells(a, 1), ws.Cells(a, 15)).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
                ws.Cells(a, RES_AMOUNT) = rw.Cells(CREDIT).Value
                ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT).Value
                ws.Cells(a, RES_PK) = 50
                ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER).Value
    '            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
                If special_account(ws.Cells(a, RES_ACCOUNT)) Then ws.Cells(a, RES_PK).Value = 31
                ' TAX code is empty
                a = a + 1
            ElseIf Not IsEmpty(rw.Cells(DEBIT)) And IsNumeric(rw.Cells(DEBIT)) And rw.Cells(DEBIT) <> 0 Then
    '            MsgBox "Debit!"
                ws.Cells(a, RES_DESC) = rw.Cells(DESC)
                Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
                ws.Cells(a, RES_AMOUNT) = rw.Cells(DEBIT)
                ws.Cells(a, RES_PK) = 40
                ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER)
                ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT)
                If special_account(ws.Cells(a, RES_ACCOUNT)) Then ws.Cells(a, RES_PK).Value = 21
                ' TAX code is empty
                a = a + 1
            End If
        End If
    Next rw
    ActiveWorkbook.Close
End Sub

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

Sub Turkije()
    Dim ws As Worksheet
    Dim rangeNom As String
    Dim nextRow As Long

    Set base_book = ThisWorkbook
    Set ws = ActiveSheet
    rangeNom = "PK"
    
    'nextRow = ws.Columns'
    Dim a As Long
    
    'a = ws.Range("A65000").End(xlUp).Row + 1
    a = 13
    
    Dim f As Workbook
    
    Set f = Workbooks.Open(pick_file("Select Turkey"), 0)
    
    Dim num_of_empty_rows As Integer
    num_of_empty_rows = 0
    
    Const ACCOUNT = 4 'GL Account
    Const DESC = 7 ' Description
    Const COST_CENTER = 14
    Const DEBIT = 9
    Const CREDIT = 10
    
    For Each rw In f.Worksheets(1).Rows
        If IsEmpty(rw.Cells(DEBIT)) And IsEmpty(rw.Cells(CREDIT)) Then
            ' Max of 2 empty rows are allowed after each other
            If num_of_empty_rows > 1 Then
                Exit For
            Else
                num_of_empty_rows = num_of_empty_rows + 1
            End If
        Else
            num_of_empty_rows = 0
        End If
        
        If IsNumeric(rw.Cells(CREDIT)) And rw.Cells(CREDIT) <> 0 Then
'            MsgBox "Credit!"
            Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, RES_DESC) = rw.Cells(DESC)
            ws.Cells(a, RES_AMOUNT) = rw.Cells(CREDIT)
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT)
            ws.Cells(a, RES_PK) = 50
            If special_account_turkey(ws.Cells(a, RES_ACCOUNT)) Then
                ws.Cells(a, RES_PK) = 31
                Dim split_account() As String
                split_account = Split(ws.Cells(a, RES_ACCOUNT), ".")
                ws.Cells(a, RES_ACCOUNT) = split_account(UBound(split_account))
            End If
            If rw.Cells(ACCOUNT) Like "5*" Then
                ws.Cells(a, RES_TAX_CODE) = "V0"
                ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER)
            End If
            a = a + 1
        ElseIf IsNumeric(rw.Cells(DEBIT)) And rw.Cells(DEBIT) <> 0 Then
'            MsgBox "Debit!"
            Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, RES_DESC) = rw.Cells(DESC)
            ws.Cells(a, RES_AMOUNT) = rw.Cells(DEBIT)
            ws.Cells(a, RES_PK) = 40
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT)
            If special_account_turkey(ws.Cells(a, RES_ACCOUNT)) Then
                ws.Cells(a, RES_PK) = 21
                ws.Cells(a, RES_TAX_CODE) = "**"
                'Dim split_account() As String
                split_account = Split(ws.Cells(a, RES_ACCOUNT), ".")
                ws.Cells(a, RES_ACCOUNT) = split_account(UBound(split_account))
            End If
            If rw.Cells(ACCOUNT) Like "5*" Then
                ws.Cells(a, RES_TAX_CODE) = "V0"
                ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER)
            End If
            a = a + 1
        End If
    Next rw
    
    ActiveWorkbook.Close
End Sub

Sub ITALY_SUB()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim a As Long
'    a = ws.Range("A65000").End(xlUp).Row + 1
    a = 13
    
    Dim vendors_file_path As String
    vendors_file_path = Application.ActiveWorkbook.Path + "\Vendors Italy.xlsx"
'    vendors_file_path = vendors_file_path + "bla"
    
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
    
    Dim f As Workbook
    Set f = Workbooks.Open(pick_file("Select Italy"), 0)
    
    Dim num_of_empty_rows As Integer
    num_of_empty_rows = 0
    
    For Each rw In f.Worksheets(1).Rows
        Const DESC = 8 ' Description
        Const desc_1 = 7 ' Description to search in the vendors list
        Const ACCOUNT = 3 'GL Account
        Const COST_CENTER = 5
        Const DEBIT = 10
        Const CREDIT = 11
    
        If IsEmpty(rw.Cells(DEBIT)) And IsEmpty(rw.Cells(CREDIT)) Then
            If num_of_empty_rows > 3 Then
                Exit For
            Else
                num_of_empty_rows = num_of_empty_rows + 1
            End If
        Else
            num_of_empty_rows = 0
        End If
        
        If IsNumeric(rw.Cells(CREDIT)) And rw.Cells(CREDIT) <> 0 Then
'            MsgBox "Credit!"
            ws.Cells(a, RES_DESC) = rw.Cells(DESC).Value
            ws.Cells(a, RES_DESC).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, RES_AMOUNT) = rw.Cells(CREDIT)
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT)
            ws.Cells(a, RES_PK) = 50
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER)
            If special_account(ws.Cells(a, RES_ACCOUNT)) Then
                ws.Cells(a, RES_PK).Value = 31
                If Not vendors_file Is Nothing Then
                    fill_italy_vendor ws.Cells(a, RES_ACCOUNT), rw.Cells(desc_1), rw.Cells(DESC), vendors_list
                Else
                   ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                End If
            End If
            ' TAX code is empty
            a = a + 1
        ElseIf IsNumeric(rw.Cells(DEBIT)) And rw.Cells(DEBIT) <> 0 Then
'            MsgBox "Debit!"
            ws.Cells(a, RES_DESC) = rw.Cells(DESC)
            ws.Cells(a, RES_DESC).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            Range(ws.Cells(a, 1), ws.Cells(a, 12)).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, RES_AMOUNT) = rw.Cells(DEBIT)
            ws.Cells(a, RES_PK) = 40
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER)
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT)
            If special_account(ws.Cells(a, RES_ACCOUNT)) Then
                ws.Cells(a, RES_PK).Value = 21
                If Not vendors_file Is Nothing Then
                    fill_italy_vendor ws.Cells(a, RES_ACCOUNT), rw.Cells(desc_1), rw.Cells(DESC), vendors_list
                Else
                   ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                End If
            End If
            ' TAX code is empty
            a = a + 1
        End If
    Next rw
    ActiveWorkbook.Close
    If Not vendors_file Is Nothing Then ActiveWorkbook.Close
End Sub
