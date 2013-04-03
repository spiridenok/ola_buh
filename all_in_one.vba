Function special_account(account_number As String) As Boolean
    If account_number = 212100 Or account_number = 212110 Or account_number = 214401 Or account_number = 212230 Then
        special_account = True
    Else
        special_account = False
    End If
End Function

Sub Turkije()
    Dim filePicker As FileDialog

    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)

    With filePicker

        'setup File Dialog'
        .AllowMultiSelect = False
        .ButtonName = "Select Turkije"
        .InitialView = msoFileDialogViewList
        .Title = "Select Turkije"
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

        Dim selectedFile As String
        selectedFile = filePicker.SelectedItems(1)

    End If
    
    Dim ws As Worksheet
    Dim rangeNom As String
    Dim nextRow As Long

    Set base_book = ThisWorkbook
    Set ws = ActiveSheet
    rangeNom = "PK"
    
    'nextRow = ws.Columns'
    Dim a As Long
    
    a = ws.Range("A65000").End(xlUp).Row + 1
'    a = ActiveSheet.UsedRange.Rows.Count
    
    Dim f As Workbook
    
    Set f = Workbooks.Open(selectedFile)
    Dim active_f As Sheets
    
    Dim i As Integer
    
    For Each rw In f.Worksheets(1).Rows
'    For Each rw In f.Worksheets("Sayfa1").Rows
        If IsEmpty(rw.Cells(1).Value) Then Exit For
        
        If IsEmpty(rw.Cells(8).Value) Then
'            MsgBox "Empty!"
        ElseIf rw.Cells(8).Value = 0 Then
'            MsgBox "Credit!"
            ws.Cells(a, 11) = rw.Cells(6).Value
            ws.Cells(a, 3) = rw.Cells(9).Value
            ws.Cells(a, 2) = rw.Cells(4).Value
            ws.Cells(a, 1) = 50
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, 2) = 212100 Or ws.Cells(a, 2) = 212110 Or ws.Cells(a, 2) = 214401 Or ws.Cells(a, 2) = 212230 Then ws.Cells(a, 1).Value = 31
            If rw.Cells(4).Value Like "5*" Then ws.Cells(a, 4) = "V0"
            If rw.Cells(4).Value Like "5*" Then ws.Cells(a, 6).Value = rw.Cells(13).Value
            a = a + 1
        ElseIf IsNumeric(rw.Cells(8).Value) Then
'            MsgBox "Debit!"
            ws.Cells(a, 11) = rw.Cells(6).Value
            ws.Cells(a, 3) = rw.Cells(8).Value
            ws.Cells(a, 1) = 40
            ws.Cells(a, 2) = rw.Cells(4).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, 2) = 212100 Or ws.Cells(a, 2) = 212110 Or ws.Cells(a, 2) = 214401 Or ws.Cells(a, 2) = 212230 Then ws.Cells(a, 1).Value = 21
            If ws.Cells(a, 1).Value = 21 Then ws.Cells(a, 4) = "**"
            If rw.Cells(4).Value Like "5*" Then ws.Cells(a, 4) = "V0"
            If rw.Cells(4).Value Like "5*" Then ws.Cells(a, 6).Value = rw.Cells(13).Value
            a = a + 1
        Else
'            MsgBox "Empty!"
        End If
        i = i + 1
    Next rw
'    MsgBox f.Worksheets("Sayfa1").Cells(3, 8).Value
    
'    For Each s In f.Worksheets
'        MsgBox s.Name
'    Next s
     
    
End Sub

Sub Greece()
    Dim filePicker As FileDialog

    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)

    With filePicker

        'setup File Dialog'
        .AllowMultiSelect = False
        .ButtonName = "Select Griekenland"
        .InitialView = msoFileDialogViewList
        .Title = "Select Greece"
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

        Dim selectedFile As String
        selectedFile = filePicker.SelectedItems(1)

    End If
    
    Dim ws As Worksheet
'    Dim rangeNom As String
'    Dim nextRow As Long

'    Set base_book = ThisWorkbook
    Set ws = ActiveSheet
'    rangeNom = "PK"
    
    Dim a As Long
    
    a = ws.Range("A65000").End(xlUp).Row + 1
    
    Dim f As Workbook
    
    ' 2nd parameter == 0 suppresses the "Update links" message
    Set f = Workbooks.Open(selectedFile, 0)
    Dim active_f As Sheets
    
    Dim i As Integer
    i = 0
    
    For Each rw In f.Worksheets(1).Rows
        Const RES_ACCOUNT = 2
        Const RES_TAX_CODE = 4
        Const RES_DESC = 11
        Const RES_COST_CENTER = 6
        Const DESC = 7 ' Description
        Const GR_ACCOUNT = 5 'GL Account
        Const COST_CENTER = 10
    
        'If IsEmpty(rw.Cells(1).Value) Then Exit For
        If i > 32000 Then Exit For
                
        If IsEmpty(rw.Cells(2).Value) Then
'            MsgBox "Empty!"
        ElseIf rw.Cells(8).Value = 0 Then
'            MsgBox "Credit!"
            ws.Cells(a, RES_DESC) = rw.Cells(DESC).Value
            ws.Cells(a, RES_DESC).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, 3) = rw.Cells(9).Value
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(GR_ACCOUNT).Value
            ws.Cells(a, 1) = 50
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, RES_ACCOUNT) = 212100 Or ws.Cells(a, RES_ACCOUNT) = 212110 Or ws.Cells(a, RES_ACCOUNT) = 214401 Or ws.Cells(a, RES_ACCOUNT) = 212230 Then ws.Cells(a, 1).Value = 31
            ' TAX code is empty
            a = a + 1
        ElseIf IsNumeric(rw.Cells(8).Value) Then
'            MsgBox "Debit!"
            ws.Cells(a, RES_DESC) = rw.Cells(DESC).Value
            ws.Cells(a, RES_DESC).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, 3) = rw.Cells(8).Value
            ws.Cells(a, 1) = 40
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER).Value
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(GR_ACCOUNT).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If special_account(ws.Cells(a, RES_ACCOUNT)) Then ws.Cells(a, 1).Value = 21
'            If ws.Cells(a, 1).Value = 21 Then ws.Cells(a, 4) = "**"
            ' TAX code is empty
            a = a + 1
        Else
'            MsgBox "Empty!"
        End If
        i = i + 1
    Next rw
'    MsgBox f.Worksheets("Sayfa1").Cells(3, 8).Value
    
'    For Each s In f.Worksheets
'        MsgBox s.Name
'    Next s
     
    
End Sub

Sub ClearStatementData()
    For Each Cell In Range("A13:K1000").Cells
        Cell.ClearContents
        Cell.Interior.ColorIndex = 2
        Cell.Borders.ColorIndex = 15
    Next Cell
    ActiveWorkbook.Save
    Range("A13").Select
End Sub

Sub dspi_test()
    Italy
End Sub

Sub Italy()
    Dim filePicker As FileDialog

    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)

    With filePicker

        'setup File Dialog'
        .AllowMultiSelect = False
        .ButtonName = "Select Italy"
        .InitialView = msoFileDialogViewList
        .Title = "Select Italy"
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

        Dim selectedFile As String
        selectedFile = filePicker.SelectedItems(1)

    End If
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim a As Long
'    a = ws.Range("A65000").End(xlUp).Row + 1
    a = 13
    
    Dim vendors_file_path As String
    vendors_file_path = Application.ActiveWorkbook.Path + "\Vendors Italy.xlsx"
    vendors_file_path = vendors_file_path + "bla"
    
    Dim vendors_file As Workbook
    If Dir(vendors_file_path) = "" Then
        MsgBox "File with Italy vendors is not found, please select a file (press 'Cancel' in the next file open dialog to continue without vendors list)"
        Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    
        With filePicker
    
            'setup File Dialog'
            .AllowMultiSelect = False
            .ButtonName = "Select Italy vendors list"
            .InitialView = msoFileDialogViewList
            .Title = "Select list"
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
            Set vendors_file = Workbooks.Open(filePicker.SelectedItems(1))
        End If
    Else
        Set vendors_file = Workbooks.Open(vendors_file_path)
    End If
    
    If Not vendors_file Is Nothing Then
        Dim vendors_list As Worksheet
        Set vendors_list = vendors_file.Worksheets(1)
    End If
    
    Dim f As Workbook
    Set f = Workbooks.Open(selectedFile, 0)
    
    Dim i As Integer
    i = 0
    
    For Each rw In f.Worksheets(1).Rows
        Const RES_ACCOUNT = 2
        Const RES_AMOUNT = 3
        Const RES_TAX_CODE = 4
        Const RES_DESC = 11
        Const RES_COST_CENTER = 6
        Const RES_PK = 1
        Const DESC = 8 ' Description
        Const DESC_1 = 7 ' Description to search in the vendors list
        Const ACCOUNT = 3 'GL Account
        Const COST_CENTER = 5
        
        Const DEBIT = 10
        Const CREDIT = 11
    
        'If IsEmpty(rw.Cells(1).Value) Then Exit For
        If i > 32000 Then Exit For
                
        If Not IsEmpty(rw.Cells(1).Value) Then
'            MsgBox "Empty!"
        ElseIf IsEmpty(rw.Cells(ACCOUNT).Value) Then
            Exit For
        ElseIf rw.Cells(DEBIT).Value = 0 Then
'            MsgBox "Credit!"
            Dim b As String
            Dim c As String
            
            b = rw.Cells(DESC).Value
            c = ws.Cells(a, RES_DESC)
            ws.Cells(a, RES_DESC) = rw.Cells(DESC).Value
            ws.Cells(a, RES_DESC).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, RES_AMOUNT) = rw.Cells(CREDIT).Value
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT).Value
            ws.Cells(a, RES_PK) = 50
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If special_account(ws.Cells(a, RES_ACCOUNT)) Then
                ws.Cells(a, RES_PK).Value = 31
                If Not vendors_file Is Nothing Then
                    For Each vendor_row In vendors_list.Rows
                        If IsEmpty(vendor_row.Cells(2).Value) Then
                            ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                            Exit For
                        End If
                        If InStr(1, rw.Cells(DESC_1).Value, vendor_row.Cells(2).Value, vbTextCompare) > 0 Then
                            ws.Cells(a, RES_ACCOUNT) = vendor_row.Cells(1).Value
                            Exit For
                        End If
                    Next vendor_row
                Else
                   ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                End If
            End If
            ' TAX code is empty
            a = a + 1
        ElseIf IsNumeric(rw.Cells(DEBIT).Value) Then
'            MsgBox "Debit!"
            ws.Cells(a, RES_DESC) = rw.Cells(DESC).Value
            ws.Cells(a, RES_DESC).Font.ColorIndex = rw.Cells(DESC).Font.ColorIndex
            ws.Cells(a, RES_AMOUNT) = rw.Cells(DEBIT).Value
            ws.Cells(a, RES_PK) = 40
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(COST_CENTER).Value
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(ACCOUNT).Value
            If special_account(ws.Cells(a, RES_ACCOUNT)) Then
                ws.Cells(a, RES_PK).Value = 21
                If Not vendors_file Is Nothing Then
                    For Each vendor_row In vendors_list.Rows
                        If IsEmpty(vendor_row.Cells(2).Value) Then
                            ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                            Exit For
                        End If
                        If InStr(1, rw.Cells(DESC_1).Value, vendor_row.Cells(2).Value, vbTextCompare) > 0 Then
                            ws.Cells(a, RES_ACCOUNT) = vendor_row.Cells(1).Value
                            Exit For
                        End If
                    Next vendor_row
                Else
                    ws.Cells(a, RES_ACCOUNT).Interior.ColorIndex = 6
                End If
            End If
            ' TAX code is empty
            a = a + 1
        End If
        i = i + 1
    Next rw
End Sub
