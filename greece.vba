Sub dspi_test()
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
    
    ' 2nd parameter == 0 suppresses the "Update links" message
    Set f = Workbooks.Open("c:\Users\dspirydz\Documents\ola\Individual Journal Vouchers 02-2013 ORIGINAL.xls", 0)
    Dim active_f As Sheets
    
    Dim i As Integer
    i = 0
    
    For Each rw In f.Worksheets(1).Rows
        Const RES_ACCOUNT = 2
        Const RES_TAX_CODE = 4
        Const RES_DESC = 11
        Const RES_COST_CENTER = 6
        Const GR_DESC = 7 ' Description
        Const GR_ACCOUNT = 5 'GL Account
        Const GR_COST_CENTER = 10
    
        'If IsEmpty(rw.Cells(1).Value) Then Exit For
        If i > 32000 Then Exit For
                
        If IsEmpty(rw.Cells(2).Value) Then
'            MsgBox "Empty!"
        ElseIf rw.Cells(8).Value = 0 Then
'            MsgBox "Credit!"
            ws.Cells(a, RES_DESC) = rw.Cells(GR_DESC).Value
            ws.Cells(a, 3) = rw.Cells(9).Value
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(GR_ACCOUNT).Value
            ws.Cells(a, 1) = 50
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(GR_COST_CENTER).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, RES_ACCOUNT) = 212100 Or ws.Cells(a, RES_ACCOUNT) = 212110 Or ws.Cells(a, RES_ACCOUNT) = 214401 Or ws.Cells(a, RES_ACCOUNT) = 212230 Then ws.Cells(a, 1).Value = 31
            ' TAX code is empty
            a = a + 1
        ElseIf IsNumeric(rw.Cells(8).Value) Then
'            MsgBox "Debit!"
            ws.Cells(a, RES_DESC) = rw.Cells(GR_DESC).Value
            ws.Cells(a, 3) = rw.Cells(8).Value
            ws.Cells(a, 1) = 40
            ws.Cells(a, RES_COST_CENTER) = rw.Cells(GR_COST_CENTER).Value
            ws.Cells(a, RES_ACCOUNT) = rw.Cells(GR_ACCOUNT).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, RES_ACCOUNT) = 212100 Or ws.Cells(a, RES_ACCOUNT) = 212110 Or ws.Cells(a, RES_ACCOUNT) = 214401 Or ws.Cells(a, RES_ACCOUNT) = 212230 Then ws.Cells(a, 1).Value = 21
            If ws.Cells(a, 1).Value = 21 Then ws.Cells(a, 4) = "**"
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
