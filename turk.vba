Sub dspi_test()
    Dim filePicker As FileDialog

'    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
'
'    With filePicker
'
'        'setup File Dialog'
'        .AllowMultiSelect = False
'        .ButtonName = "Select"
'        .InitialView = msoFileDialogViewList
'        .Title = "Select File"
'        .InitialFileName = "test"
'
'        'add filter for all files'
'        With .Filters
'        .Clear
'        .Add "All Files", "*.*"
'        End With
'        .FilterIndex = 1
'
'        'display file dialog box'
'        .Show
'
'     End With
'    If filePicker.SelectedItems.Count > 0 Then
'
'        Dim selectedFile As String
'        selectedFile = filePicker.SelectedItems(1)
'
'    End If
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
    
    Set f = Workbooks.Open("c:\Users\dspirydz\Documents\ola\44_DIALOG 02 2013 LEDGER.xlsx")
    Dim active_f As Sheets
    
    Dim i As Integer
    
    For Each rw In f.Worksheets("Sayfa1").Rows
        If IsEmpty(rw.Cells(1).Value) Then Exit For
        
        If IsEmpty(rw.Cells(8).Value) Then
'            MsgBox "Empty!"
        ElseIf rw.Cells(8).Value = 0 Then
'            MsgBox "Credit!"
            ws.Cells(a, 11) = rw.Cells(6).Value
            ws.Cells(a, 3) = rw.Cells(9).Value
            ws.Cells(a, 2) = rw.Cells(4).Value
            ws.Cells(a, 1) = 50
            ws.Cells(a, 6).Value = rw.Cells(13).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, 2) = 212100 Or ws.Cells(a, 2) = 212110 Or ws.Cells(a, 2) = 214401 Or ws.Cells(a, 2) = 212230 Then ws.Cells(a, 1).Value = 31
            a = a + 1
        ElseIf IsNumeric(rw.Cells(8).Value) Then
'            MsgBox "Debit!"
            ws.Cells(a, 11) = rw.Cells(6).Value
            ws.Cells(a, 3) = rw.Cells(8).Value
            ws.Cells(a, 1) = 40
            ws.Cells(a, 2) = rw.Cells(4).Value
            ws.Cells(a, 6).Value = rw.Cells(13).Value
'            If Not IsEmpty(rw.Cells(15).Value) Then ws.Cells(a, 7).Value = rw.Cells(15).Value
            If ws.Cells(a, 2) = 212100 Or ws.Cells(a, 2) = 212110 Or ws.Cells(a, 2) = 214401 Or ws.Cells(a, 2) = 212230 Then ws.Cells(a, 1).Value = 21
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
