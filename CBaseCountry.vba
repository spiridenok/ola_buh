Const TURKEY = 1
Const ITALY = 2
Const GREECE = 3

Private credit As Integer
Private account As Integer
Private debit As Integer
Private cost_center As Integer
Private desc As Integer
Private alt_desc As Integer

Public Function get_col_numbers(country As Integer, rows As Range)
    Dim debit_found As Boolean
    Dim credit_found As Boolean
    debit_found = False
    credit_found = False
    Dim found_row As Range
    
    For Each rw In rows
        For Each cell In rw.Columns
'            If UCase(cell.Value.Trim) = "DEBIT" Then debit_found = True
'            If UCase(cell.Value.Trim) = "CREDIT" Then credit_found = True
            If InStr(1, cell, "DEBIT", vbTextCompare) > 0 Then
                debit_found = True
                get_debit = cell.Column
            End If
            If InStr(1, cell, "CREDIT", vbTextCompare) > 0 Then
                credit_found = True
                get_credit = cell.Column
            End If
        If cell.Column > 20 Then Exit For
        Next cell
        
        If debit_found And credit_found Then
            'MsgBox "Found on line " + CStr(rw.Row)
            Set found_row = rw
            Exit For
        End If
        
        If rw.Row > 10 Then
            If country <> ITALY Then
                MsgBox "Could not find debet/credit in the first 10 rows of a file. Make sure the right file is used."
            Else
                Set found_row = Range("A1:M1000")
            End If
            Exit For
        End If
    Next rw
    
    For Each cell In found_row.Columns
    Select Case country
        Case TURKEY
            If InStr(1, cell, "DIALOG", vbTextCompare) > 0 And InStr(1, cell, "ACCOUNT", vbTextCompare) > 0 Then get_account = cell.Column
            If InStr(1, cell, "cost", vbTextCompare) > 0 And (InStr(1, cell, "center", vbTextCompare) > 0 Or InStr(1, cell, "centre", vbTextCompare) > 0) And InStr(1, cell, "code", vbTextCompare) > 0 Then get_cost_center = cell.Column
            If InStr(1, cell, "account", vbTextCompare) > 0 And InStr(1, cell, "name", vbTextCompare) > 0 Then get_desc = cell.Column
        Case GREECE
            If InStr(1, cell, "GL", vbTextCompare) > 0 And InStr(1, cell, "ACCOUNT", vbTextCompare) > 0 Then get_account = cell.Column
            If InStr(1, cell, "cost", vbTextCompare) > 0 And InStr(1, cell, "center", vbTextCompare) > 0 Then get_cost_center = cell.Column
            If InStr(1, cell, "description", vbTextCompare) > 0 And Len(cell) = Len("Description") Then get_desc = cell.Column
        Case ITALY
            get_debit = 3
            get_credit = 3
            get_account = 2
            get_cost_center = 6
            get_desc = 11
            Exit For
'            If InStr(1, cell, "GL", vbTextCompare) > 0 And InStr(1, cell, "ACCOUNT", vbTextCompare) > 0 Then get_account = cell.Column
'            If InStr(1, cell, "cost", vbTextCompare) > 0 And InStr(1, cell, "center", vbTextCompare) > 0 Then get_cost_center = cell.Column
'            If InStr(1, cell, "description", vbTextCompare) > 0 And InStr(1, cell, "2", vbTextCompare) > 0 Then get_desc = cell.Column
'            If InStr(1, cell, "description", vbTextCompare) > 0 And InStr(1, cell, "1", vbTextCompare) > 0 Then get_alt_desc = cell.Column
    End Select
    Next cell
    
    If get_account = 0 Then MsgBox "Did not find account"
    If get_cost_center = 0 Then MsgBox "Did not find cost center"
    If get_desc = 0 Then MsgBox "Did not find description"
    
End Function

Property Get get_credit() As Integer
    get_credit = credit
End Property
Property Let get_credit(val As Integer)
    credit = val
End Property

Public Property Get get_account() As Integer
    get_account = account
End Property
Public Property Let get_account(val As Integer)
    account = val
End Property

'End Function
'
Public Property Get get_debit() As Integer
    get_debit = debit
    
End Property
Public Property Let get_debit(val As Integer)
    debit = val
    
End Property

Public Property Get get_cost_center() As Integer
    get_cost_center = cost_center
End Property

Public Property Let get_cost_center(val As Integer)
    cost_center = val
End Property
'End Function
'
Public Property Get get_desc() As Integer
    get_desc = desc
End Property

Public Property Let get_desc(val As Integer)
    desc = val
End Property

Public Property Get get_alt_desc() As Integer
    get_alt_desc = alt_desc
End Property

Public Property Let get_alt_desc(val As Integer)
    alt_desc = val
End Property

Public Sub fill_common_data(country As Integer, cost_center As Variant, ByRef account_cell As Range, ByRef cost_center_cell As Range, Optional ByRef tax_code_cell As Range)
    Select Case country
        Case GREECE
            cost_center_cell = cost_center
            If account_cell = 212100 Or account_cell = 212110 Then account_cell = 8809
            If account_cell = 214401 Then account_cell = 2413
            If (account_cell Like "5*" Or account_cell Like "4*") And IsEmpty(cost_center_cell) Then
                cost_center_cell.Interior.ColorIndex = 3
            End If
            If Not account_cell Like "5*" And Not account_cell Like "4*" And Not IsEmpty(cost_center_cell) Then
                cost_center_cell.Interior.ColorIndex = 3
            End If
            
        Case ITALY
            cost_center_cell = cost_center
            If (account_cell Like "6*" Or account_cell Like "5*" Or account_cell Like "4*") And IsEmpty(cost_center_cell) Then
                cost_center_cell.Interior.ColorIndex = 3
            End If
            If Not account_cell Like "6*" And Not account_cell Like "5*" And Not account_cell Like "4*" And Not IsEmpty(cost_center_cell) Then
                cost_center_cell.Interior.ColorIndex = 3
            End If
        
        Case TURKEY
            If account_cell Like "5*" Then
                tax_code_cell = "V0"
                cost_center_cell = cost_center
            End If
    End Select
End Sub
