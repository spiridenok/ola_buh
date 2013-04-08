Private credit As Integer
Private account As Integer
Private debit As Integer
Private cost_center As Integer
Private desc As Integer

'Public Function get_account() As Integer
'End Function
'
'Public Function get_debit() As Integer
'End Function
'
'Public Function get_credit() As Integer
'End Function
'
'Public Function get_cost_center() As Integer
'End Function
'
'Public Function get_desc() As Integer
'End Function

Public Function get_col_numbers(country As Integer)
    'Dim c As New CBaseCountry
    If country = 1 Then
        get_account = 4 'GL Account
        get_desc = 7  ' Description
        get_cost_center = 14
        get_debit = 9
        get_credit = 10
    ElseIf country = 3 Then
        get_account = 5 'GL Account
        get_desc = 7  ' Description
        get_cost_center = 10
        get_debit = 8
        get_credit = 9
    End If
    'Set get_col_numbers = c
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
'

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

