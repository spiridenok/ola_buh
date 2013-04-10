Const TURKEY = 1
Const ITALY = 2
Const GREECE = 3

Private credit As Integer
Private account As Integer
Private debit As Integer
Private cost_center As Integer
Private desc As Integer
Private alt_desc As Integer

Public Function get_col_numbers(country As Integer)
    'Dim c As New CBaseCountry
    Select Case country
    Case TURKEY
        get_account = 4 'GL Account
        get_desc = 7  ' Description
        get_cost_center = 14
        get_debit = 9
        get_credit = 10
    Case GREECE
        get_account = 5 'GL Account
        get_desc = 7  ' Description
        get_cost_center = 10
        get_debit = 8
        get_credit = 9
    Case ITALY
        get_account = 3 'GL Account
        get_desc = 8  ' Description
        get_cost_center = 5
        get_debit = 10
        get_credit = 11
        get_alt_desc = 7 ' Description to search in the vendors list
    End Select
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

Public Property Get get_alt_desc() As Integer
    get_alt_desc = alt_desc
End Property

Public Property Let get_alt_desc(val As Integer)
    alt_desc = val
End Property

