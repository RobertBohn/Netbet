VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit

Private sDays() As New OneDay
Private sAccount As String
Private nDays As Long
Private StartingBalance As Long

Private Sub Class_Initialize()
    nDays = 0
End Sub



Public Sub Populate(ByVal acct As String)
    
    Dim lResult As Long
    Dim i, x, records As Long
    Dim Transactions() As String
    Dim aLedger As New Ledger
    
    sAccount = acct
    FormDetail.Outline1.Clear
    FormDetail.DayList.Clear
    nDays = 0
    ReDim sDays(0)
   
    'Get Starting Balance
    lResult = dbSelect("Select * from Accounts Where Account = '" & sAccount & "'")
    If lResult <> 0 Then
        Exit Sub
    End If
    If Not MyTable.EOF Then
        StartingBalance = MyTable![StartingBalance]
    End If
    MyTable.Close
    
    'Read All Ledger Records For This Account
    records = 0
    lResult = dbSelect("Select Transaction from Ledger Where Account = '" & sAccount & "' AND DELETED = 0")
    If lResult <> 0 Then
        Exit Sub
    End If
    Do Until MyTable.EOF
        records = records + 1
        ReDim Preserve Transactions(records)
        Transactions(records - 1) = MyTable![Transaction]
        MyTable.MoveNext
    Loop
    MyTable.Close
    
    'Populate Ledger Records in the Wagers Array
    For i = 0 To records - 1
        Call aLedger.Populate(Transactions(i))
        For x = 0 To nDays - 1
            If aLedger.GameDate = sDays(x).GameDate Then
                Call sDays(x).Populate(Transactions(i))
                Exit For
            End If
        Next x
        If x = nDays Then
            'Need to add another day
            nDays = nDays + 1
            ReDim Preserve sDays(nDays)
            Call sDays(nDays - 1).Populate(Transactions(i))
        End If
    Next i
    
End Sub

Public Sub Display()
    
    Dim i, x, levels As Long
    Dim balance, action As Long
   
    'Use the Invisible DayList List Box to Sort by Date
    FormDetail.DayList.Clear
    For i = 0 To nDays - 1
        FormDetail.DayList.AddItem sDays(i).GameDate
    Next i
    
    'Insert Starting Balance Line
    balance = AccountSettings(ACT_STARTING_BALANCE)
    If balance >= 0 Then
        FormDetail.Outline1.AddItem "Starting Balance     +" & balance
    Else
        FormDetail.Outline1.AddItem "Starting Balance     " & balance
    End If
    
    'Insert Detail Lines
    For i = 0 To nDays - 1
        FormDetail.DayList.ListIndex = i
        For x = 0 To nDays - 1
            If sDays(x).GameDate = FormDetail.DayList.Text Then
                balance = balance + sDays(x).WinLoss
                Call sDays(x).Display
                Exit For
            End If
        Next x
    Next i
    
    'Insert InAction Amount
    If DEPOSIT_SYSTEM = True Then
        action = InAction()
        balance = balance - action
        FormDetail.Outline1.AddItem "Money In Action     $" & action
    End If
    
    
    'Insert Ending Balance Line
    If balance >= 0 Then
        FormDetail.Outline1.AddItem "Current Balance     +" & balance
    Else
        FormDetail.Outline1.AddItem "Current Balance     " & balance
    End If

    'Do Indents
    levels = 1
    For i = 0 To nDays - 1
        FormDetail.DayList.ListIndex = i
        For x = 0 To nDays - 1
            If sDays(x).GameDate = FormDetail.DayList.Text Then
                levels = levels + sDays(x).Indent(levels, 1)
                Exit For
            End If
        Next x
    Next i

End Sub

Public Function DeleteItem()

    Dim i, x, Index, result As Long

    Index = 1
    For i = 0 To nDays - 1
        FormDetail.DayList.ListIndex = i
        For x = 0 To nDays - 1
            If sDays(x).GameDate = FormDetail.DayList.Text Then
                result = sDays(x).DeleteItem(Index)
                
                If result < 0 Then
                    DeleteItem = result
                    Exit Function
                End If
                
                Index = Index + result
                Exit For
            End If
        Next x
    Next i
    DeleteItem = Index
  
End Function

Public Property Get CurrentBalance()

    Dim i, x As Long
    Dim balance As Long
   
    'Use the Invisible DayList List Box to Sort by Date
    FormDetail.DayList.Clear
    For i = 0 To nDays - 1
        FormDetail.DayList.AddItem sDays(i).GameDate
    Next i
    
    'Insert Starting Balance Line
    balance = AccountSettings(ACT_STARTING_BALANCE)
    If balance >= 0 Then
        FormDetail.Outline1.AddItem "Starting Balance     +" & balance
    Else
        FormDetail.Outline1.AddItem "Starting Balance     " & balance
    End If
    
    'Insert Detail Lines
    For i = 0 To nDays - 1
        FormDetail.DayList.ListIndex = i
        For x = 0 To nDays - 1
            If sDays(x).GameDate = FormDetail.DayList.Text Then
                balance = balance + sDays(x).WinLoss
                Call sDays(x).Display
                Exit For
            End If
        Next x
    Next i
    
    CurrentBalance = balance

End Property



Public Sub PurgeOldData(ByVal CutOffDate As Date)

    Dim i As Long
    Dim OldBalance As Long

    OldBalance = 0
    For i = 0 To nDays - 1
        If DateDiff("d", CutOffDate, sDays(i).GameDate) < 0 Then
            OldBalance = OldBalance + sDays(i).WinLoss
            sDays(i).PurgeOldData
        End If
    Next i

    Call dbUpdate("Update Accounts Set StartingBalance = " & (StartingBalance + OldBalance) & " Where Account = '" & sAccount & "'")

End Sub

Public Property Get WinLoss(ByVal thedate As Date) As String

    Dim i As Long
    WinLoss = ""
    For i = 0 To nDays - 1
        If DateDiff("d", thedate, sDays(i).GameDate) = 0 Then
            WinLoss = sDays(i).WinLoss
            Exit For
        End If
    Next i
    
End Property



Public Property Get action(ByVal thedate As Date) As String
    Dim i As Long
    action = ""
    For i = 0 To nDays - 1
        If DateDiff("d", thedate, sDays(i).GameDate) = 0 Then
            action = sDays(i).action
            Exit For
        End If
    Next i
End Property



Public Function DisplayItem()
    Dim i, x, Index, result As Long
    
    Index = 1
    For i = 0 To nDays - 1
        FormDetail.DayList.ListIndex = i
        For x = 0 To nDays - 1
            If sDays(x).GameDate = FormDetail.DayList.Text Then
                result = sDays(x).DisplayItem(Index)
                
                If result < 0 Then
                    DisplayItem = result
                    Exit Function
                End If
                
                Index = Index + result
                Exit For
            End If
        Next x
    Next i
    DisplayItem = Index
End Function




Public Property Get InAction() As Long
    Dim i, action As Long
    
    action = 0
    For i = 0 To nDays - 1
        action = action + sDays(i).InAction
    Next i

    InAction = 0 - action

End Property


