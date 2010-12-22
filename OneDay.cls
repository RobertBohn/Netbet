VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit

Private sLedgers() As New Ledger
Private sGameDate As String
Private nLedgers As Long

Public Property Get GameDate()
    GameDate = sGameDate
End Property



Public Sub Populate(ByVal Nbr As String)
    nLedgers = nLedgers + 1
    ReDim Preserve sLedgers(nLedgers)
    Call sLedgers(nLedgers - 1).Populate(Nbr)
    sGameDate = sLedgers(nLedgers - 1).GameDate
End Sub

Private Sub Class_Initialize()
    nLedgers = 0
End Sub



Public Sub Display()
    
    Dim i, wl As Long
    
    'Get Days Win Loss
    wl = 0
    For i = 0 To nLedgers - 1
        wl = wl + sLedgers(i).WinLoss
    Next i

    'Insert Line For the Day
    If wl >= 0 Then
        FormDetail.Outline1.AddItem Format(sGameDate, "ddd m/d/yyyy") & "     +" & wl
    Else
        FormDetail.Outline1.AddItem Format(sGameDate, "ddd m/d/yyyy") & "     " & wl
    End If
    
    'Insert Ledger Objects
    For i = 0 To nLedgers - 1
        Call sLedgers(i).Display
    Next i
    
End Sub

Public Function Indent(ByVal Index As Long, ByVal Indnt As Long) As Long

    Dim i As Long
    Dim levels As Long
    
    FormDetail.Outline1.Indent(Index) = Indnt
    levels = 1
    For i = 0 To nLedgers - 1
        levels = levels + sLedgers(i).Indent(Index + levels, Indnt + 1)
    Next i
    Indent = levels

End Function

Public Function WinLoss() As Long

    Dim i, wl As Long
    
    wl = 0
    For i = 0 To nLedgers - 1
        wl = wl + sLedgers(i).WinLoss()
    Next i
    WinLoss = wl

End Function

Public Function action() As Long

    Dim i, wl As Long
    
    wl = 0
    For i = 0 To nLedgers - 1
        wl = wl + sLedgers(i).action()
    Next i
    action = wl

End Function

Public Function DeleteItem(ByVal level As Long)

    Dim i, Index As Long
    Dim result As Long
    
    Index = 1
    For i = 0 To nLedgers - 1
        result = sLedgers(i).DeleteItem(level + Index)
        If result < 0 Then
            DeleteItem = result
            Exit Function
        End If
        Index = Index + result
    Next i
    DeleteItem = Index

End Function

Public Sub PurgeOldData()

    Dim i As Long
    For i = 0 To nLedgers - 1
        sLedgers(i).PurgeOldData
    Next i

End Sub

Public Function DisplayItem(ByVal level As Long)
    Dim i, Index As Long
    Dim result As Long
    Index = 1
    For i = 0 To nLedgers - 1
        result = sLedgers(i).DisplayItem(level + Index)
        If result < 0 Then
            DisplayItem = result
            Exit Function
        End If
        Index = Index + result
    Next i
    DisplayItem = Index
End Function

Public Function InAction()
    Dim i, action As Long
    
    action = 0
    For i = 0 To nLedgers - 1
        action = action + sLedgers(i).InAction()
    Next i
    InAction = action

End Function
