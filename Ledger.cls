VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit

Private sEntryTime As String
Private sTransaction As String
Private sTimeStamp As String
Private sType As String
Private sAmount As String
Private sOdds As String
Private sGames(4) As New Game
Private sLedgerDate As String







Public Sub Populate(ByVal Nbr As String)

    Dim lResult As Long
    Dim i As Long
    Dim gms(4, 4) As String
    
    sTransaction = Nbr
    
    lResult = dbSelect("Select * from Ledger Where Transaction = " & sTransaction)
    If lResult = 0 Then
        If Not MyTable.EOF Then
            If Not IsNull(MyTable![TimeStamp]) Then sTimeStamp = Format(MyTable![TimeStamp], "yyyy/mm/dd")
            If Not IsNull(MyTable![TimeStamp]) Then sEntryTime = MyTable![TimeStamp]
            If Not IsNull(MyTable![Type]) Then sType = MyTable![Type]
            If Not IsNull(MyTable![amount]) Then sAmount = MyTable![amount]
            If Not IsNull(MyTable![odds]) Then sOdds = MyTable![odds]
            
            If Not IsNull(MyTable![Game1]) Then gms(0, 0) = MyTable![Game1]
            If Not IsNull(MyTable![Game2]) Then gms(1, 0) = MyTable![Game2]
            If Not IsNull(MyTable![Game3]) Then gms(2, 0) = MyTable![Game3]
            If Not IsNull(MyTable![Game4]) Then gms(3, 0) = MyTable![Game4]
            
            If Not IsNull(MyTable![Side1]) Then gms(0, 1) = MyTable![Side1]
            If Not IsNull(MyTable![Side2]) Then gms(1, 1) = MyTable![Side2]
            If Not IsNull(MyTable![Side3]) Then gms(2, 1) = MyTable![Side3]
            If Not IsNull(MyTable![Side4]) Then gms(3, 1) = MyTable![Side4]
            
            If Not IsNull(MyTable![line1]) Then gms(0, 2) = MyTable![line1]
            If Not IsNull(MyTable![Line2]) Then gms(1, 2) = MyTable![Line2]
            If Not IsNull(MyTable![Line3]) Then gms(2, 2) = MyTable![Line3]
            If Not IsNull(MyTable![Line4]) Then gms(3, 2) = MyTable![Line4]
        
            If Not IsNull(MyTable![Price1]) Then gms(0, 3) = MyTable![Price1]
            If Not IsNull(MyTable![Price2]) Then gms(1, 3) = MyTable![Price2]
            If Not IsNull(MyTable![Price3]) Then gms(2, 3) = MyTable![Price3]
            If Not IsNull(MyTable![Price4]) Then gms(3, 3) = MyTable![Price4]
        End If
    End If
    MyTable.Close
    
    For i = 0 To 3
        Call sGames(i).Populate(gms(i, 0), sType, gms(i, 1), gms(i, 2), gms(i, 3))
    Next i

    'For Adjustments or Payments use the transaction date
    If Val(sType) = BET_TYPE_ADJUSTMENT Or Val(sType) = BET_TYPE_PAYMENT Then
        sLedgerDate = sTimeStamp
        Exit Sub
    End If
    
    'Straight bets use the game date
    If Val(sType) = BET_TYPE_STRAIGHT Then
        sLedgerDate = sGames(0).GameDate
        Exit Sub
    End If
    
    'All others use max game date
    sLedgerDate = sGames(0).GameDate
    For i = 1 To 3
        If Len(sGames(i).Text) = 0 Then Exit For
            If DateDiff("d", sLedgerDate, sGames(i).GameDate) >= 0 Then
            sLedgerDate = sGames(i).GameDate
        End If
    Next i

End Sub

Public Property Get GameDate()
    GameDate = sLedgerDate
End Property



Public Sub Display()

    Dim s, sResult As String
    Dim i As Long

    If Val(sType) = BET_TYPE_ADJUSTMENT Then
        If Val(sAmount) >= 0 Then
            FormDetail.Outline1.AddItem "Adjustment     +" & sAmount
        Else
            FormDetail.Outline1.AddItem "Adjustment     " & sAmount
        End If
        Exit Sub
    End If

    If Val(sType) = BET_TYPE_PAYMENT And Val(sAmount) < 0 Then
        FormDetail.Outline1.AddItem "Payment     " & sAmount
        Exit Sub
    End If

    If Val(sType) = BET_TYPE_PAYMENT And Val(sAmount) > 0 Then
        FormDetail.Outline1.AddItem "Collection     +" & sAmount
        Exit Sub
    End If

    If Val(sType) = BET_TYPE_STRAIGHT Then
        If sGames(0).WinLoss = "?" Then
            FormDetail.Outline1.AddItem "$" & sAmount & " Straight Bet  " & sGames(0).Text & "    ?"
        Else
            If WinLoss >= 0 Then
                FormDetail.Outline1.AddItem "$" & sAmount & " Straight Bet  " & sGames(0).Text & "    +" & WinLoss
            Else
                FormDetail.Outline1.AddItem "$" & sAmount & " Straight Bet  " & sGames(0).Text & "    " & WinLoss
            End If
        End If
        Call sGames(0).Display
        Exit Sub
    End If

    If Val(sType) = BET_TYPE_PARLAY _
    Or Val(sType) = BET_TYPE_4TEASER _
    Or Val(sType) = BET_TYPE_6TEASER _
    Or Val(sType) = BET_TYPE_7TEASER Then
        sResult = ""
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "?" Then sResult = "?"
        Next i
    
        s = "$" & sAmount & " " & i
        
        Select Case Val(sType)
        Case BET_TYPE_PARLAY
            s = s & "-Team Parlay     "
        Case BET_TYPE_4TEASER
            s = s & "-Team 4pt Teaser     "
        Case BET_TYPE_6TEASER
            s = s & "-Team 6pt Teaser     "
        Case BET_TYPE_7TEASER
            s = s & "-Team 7pt Teaser     "
        End Select
        
        If sResult = "?" Then
            s = s & "?"
        Else
            If WinLoss >= 0 Then
                s = s & "+" & WinLoss
            Else
                s = s & WinLoss
            End If
        End If
        FormDetail.Outline1.AddItem s
        
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
            s = sGames(i).Text
            
            If sGames(i).WinLoss = "W" Then s = s & "     Win"
            If sGames(i).WinLoss = "L" Then s = s & "     Lose"
            If sGames(i).WinLoss = "T" Then s = s & "     Push"
            If sGames(i).WinLoss = "?" Then s = s & "     ?"
            
            FormDetail.Outline1.AddItem s
            Call sGames(i).Display
        Next i
        Exit Sub
    End If
    
End Sub



Public Function Indent(ByVal Index As Long, ByVal Indnt As Long) As Long

    Dim i

    If Val(sType) = BET_TYPE_ADJUSTMENT Or Val(sType) = BET_TYPE_PAYMENT Then
        FormDetail.Outline1.Indent(Index) = Indnt
        Indent = 1
        Exit Function
    End If

    If Val(sType) = BET_TYPE_STRAIGHT Then
        FormDetail.Outline1.Indent(Index) = Indnt
        Indent = sGames(0).Indent(Index + 1, Indnt + 1) + 1
        Exit Function
    End If

    If Val(sType) = BET_TYPE_PARLAY _
    Or Val(sType) = BET_TYPE_4TEASER _
    Or Val(sType) = BET_TYPE_6TEASER _
    Or Val(sType) = BET_TYPE_7TEASER Then
        FormDetail.Outline1.Indent(Index) = Indnt
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
            FormDetail.Outline1.Indent((i * 3) + Index + 1) = Indnt + 1
            Call sGames(i).Indent((i * 3) + Index + 2, Indnt + 2)
        Next i
        Indent = (i * 3) + 1
        Exit Function
    End If

End Function

Public Function WinLoss() As Long

    Dim result, odds As String
    Dim i, w, l, t As Long
    Dim hockey, football, baseball As Long

    If Val(sType) = BET_TYPE_ADJUSTMENT Or Val(sType) = BET_TYPE_PAYMENT Then
        WinLoss = Val(sAmount)
        Exit Function
    End If

    If Val(sType) = BET_TYPE_STRAIGHT Then
        WinLoss = 0
        result = sGames(0).WinLoss
        If result = "W" Or result = "L" Then
            
            'Football and Basketball
            If sGames(0).Sport = "NFL" Or sGames(0).Sport = "NBA" Then
                If result = "W" Then
                    WinLoss = Val(sAmount)
                Else
                    WinLoss = 0 - ((Val(sAmount) * (0 - Val(sOdds))) / 100)
                End If
            End If
             
            'Baseball Sides
            If sGames(0).Sport = "MLB" Then
                If sGames(0).Side = "H" Or sGames(0).Side = "R" Then
                    odds = sGames(0).GamePrice
                    If result = "W" Then
                        If Val(odds) < 0 Then
                            WinLoss = Val(sAmount)
                        Else
                            WinLoss = (Val(sAmount) * Val(odds)) / 5
                        End If
                    Else
                        If Val(odds) > 0 Then
                            WinLoss = 0 - Val(sAmount)
                        Else
                            WinLoss = (Val(sAmount) * Val(odds)) / 5
                        End If
                    End If
                End If
            End If
           
           'Baseball and Hockey Totals
            If sGames(0).Sport = "MLB" Or sGames(0).Sport = "NHL" Then
                If sGames(0).Side = "O" Or sGames(0).Side = "U" Then
                    odds = sGames(0).GamePrice
                    If result = "W" Then
                        If Val(odds) < 0 Then
                            WinLoss = Val(sAmount)
                        Else
                            WinLoss = (Val(sAmount) * Val(odds)) / 100
                        End If
                    Else
                        If Val(odds) > 0 Then
                            WinLoss = 0 - Val(sAmount)
                        Else
                            WinLoss = (Val(sAmount) * Val(odds)) / 100
                        End If
                    End If
                End If
           End If
       
           'Hockey Sides
            If sGames(0).Sport = "NHL" Then
                If sGames(0).Side = "H" Or sGames(0).Side = "R" Then
                    odds = sGames(0).GamePrice
                    If result = "W" Then
                        If Val(odds) < 0 Then
                            WinLoss = Val(sAmount)
                        Else
                            WinLoss = (Val(sAmount) * Val(odds)) / 100
                        End If
                    Else
                        If Val(odds) > 0 Then
                            WinLoss = 0 - Val(sAmount)
                        Else
                            WinLoss = (Val(sAmount) * Val(odds)) / 100
                        End If
                    End If
                End If
           End If
        End If
        Exit Function
    End If
    
    
    If Val(sType) = BET_TYPE_PARLAY Then
        WinLoss = w = l = 0
        hockey = football = baseball = 0
        'Count winners and losers and sports
        For i = 0 To 3
            'Any ? Game is a ? Parlay
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "?" Then Exit Function
            If sGames(i).WinLoss = "W" Then w = w + 1
            If sGames(i).WinLoss = "L" Then l = l + 1
            If sGames(i).Sport = "MLB" Then baseball = baseball + 1
            If sGames(i).Sport = "NHL" Then hockey = hockey + 1
            If sGames(i).Sport = "NFL" Then football = football + 1
            If sGames(i).Sport = "NBA" Then football = football + 1
        Next i
        If l > 0 Then
            'Any loss is a losing Parlay
            WinLoss = 0 - Val(sAmount)
            Exit Function
        End If
        If w = 0 Then
            'All ties is a push
            WinLoss = 0
            Exit Function
        End If
        
        If baseball = 0 And hockey = 0 Then
            'NFL or NBA regular parlay
            Select Case w
            Case 1
                WinLoss = (Val(sAmount) * 100) / (0 - Val(AccountSettings(ACT_STRAIGHT_BET)))
            Case 2
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_PARLAY2))) / 100
            Case 3
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_PARLAY3))) / 100
            Case 4
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_PARLAY4))) / 100
            End Select
            Exit Function
        End If
       
        'hockey and baseball requires calculating
        result = Val(sAmount)
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "W" Then
                odds = sGames(i).GamePrice
                'Default price is 11-10
                If Len(odds) = 0 Then odds = "-110"
                'Baseball is a nickel line
                If sGames(i).Sport = "MLB" Then odds = Val(odds) * 20
            
                If Val(odds) >= 0 Then
                    'dog
                    result = Val(result) + ((Val(result) * Val(odds)) / 100)
                Else
                    'favorite
                    result = Val(result) + ((Val(result) * 100) / (0 - Val(odds)))
                End If
            End If
        Next i
        'baseball only - deduct calc fee
        If football = 0 And hockey = 0 Then
            If w > 1 Then
                result = Val(result) - (Val(result) * Val(AccountSettings(ACT_PARLAY_FEE)))
            End If
        End If
        'Deduct original bet
        result = Val(result) - Val(sAmount)
        WinLoss = Val(result)
        Exit Function
    End If
    
    
    If Val(sType) = BET_TYPE_4TEASER Then
        WinLoss = 0
        w = 0
        l = 0
        t = 0
        'Count winners and losers
        For i = 0 To 3
            'Any ? Game is a ? Teaser
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "?" Then Exit Function
            If sGames(i).WinLoss = "W" Then w = w + 1
            If sGames(i).WinLoss = "L" Then l = l + 1
            If sGames(i).WinLoss = "T" Then t = t + 1
        Next i
        If l > 0 Then
            'Any loss is a losing Teaser
            If i = 2 And Val(AccountSettings(ACT_4PT_TEASER2)) < 0 Then
                WinLoss = 0 - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_4PT_TEASER2)))) / 100)
            Else
                WinLoss = 0 - Val(sAmount)
            End If
            Exit Function
        End If
        'deposit system ties -> push
        If ALL_TIES_PUSH = False Then
            If t > 0 And Val(AccountSettings(ACT_4PT_TEASER)) = 0 Then
                'Ties Lose
                If i = 2 And Val(AccountSettings(ACT_4PT_TEASER2)) < 0 Then
                    WinLoss = 0 - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_4PT_TEASER2)))) / 100)
                Else
                    WinLoss = 0 - Val(sAmount)
                End If
                Exit Function
            End If
        End If
        If w = 3 Then
            'Winning 4pt 3-Team Teaser
            If Val(AccountSettings(ACT_4PT_TEASER3)) >= 0 Then
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_4PT_TEASER3))) / 100
            Else
                WinLoss = Val(sAmount)
            End If
        End If
        If w = 2 Then
            'Winning 4pt 2-Team Teaser
            If Val(AccountSettings(ACT_4PT_TEASER2)) >= 0 Then
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_4PT_TEASER2))) / 100
            Else
                WinLoss = Val(sAmount)
            End If
        End If
        If w = 1 Or w = 0 Then
            '4pt 2 Team Teaser With a Tie Pushes
            WinLoss = 0
        End If
        Exit Function
    End If
    
    
    If Val(sType) = BET_TYPE_6TEASER Then
        WinLoss = 0
        w = 0
        l = 0
        t = 0
        'Count winners and losers
        For i = 0 To 3
            'Any ? Game is a ? Teaser
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "?" Then Exit Function
            If sGames(i).WinLoss = "W" Then w = w + 1
            If sGames(i).WinLoss = "L" Then l = l + 1
            If sGames(i).WinLoss = "T" Then t = t + 1
        Next i
        If l > 0 Then
            'Any loss is a losing Teaser
            If i = 2 And Val(AccountSettings(ACT_6PT_TEASER2)) < 0 Then
                WinLoss = 0 - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_6PT_TEASER2)))) / 100)
            Else
                WinLoss = 0 - Val(sAmount)
            End If
            Exit Function
        End If
        
        'deposit system ties -> push
        If ALL_TIES_PUSH = False Then
            If t > 0 And Val(AccountSettings(ACT_6PT_TEASER)) = 0 Then
                'Ties Lose
                If i = 2 And Val(AccountSettings(ACT_6PT_TEASER2)) < 0 Then
                    WinLoss = 0 - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_6PT_TEASER2)))) / 100)
                Else
                    WinLoss = 0 - Val(sAmount)
                End If
                Exit Function
            End If
        End If
'        If i > w And Val(AccountSettings(ACT_6PT_TEASER2)) < 0 Then
'            'If Player lays more than even money, ties lose
'            'Any Tie is a Loss (GamesBet < GamesWon)
'            WinLoss = 0 - Val(sAmount)
'            Exit Function
'        End If
        If w = 3 Then
            'Winning 6pt 3-Team Teaser
            If Val(AccountSettings(ACT_6PT_TEASER3)) >= 0 Then
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_6PT_TEASER3))) / 100
            Else
                WinLoss = Val(sAmount)
            End If
        End If
        If w = 2 Then
            'Winning 6pt 2-Team Teaser
            If Val(AccountSettings(ACT_6PT_TEASER2)) >= 0 Then
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_6PT_TEASER2))) / 100
            Else
                WinLoss = Val(sAmount)
            End If
        End If
        If w = 1 Or w = 0 Then
            'Ties down to 0 or 1 game is a push
            WinLoss = 0
        End If
        Exit Function
    End If
    
    
    If Val(sType) = BET_TYPE_7TEASER Then
        WinLoss = 0
        w = 0
        l = 0
        t = 0
        'Count winners and losers
        For i = 0 To 3
            'Any ? Game is a ? Teaser
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "?" Then Exit Function
            If sGames(i).WinLoss = "W" Then w = w + 1
            If sGames(i).WinLoss = "L" Then l = l + 1
            If sGames(i).WinLoss = "T" Then t = t + 1
        Next i
        If l > 0 Then
            'Any loss is a losing Teaser
            If i = 2 And Val(AccountSettings(ACT_7PT_TEASER2)) < 0 Then
                WinLoss = 0 - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_7PT_TEASER2)))) / 100)
            Else
                WinLoss = 0 - Val(sAmount)
            End If
            Exit Function
        End If
        'deposit system ties -> push
        If ALL_TIES_PUSH = False Then
            If t > 0 And Val(AccountSettings(ACT_7PT_TEASER)) = 0 Then
                'Ties Lose
                If i = 2 And Val(AccountSettings(ACT_7PT_TEASER2)) < 0 Then
                    WinLoss = 0 - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_7PT_TEASER2)))) / 100)
                Else
                    WinLoss = 0 - Val(sAmount)
                End If
                Exit Function
            End If
        End If
'        If i > w Then
'            'Any Tie is a Loss (GamesBet < GamesWon)
'            WinLoss = 0 - Val(sAmount)
'            Exit Function
'        End If
        If w = 3 Then
            'Winning 7pt 3-Team Teaser
            If Val(AccountSettings(ACT_7PT_TEASER3)) >= 0 Then
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_7PT_TEASER3))) / 100
            Else
                WinLoss = Val(sAmount)
            End If
        End If
        If w = 2 Then
            'Winning 7pt 2-Team Teaser
            If Val(AccountSettings(ACT_7PT_TEASER2)) >= 0 Then
                WinLoss = (Val(sAmount) * Val(AccountSettings(ACT_7PT_TEASER2))) / 100
            Else
                WinLoss = Val(sAmount)
            End If
        End If
        If w = 1 Or w = 0 Then
            'Ties down to 0 or 1 game is a push
            WinLoss = 0
        End If
        Exit Function
    End If
    
End Function



Public Function action() As Long

    Dim result, odds As String
    Dim i, w, l, t As Long
    Dim hockey, football, baseball As Long

    If Val(sType) = BET_TYPE_PAYMENT Then
        action = 0
    Else
        action = WinLoss
    End If

End Function




Public Function DeleteItem(ByVal level As Long)

    Dim i As Long
    Dim lResult As Long
    
    If Val(sType) = BET_TYPE_ADJUSTMENT Or Val(sType) = BET_TYPE_PAYMENT Then
        DeleteItem = 1
        Exit Function
    End If

    If level = FormDetail.Outline1.ListIndex Then
        'Check Game Dates
        If WAGER_AFTER_KICKOFF = False Then
            For i = 0 To 3
                If Len(sGames(i).Text) = 0 Then Exit For
                If DateDiff("s", Now, sGames(i).GameTime) < 0 Then
                    lResult = MessageBox(0, "You can not delete a wager after a game has started.", "Deleting Wager", MB_OK)
                    GoTo KeepGoing
                End If
            Next i
        End If
        
        'Confirm Deletion
        lResult = MessageBox(0, "Are you sure you want to delete this wager?", "Deleting Wager", MB_YESNO)
        If lResult = IDYES Then
            dbUpdate ("Update Ledger Set Deleted = 1 Where Transaction = " & sTransaction)
            DeleteItem = -1
            Exit Function
        End If
    End If

KeepGoing:

    If Val(sType) = BET_TYPE_STRAIGHT Then
        DeleteItem = 3
        Exit Function
    End If

    If Val(sType) = BET_TYPE_PARLAY _
    Or Val(sType) = BET_TYPE_4TEASER _
    Or Val(sType) = BET_TYPE_6TEASER _
    Or Val(sType) = BET_TYPE_7TEASER Then
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
        Next i
        DeleteItem = (i * 3) + 1
        Exit Function
    End If

End Function

Public Sub PurgeOldData()
    Dim i, notdecided As Long
            
    notdecided = 0
    For i = 0 To 3
        If Len(sGames(i).Text) = 0 Then Exit For
        If sGames(i).WinLoss = "?" Then notdecided = 1
    Next i
    
    If notdecided = 0 Then
        Call dbDelete("Delete from Ledger Where Transaction = " & sTransaction)
    End If
    
End Sub



Public Function DisplayItem(ByVal level As Long)
    Dim i As Long
    Dim lResult As Long

    If level = FormDetail.Outline1.ListIndex Then
        Call MessageBox(0, "Transaction #" & sTransaction & "  " & sEntryTime, "Wager Detail", MB_OK)
        DisplayItem = -1
        Exit Function
    End If

    If Val(sType) = BET_TYPE_ADJUSTMENT Or Val(sType) = BET_TYPE_PAYMENT Then
        DisplayItem = 1
    End If

    If Val(sType) = BET_TYPE_STRAIGHT Then
        DisplayItem = 3
    End If

    If Val(sType) = BET_TYPE_PARLAY _
    Or Val(sType) = BET_TYPE_4TEASER _
    Or Val(sType) = BET_TYPE_6TEASER _
    Or Val(sType) = BET_TYPE_7TEASER Then
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
        Next i
        DisplayItem = (i * 3) + 1
    End If

End Function




Public Function InAction() As Long
    Dim odds, result As String
    Dim i, teams, notdecided As Long

    InAction = 0

    If Val(sType) = BET_TYPE_ADJUSTMENT Or Val(sType) = BET_TYPE_PAYMENT Then
        Exit Function
    End If

    If Val(sType) = BET_TYPE_STRAIGHT Then
        result = sGames(0).WinLoss
        If result = "?" Then
            
            'Football and Basketball
            If sGames(0).Sport = "NFL" Or sGames(0).Sport = "NBA" Then
                InAction = 0 - ((Val(sAmount) * (0 - Val(sOdds))) / 100)
                Exit Function
            End If
             
            'Baseball Sides
            If sGames(0).Sport = "MLB" Then
                If sGames(0).Side = "H" Or sGames(0).Side = "R" Then
                    odds = sGames(0).GamePrice
                    If Val(odds) > 0 Then
                        InAction = 0 - Val(sAmount)
                    Else
                        InAction = (Val(sAmount) * Val(odds)) / 5
                    End If
                    Exit Function
                End If
            End If
           
           'Baseball and Hockey Totals
            If sGames(0).Sport = "MLB" Or sGames(0).Sport = "NHL" Then
                If sGames(0).Side = "O" Or sGames(0).Side = "U" Then
                    odds = sGames(0).GamePrice
                    If Val(odds) > 0 Then
                        InAction = 0 - Val(sAmount)
                    Else
                        InAction = (Val(sAmount) * Val(odds)) / 100
                    End If
                    Exit Function
                End If
           End If
       
           'Hockey Sides
            If sGames(0).Sport = "NHL" Then
                If sGames(0).Side = "H" Or sGames(0).Side = "R" Then
                    odds = sGames(0).GamePrice
                    If Val(odds) > 0 Then
                        InAction = 0 - Val(sAmount)
                    Else
                        InAction = (Val(sAmount) * Val(odds)) / 100
                    End If
                    Exit Function
                End If
           End If
        End If
        Exit Function
    End If
    
    If Val(sType) = BET_TYPE_PARLAY Or _
    Val(sType) = BET_TYPE_4TEASER Or _
    Val(sType) = BET_TYPE_6TEASER Or _
    Val(sType) = BET_TYPE_7TEASER Then
        'Look for '?'
        teams = 0
        notdecided = 0
        For i = 0 To 3
            If Len(sGames(i).Text) = 0 Then Exit For
            If sGames(i).WinLoss = "?" Then notdecided = 1
            teams = teams + 1
        Next i
        
        If notdecided = 1 Then
            If Val(sType) = BET_TYPE_PARLAY Then
                InAction = 0 - Val(sAmount)
            End If
            If Val(sType) = BET_TYPE_4TEASER Then
                If teams = 2 And Val(AccountSettings(ACT_4PT_TEASER2)) < 0 Then
                    InAction = InAction - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_4PT_TEASER2)))) / 100)
                Else
                    InAction = 0 - Val(sAmount)
                End If
            End If
            If Val(sType) = BET_TYPE_6TEASER Then
                If teams = 2 And Val(AccountSettings(ACT_6PT_TEASER2)) < 0 Then
                    InAction = InAction - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_6PT_TEASER2)))) / 100)
                Else
                    InAction = 0 - Val(sAmount)
                End If
            End If
            If Val(sType) = BET_TYPE_7TEASER Then
                If teams = 2 And Val(AccountSettings(ACT_7PT_TEASER2)) < 0 Then
                    InAction = InAction - ((Val(sAmount) * (0 - Val(AccountSettings(ACT_7PT_TEASER2)))) / 100)
                Else
                    InAction = 0 - Val(sAmount)
                End If
            End If
        End If
    End If

End Function
