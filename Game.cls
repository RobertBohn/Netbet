VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit

Private sGameNumber As String
Private sRoadTeam As String
Private sHomeTeam As String
Private sSide As String
Private sLine As String
Private sPrice As String
Private sSport As String
Private sHomeScore As String
Private sRoadScore As String
Private sWagerType As String
Private sGameDate As String
Private sGameTime As String
Private sCancelled As String






Public Property Get GameDate()
    GameDate = sGameDate
End Property



Public Sub Populate(ByVal Nbr As String, ByVal Wgr As String, ByVal Sid As String, ByVal Lin As String, ByVal Pric As String)
    
    Dim lResult As Long
    
    sGameNumber = Nbr
    sWagerType = Wgr
    sSide = Sid
    sLine = Lin
    sPrice = Pric
   
    If Len(sGameNumber) > 0 Then
        lResult = dbSelect("Select * from Schedule Where GameNumber = " & sGameNumber)
        If lResult = 0 Then
            If Not MyTable.EOF Then
                If Not IsNull(MyTable![RoadTeam]) Then sRoadTeam = MyTable![RoadTeam]
                If Not IsNull(MyTable![HomeTeam]) Then sHomeTeam = MyTable![HomeTeam]
                If Not IsNull(MyTable![Sport]) Then sSport = MyTable![Sport]
                If Not IsNull(MyTable![HomeScore]) Then sHomeScore = MyTable![HomeScore]
                If Not IsNull(MyTable![RoadScore]) Then sRoadScore = MyTable![RoadScore]
                If Not IsNull(MyTable![GameDate]) Then sGameDate = Format(MyTable![GameDate], "yyyy/mm/dd")
                If Not IsNull(MyTable![GameDate]) Then sGameTime = MyTable![GameDate]
                If Not IsNull(MyTable![Cancelled]) Then sCancelled = MyTable![Cancelled]
            Else
                sRoadTeam = "Missing Game"
                sHomeTeam = "Missing Game"
                sSport = 1
                sHomeScore = 0
                sRoadScore = 0
                sGameDate = Now()
                sGameTime = Now()
                sCancelled = True
            End If
        End If
        MyTable.Close
    End If

End Sub

Public Sub Display()
    If sCancelled = "C" Then
        FormDetail.Outline1.AddItem sRoadTeam & "    *** POSTPONED ***"
        FormDetail.Outline1.AddItem sHomeTeam
    Else
        If Len(sRoadScore) > 0 Then
            FormDetail.Outline1.AddItem sRoadTeam & "     " & sRoadScore
            FormDetail.Outline1.AddItem sHomeTeam & "     " & sHomeScore
        Else
            FormDetail.Outline1.AddItem sRoadTeam & "     ?"
            FormDetail.Outline1.AddItem sHomeTeam & "     ?"
        End If
    End If
End Sub



Public Function Indent(ByVal Index As Long, ByVal Indnt As Long) As Long
    FormDetail.Outline1.Indent(Index) = Indnt
    FormDetail.Outline1.Indent(Index + 1) = Indnt
    Indent = 2
End Function



Public Function Text() As String
    Dim s, line, total, Price As String

    strGameTotal = ""
    strOverPrice = ""
    strUnderPrice = ""
    strRoadSidePrice = ""
    strHomeSidePrice = ""
    strRoadLine = ""
    strHomeLine = ""
    
    If Len(sGameNumber) = 0 Then
        Text = ""
        Exit Function
    End If
    
    line = sLine
    total = sLine
    Price = sPrice
    
    'Adjust Ticket Display For Football Teasers
    If Val(sWagerType) = BET_TYPE_6TEASER Then
        If sSide = "H" Then line = Val(line) + 6
        If sSide = "R" Then line = Val(line) - 6
        If sSide = "O" Then total = Val(total) - 4
        If sSide = "U" Then total = Val(total) + 4
    End If
    If Val(sWagerType) = BET_TYPE_7TEASER Then
        If sSide = "H" Then line = Val(line) + 7
        If sSide = "R" Then line = Val(line) - 7
        If sSide = "O" Then total = Val(total) - 4
        If sSide = "U" Then total = Val(total) + 4
    End If

    'Adjust Ticket Display For Basketball Teasers
    If Val(sWagerType) = BET_TYPE_4TEASER Then
        If sSide = "H" Then line = Val(line) + 4
        If sSide = "R" Then line = Val(line) - 4
    End If
    
    'NFL and NBA Lines
    s = line
    If Len(s) > 0 And (sSport = "NFL" Or sSport = "NBA") Then
        If Val(s) = 0 Then
            strRoadLine = "Pick"
            strHomeLine = "Pick"
        Else
            If Val(s) < 0 Then s = 0 - Val(s)
            If (Val(s) * 10) Mod 10 = 5 Then
                 s = Val(s) - 0.5
                 If Val(s) = 0 Then s = ""
                 s = s & OneHalf
            End If
            If Val(line) < 0 Then
                strHomeLine = "-" & s
                strRoadLine = "+" & s
            Else
                strRoadLine = "-" & s
                strHomeLine = "+" & s
            End If
        End If
    End If
    
    'Baseball Home Line
    s = Price
    If Len(s) > 0 And sSport = "MLB" Then
        If Val(s) < 0 Then s = 0 - Val(s)
        If (Val(s) * 10) Mod 10 = 5 Then
             s = Val(s) - 0.5
             s = s & OneHalf
        End If
        If Val(Price) < 0 Then
            strHomeLine = "-" & s
        Else
            strHomeLine = "+" & s
        End If
    End If
    
    'Baseball Road Line
    s = Price
    If Len(s) > 0 And sSport = "MLB" Then
        If Val(s) < 0 Then s = 0 - Val(s)
        If (Val(s) * 10) Mod 10 = 5 Then
             s = Val(s) - 0.5
             s = s & OneHalf
        End If
        If Val(Price) < 0 Then
            strRoadLine = "-" & s
        Else
            strRoadLine = "+" & s
        End If
    End If
    
    'NHL Lines
    s = line
    If Len(s) > 0 And sSport = "NHL" Then
        If Val(s) = 0 Then
            strHomeLine = "ev"
        Else
            If Val(s) < 0 Then s = 0 - Val(s)
            If (Val(s) * 10) Mod 10 = 5 Then
                 s = Val(s) - 0.5
                 If Val(s) = 0 Then s = ""
                 s = s & OneHalf
            End If
            If Val(line) < 0 Then
                strHomeLine = "-" & s
            Else
                strHomeLine = "+" & s
            End If
        End If
    
        s = 0 - Val(line) - 0.5
        If Val(s) = 0 Then
            strRoadLine = "ev"
        Else
            If Val(s) < 0 Then s = 0 - Val(s)
            If (Val(s) * 10) Mod 10 = 5 Then
                 s = Val(s) - 0.5
                 If Val(s) = 0 Then s = ""
                 s = s & OneHalf
            End If
            If (0 - Val(line) - 0.5) < 0 Then
                strRoadLine = "-" & s
            Else
                strRoadLine = "+" & s
            End If
        End If
    End If
    
    'Display Side Prices for NHL Only
    If sSport = "NHL" Then
        'Display Road Side Price
        s = Price
        If Len(s) Then
            If Val(s) > 0 Then s = "+" & s
            strRoadSidePrice = s
        End If
                
        'Display Home Side Price
        s = Price
        If Len(s) Then
            If Val(s) > 0 Then s = "+" & s
            strHomeSidePrice = s
        End If
    End If
          
    'Display Totals
    s = total
    If Len(s) > 0 Then
        'Display Total Line
        If (Val(s) * 10) Mod 10 = 5 Then
             s = Val(s) - 0.5
             s = s & OneHalf
        End If
        strGameTotal = s
                    
        'Display Over Total Price
        s = Price
        If Len(s) And Val(s) <> -110 Then
            If Val(s) = 100 Then
                s = "even"
            Else
                If Val(s) > 0 Then s = "+" & s
            End If
            strOverPrice = s
        End If
                    
        'Display Under Total Price
        s = Price
        If Len(s) And Val(s) <> -110 Then
            If Val(s) = 100 Then
                s = "even"
            Else
                If Val(s) > 0 Then s = "+" & s
            End If
            strUnderPrice = s
        End If
    End If
    
    If sSide = "O" Then Text = sRoadTeam & " Over " & strGameTotal & " " & strOverPrice
    If sSide = "U" Then Text = sRoadTeam & " Under " & strGameTotal & " " & strUnderPrice
    If sSide = "R" Then Text = sRoadTeam & " " & strRoadLine & " " & strRoadSidePrice
    If sSide = "H" Then Text = sHomeTeam & " " & strHomeLine & " " & strHomeSidePrice

End Function

Public Function WinLoss() As String

    Dim s As String, line As String, total As String, Price As String

    line = sLine
  
    If sCancelled = "C" Then
        WinLoss = "T"
        Exit Function
    End If
  
    If Len(sRoadScore) = 0 Then
        WinLoss = "?"
        Exit Function
    End If
    
    'Adjust Ticket Display For Football Teasers
    If Val(sWagerType) = BET_TYPE_6TEASER Then
        If sSide = "H" Then line = Val(line) + 6
        If sSide = "R" Then line = Val(line) - 6
        If sSide = "O" Then line = Val(line) - 4
        If sSide = "U" Then line = Val(line) + 4
    End If
    If Val(sWagerType) = BET_TYPE_7TEASER Then
        If sSide = "H" Then line = Val(line) + 7
        If sSide = "R" Then line = Val(line) - 7
        If sSide = "O" Then line = Val(line) - 4
        If sSide = "U" Then line = Val(line) + 4
    End If

    'Adjust Ticket Display For Basketball Teasers
    If Val(sWagerType) = BET_TYPE_4TEASER Then
        If sSide = "H" Then line = Val(line) + 4
        If sSide = "R" Then line = Val(line) - 4
    End If
    
    'NFL and NBA
    s = line
    If Len(s) > 0 And (sSport = "NFL" Or sSport = "NBA") Then
        If sSide = "H" Or sSide = "R" Then
            s = Val(sHomeScore) + Val(line) - Val(sRoadScore)
            If Val(s) = 0 Then WinLoss = "T"
            If Val(s) > 0 Then
                If sSide = "H" Then
                    WinLoss = "W"
                Else
                    WinLoss = "L"
                End If
            End If
            If Val(s) < 0 Then
                If sSide = "R" Then
                    WinLoss = "W"
                Else
                    WinLoss = "L"
                End If
            End If
        Else
            s = Val(sHomeScore) + Val(sRoadScore) - Val(line)
            If Val(s) = 0 Then WinLoss = "T"
            If Val(s) > 0 Then
                If sSide = "O" Then
                    WinLoss = "W"
                Else
                    WinLoss = "L"
                End If
            End If
            If Val(s) < 0 Then
                If sSide = "U" Then
                    WinLoss = "W"
                Else
                    WinLoss = "L"
                End If
            End If
        End If
        Exit Function
    End If
    
    'Baseball Sides
    If (sSide = "R" Or sSide = "H") And sSport = "MLB" Then
        s = Val(sHomeScore) - Val(sRoadScore)
        If Val(s) = 0 Then WinLoss = "T"
        If Val(s) > 0 Then
            If sSide = "H" Then
                WinLoss = "W"
            Else
                WinLoss = "L"
            End If
        End If
        If Val(s) < 0 Then
            If sSide = "R" Then
                WinLoss = "W"
            Else
                WinLoss = "L"
            End If
        End If
        Exit Function
    End If
    
    'Baseball and Hockey Totals
    If (sSide = "O" Or sSide = "U") And (sSport = "MLB" Or sSport = "NHL") Then
        s = Val(sHomeScore) + Val(sRoadScore) - Val(line)
        If Val(s) = 0 Then WinLoss = "T"
        If Val(s) > 0 Then
            If sSide = "O" Then
                WinLoss = "W"
            Else
                WinLoss = "L"
            End If
        End If
        If Val(s) < 0 Then
            If sSide = "U" Then
                WinLoss = "W"
            Else
                WinLoss = "L"
            End If
        End If
        Exit Function
    End If
    
    'Hockey Sides
    If (sSide = "R" Or sSide = "H") And sSport = "NHL" Then
        If sSide = "H" Then
            s = Val(sHomeScore) - Val(sRoadScore) + Val(line)
            If Val(s) = 0 Then WinLoss = "T"
            If Val(s) > 0 Then WinLoss = "W"
            If Val(s) < 0 Then WinLoss = "L"
        End If
        If sSide = "R" Then
            s = Val(sRoadScore) - Val(sHomeScore) - Val(line) - 0.5
            If Val(s) = 0 Then WinLoss = "T"
            If Val(s) > 0 Then WinLoss = "W"
            If Val(s) < 0 Then WinLoss = "L"
        End If
        Exit Function
    End If
  
End Function

Public Property Get Sport()
    Sport = sSport
End Property



Public Property Get Side()
    Side = sSide
End Property






Public Property Get GamePrice()
    GamePrice = sPrice
End Property



Public Property Get GameTime()
    GameTime = sGameTime
End Property


