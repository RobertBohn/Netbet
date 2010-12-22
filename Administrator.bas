Attribute VB_Name = "Module1"
'All variables MUST be defined
Option Explicit

'Constants
Public Const NFL_TAB = 0           'Football Tab Index
Public Const NBA_TAB = 1           'Basketball Tab Index
Public Const MLB_TAB = 2           'Baseball Tab Index
Public Const NHL_TAB = 3           'Hockey Tab Index
Public Const SPORTS = 4            'Number of Sports

Public Const ACT_ACCOUNT = 0        'Account Settings Index
Public Const ACT_LIMIT = 1          'Betting limit
Public Const ACT_STARTING_BALANCE = 2
Public Const ACT_BALANCE = 3        'Current Balance
Public Const ACT_STRAIGHT_BET = 4   'Straight Bet Price
Public Const ACT_PARLAY2 = 5        '2 Team Parlay Price
Public Const ACT_PARLAY3 = 6        '3 Team Parlay Price
Public Const ACT_PARLAY4 = 7        '4 Team Parlay
Public Const ACT_PARLAY_FEE = 8     'Baseball parlay calculation fee
Public Const ACT_4PT_TEASER2 = 9    '4pt 2-Team Teaser Odds
Public Const ACT_4PT_TEASER3 = 10   '4pt 3-Team Teaser Odds
Public Const ACT_6PT_TEASER2 = 11   '6pt 2-Team Teaser Odds
Public Const ACT_6PT_TEASER3 = 12   '6pt 3-Team Teaser Odds
Public Const ACT_7PT_TEASER2 = 13   '7pt 2-Team Teaser Odds
Public Const ACT_7PT_TEASER3 = 14   '7pt 3-Team Teaser Odds
Public Const ACT_PITCHERS = 15      'Pitchers Musr Start
Public Const ACT_4PT_TEASER = 16    '4pt Teaser Ties Push
Public Const ACT_6PT_TEASER = 17    '6pt Teaser Ties Push
Public Const ACT_7PT_TEASER = 18    '7pt Teaser Ties Push
Public Const ACT_IN_ACTION = 19     'In Action
Public Const ACT_LAST_WEEK = 20     'Last Week Balance
Public Const MAX_ACT_ITEMS = 21     'Number of Account Settings

Public Const SCH_GAME_NBR = 0       'Schedule Array Indexes
Public Const SCH_DATE = 1
Public Const SCH_ROTATION = 2
Public Const SCH_ROAD_TEAM = 3
Public Const SCH_HOME_TEAM = 4
Public Const SCH_LINE = 5
Public Const SCH_TOTAL = 6
Public Const SCH_ROAD_PRICE = 7
Public Const SCH_HOME_PRICE = 8
Public Const SCH_OVER_PRICE = 9
Public Const SCH_UNDER_PRICE = 10
Public Const SCH_ROAD_SCORE = 11
Public Const SCH_HOME_SCORE = 12
Public Const MAX_SCH_ITEMS = 13

Public Const BET_TYPE_STRAIGHT = 0
Public Const BET_TYPE_PARLAY = 1
Public Const BET_TYPE_4TEASER = 2
Public Const BET_TYPE_6TEASER = 3
Public Const BET_TYPE_7TEASER = 4
Public Const BET_TYPE_ADJUSTMENT = 8
Public Const BET_TYPE_PAYMENT = 9

Public Const OneHalf = "½"

'Global Variables
Public MyDb As DATABASE                'Database Object
Public MyTable As Recordset            'Table Object
Public DefaultSettings(MAX_ACT_ITEMS) As String
Public AccountSettings(MAX_ACT_ITEMS) As String
Public sGames() As String              'Scheduled Games by Sport
Public nGames(SPORTS) As Long       'Number of Scheduled Games by Sport
Public nSchNbr() As Long            'Game Numbers
Public sAdjustments() As String
Public sPayments() As String
Public sCurrentAccount As String

'Formatted Line Strings
Public strGameTotal As String
Public strOverPrice As String
Public strUnderPrice As String
Public strRoadSidePrice As String
Public strHomeSidePrice As String
Public strRoadLine As String
Public strHomeLine As String
Public strBetDisplay As String

'Win32API Types and Function Declarations
'These Were Copied From the File WINAPI\WIN32API.TXT
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_YESNOCANCEL = &H3&
Public Const MB_YESNO = &H4&
Public Const MB_RETRYCANCEL = &H5&
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7

Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Sub ApplyLineChanges()

     Dim i, x, cur, changed As Long
    Dim SQL As String
    Dim lResult As Long
    
    cur = 0
    For i = 0 To SPORTS - 1
        For x = 0 To nGames(i) - 1
                
            changed = 0
            
            FormLines.GridLines.Row = (cur * 2) + 1
            
            FormLines.GridLines.Col = 6
            If sGames(i, (x * MAX_SCH_ITEMS) + SCH_ROAD_PRICE) <> FormLines.GridLines.Text Then changed = 1
           
            FormLines.GridLines.Col = 8
            If sGames(i, (x * MAX_SCH_ITEMS) + SCH_OVER_PRICE) <> FormLines.GridLines.Text Then changed = 1
  
            FormLines.GridLines.Row = (cur * 2) + 2
            
            FormLines.GridLines.Col = 5
            If sGames(i, (x * MAX_SCH_ITEMS) + SCH_LINE) <> FormLines.GridLines.Text Then changed = 1
            
            FormLines.GridLines.Col = 6
            If sGames(i, (x * MAX_SCH_ITEMS) + SCH_HOME_PRICE) <> FormLines.GridLines.Text Then changed = 1
            
            FormLines.GridLines.Col = 7
            If sGames(i, (x * MAX_SCH_ITEMS) + SCH_TOTAL) <> FormLines.GridLines.Text Then changed = 1
            
            FormLines.GridLines.Col = 8
            If sGames(i, (x * MAX_SCH_ITEMS) + SCH_UNDER_PRICE) <> FormLines.GridLines.Text Then changed = 1
            
            If changed > 0 Then
            
                FormLines.GridLines.Row = (cur * 2) + 1
                
                FormLines.GridLines.Col = 6
                If Len(FormLines.GridLines.Text) > 0 Then
                    SQL = "Update Schedule Set RoadPrice = " & Val(FormLines.GridLines.Text) & " Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                Else
                    SQL = "Update Schedule Set RoadPrice = '' Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                End If
                lResult = dbUpdate(SQL)
                If lResult <> 0 Then Exit Sub
           
                FormLines.GridLines.Col = 8
                If Len(FormLines.GridLines.Text) > 0 Then
                    SQL = "Update Schedule Set OverPrice = " & Val(FormLines.GridLines.Text) & " Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                Else
                    SQL = "Update Schedule Set OverPrice = '' Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                End If
                lResult = dbUpdate(SQL)
                If lResult <> 0 Then Exit Sub
                
                FormLines.GridLines.Row = (cur * 2) + 2
           
                FormLines.GridLines.Col = 5
                If Len(FormLines.GridLines.Text) > 0 Then
                    SQL = "Update Schedule Set Line = " & Val(FormLines.GridLines.Text) & " Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                Else
                    SQL = "Update Schedule Set Line = '' Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                End If
                lResult = dbUpdate(SQL)
                If lResult <> 0 Then Exit Sub
           
                FormLines.GridLines.Col = 6
                If Len(FormLines.GridLines.Text) > 0 Then
                    SQL = "Update Schedule Set HomePrice = " & Val(FormLines.GridLines.Text) & " Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                Else
                    SQL = "Update Schedule Set HomePrice = '' Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                End If
                lResult = dbUpdate(SQL)
                If lResult <> 0 Then Exit Sub
            
                FormLines.GridLines.Col = 7
                If Len(FormLines.GridLines.Text) > 0 Then
                    SQL = "Update Schedule Set Total = " & Val(FormLines.GridLines.Text) & " Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                Else
                    SQL = "Update Schedule Set Total = '' Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                End If
                lResult = dbUpdate(SQL)
                If lResult <> 0 Then Exit Sub
       
                FormLines.GridLines.Col = 8
                If Len(FormLines.GridLines.Text) > 0 Then
                    SQL = "Update Schedule Set UnderPrice = " & Val(FormLines.GridLines.Text) & " Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                Else
                    SQL = "Update Schedule Set UnderPrice = '' Where GameNumber = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                End If
                lResult = dbUpdate(SQL)
                If lResult <> 0 Then Exit Sub
            
            End If
            
            cur = cur + 1
        Next x
    Next i
    ReadSchedule
End Sub

Public Sub DisplayLineChanges()

    Dim i As Long, x As Long, cur As Long
    Dim iH As Long, iR As Long, iO As Long, iU As Long
    Dim allSides As Long, allTotals As Long
    Dim lResult As Long

    allSides = 0
    allTotals = 0

    'set number of rows
    x = 0
    For i = 0 To SPORTS - 1
        x = x + nGames(i)
    Next i
    FormLines.GridLines.Rows = (x * 2) + 1

    'set column headings
    FormLines.GridLines.Row = 0
    FormLines.GridLines.Col = 0
    FormLines.GridLines.Text = "Nbr"
    FormLines.GridLines.Col = 1
    FormLines.GridLines.Text = "Date/Time"
    FormLines.GridLines.Col = 2
    FormLines.GridLines.Text = "Teams"
    FormLines.GridLines.Col = 3
    FormLines.GridLines.Text = "$Side$"
    FormLines.GridLines.Col = 4
    FormLines.GridLines.Text = "$OvUn$"
    FormLines.GridLines.Col = 5
    FormLines.GridLines.Text = "Line"
    FormLines.GridLines.Col = 6
    FormLines.GridLines.Text = "Price"
    FormLines.GridLines.Col = 7
    FormLines.GridLines.Text = "Total"
    FormLines.GridLines.Col = 8
    FormLines.GridLines.Text = "Price"
    
    'set column widths
    FormLines.GridLines.ColWidth(0) = 500
    FormLines.GridLines.ColWidth(1) = 1100
    FormLines.GridLines.ColWidth(2) = 2400
    FormLines.GridLines.ColWidth(3) = 700
    FormLines.GridLines.ColWidth(4) = 700
    FormLines.GridLines.ColWidth(5) = 700
    FormLines.GridLines.ColWidth(6) = 700
    FormLines.GridLines.ColWidth(7) = 700
    FormLines.GridLines.ColWidth(8) = 700
    
    'populate the cells
    cur = 0
    For i = 0 To SPORTS - 1
        For x = 0 To nGames(i) - 1
            FormLines.GridLines.Row = (cur * 2) + 1
        
            iO = 0
            iU = 0
            iH = 0
            iR = 0
            lResult = dbSelect("Select * from Ledger where deleted = 0 and type = " & BET_TYPE_STRAIGHT & " and game1 = " & sGames(i, (x * MAX_SCH_ITEMS) + SCH_GAME_NBR))
            If lResult = 0 Then
                Do Until MyTable.EOF
                    If MyTable![Side1] = "H" Then iH = iH + MyTable![amount]
                    If MyTable![Side1] = "R" Then iR = iR + MyTable![amount]
                    If MyTable![Side1] = "O" Then iO = iO + MyTable![amount]
                    If MyTable![Side1] = "U" Then iU = iU + MyTable![amount]
                    MyTable.MoveNext
                Loop
                MyTable.Close
            End If
        
            FormLines.GridLines.Col = 0
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_ROTATION)
            FormLines.GridLines.Col = 1
            FormLines.GridLines.Text = Format(sGames(i, (x * MAX_SCH_ITEMS) + SCH_DATE), "ddd m/d")
            FormLines.GridLines.Col = 2
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_ROAD_TEAM)
            
            FormLines.GridLines.Col = 3
            FormLines.GridLines.Text = iR
            FormLines.GridLines.Col = 4
            FormLines.GridLines.Text = iO
            
            FormLines.GridLines.Col = 6
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)
            
            FormLines.GridLines.Col = 8
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_OVER_PRICE)
            
            FormLines.GridLines.Row = (cur * 2) + 2
            FormLines.GridLines.Col = 0
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_ROTATION) + 1
            FormLines.GridLines.Col = 1
            FormLines.GridLines.Text = Format(sGames(i, (x * MAX_SCH_ITEMS) + SCH_DATE), "h:mm AMPM")
            FormLines.GridLines.Col = 2
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_HOME_TEAM)
            
            FormLines.GridLines.Col = 3
            FormLines.GridLines.Text = iH
            FormLines.GridLines.Col = 4
            FormLines.GridLines.Text = iU
            
            allSides = allSides + iH + iR
            allTotals = allTotals + iO + iU
            
            FormLines.GridLines.Col = 5
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_LINE)
            
            FormLines.GridLines.Col = 6
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_HOME_PRICE)
            
            FormLines.GridLines.Col = 7
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_TOTAL)
            
            FormLines.GridLines.Col = 8
            FormLines.GridLines.Text = sGames(i, (x * MAX_SCH_ITEMS) + SCH_UNDER_PRICE)
           
            cur = cur + 1
        Next x
    Next i
    
    FormLines.lblSides.Caption = allSides
    FormLines.lblTotals.Caption = allTotals
    
    'set edit tag to off
    FormLines.GridLines.Tag = "0"
    
End Sub


Public Function NoQuotes(s As String) As String
    Dim ss As String
    Dim i As Integer
    
    NoQuotes = s
    If InStr(1, s, "'") Then
        ss = ""
        For i = 1 To Len(s)
            If Mid(s, i, 1) <> "'" Then ss = ss & Mid(s, i, 1)
        Next i
        NoQuotes = ss
    End If

End Function

Public Sub ReadDefaults()
   
    Dim lResult As Long
    
    'Read System Defaults
    lResult = dbSelect("Select * From Defaults")
    If lResult <> 0 Then Exit Sub
    
    If Not MyTable.EOF Then
        DefaultSettings(ACT_LIMIT) = MyTable![LIMIT]
        DefaultSettings(ACT_STRAIGHT_BET) = MyTable![StraightBet]
        DefaultSettings(ACT_PARLAY2) = MyTable![PARLAY2]
        DefaultSettings(ACT_PARLAY3) = MyTable![PARLAY3]
        DefaultSettings(ACT_PARLAY4) = MyTable![PARLAY4]
        DefaultSettings(ACT_PARLAY_FEE) = MyTable![PARLAYFEE]
        DefaultSettings(ACT_4PT_TEASER2) = MyTable![Teaser4Two]
        DefaultSettings(ACT_4PT_TEASER3) = MyTable![Teaser4Three]
        DefaultSettings(ACT_6PT_TEASER2) = MyTable![Teaser6Two]
        DefaultSettings(ACT_6PT_TEASER3) = MyTable![Teaser6Three]
        DefaultSettings(ACT_7PT_TEASER2) = MyTable![Teaser7Two]
        DefaultSettings(ACT_7PT_TEASER3) = MyTable![Teaser7Three]
        DefaultSettings(ACT_PITCHERS) = MyTable![Pitchers]
        DefaultSettings(ACT_4PT_TEASER) = MyTable![Teaser4Ties]
    End If
     
    MyTable.Close

End Sub


Public Sub ReadSchedule()
        
    Dim lResult As Long
    Dim i As Long       'Local Temporary Counter
    Dim x As Long       'Local Temporary Counter
    Dim idx As Long     'Sport Index
    
    For i = 0 To SPORTS - 1
        nGames(i) = 0
    Next i
    
    'Open the Schedule Table
    lResult = dbSelect("Select * From Schedule Order By RotationNumber")
    If lResult <> 0 Then Exit Sub
    If MyTable.EOF Then Exit Sub
    
    Do Until MyTable.EOF
        'Only Look at Games of Today or Later
        If DateDiff("d", Now, MyTable![GameDate]) >= 0 Then
            If MyTable![Sport] = "NFL" Then
                nGames(NFL_TAB) = nGames(NFL_TAB) + 1
            End If
            If MyTable![Sport] = "NBA" Then
                If DateDiff("d", Now, MyTable![GameDate]) = 0 Or NBA_TODAY_ONLY = False Then
                    nGames(NBA_TAB) = nGames(NBA_TAB) + 1
                End If
            End If
            If MyTable![Sport] = "MLB" Then
                If DateDiff("d", Now, MyTable![GameDate]) = 0 Or MLB_TODAY_ONLY = False Then
                    nGames(MLB_TAB) = nGames(MLB_TAB) + 1
                End If
            End If
            If MyTable![Sport] = "NHL" Then
                If DateDiff("d", Now, MyTable![GameDate]) = 0 Or NHL_TODAY_ONLY = False Then
                    nGames(NHL_TAB) = nGames(NHL_TAB) + 1
                End If
            End If
       End If
       MyTable.MoveNext
    Loop
 
  
    'ReDim the Schedule Array to Largest Sport
    i = nGames(NFL_TAB)
    If nGames(NBA_TAB) > i Then i = nGames(NBA_TAB)
    If nGames(MLB_TAB) > i Then i = nGames(MLB_TAB)
    If nGames(NHL_TAB) > i Then i = nGames(NHL_TAB)
    ReDim sGames(SPORTS, (i * MAX_SCH_ITEMS))
    
    'Move Back to the Top of the Scheduled Games
    MyTable.MoveFirst
    
    For i = 0 To SPORTS - 1
        nGames(i) = 0
    Next i

    'Read The Schedule Again
    i = 0
    Do Until MyTable.EOF
        'Only Look at Games of Today or Later
        If DateDiff("d", Now, MyTable![GameDate]) >= 0 Then
        
            If MyTable![Sport] = "MLB" And DateDiff("d", Now, MyTable![GameDate]) > 0 And MLB_TODAY_ONLY = True Then
                GoTo SKIPIT
            End If
            
            If MyTable![Sport] = "NBA" And DateDiff("d", Now, MyTable![GameDate]) > 0 And NBA_TODAY_ONLY = True Then
                GoTo SKIPIT
            End If
            
            If MyTable![Sport] = "NHL" And DateDiff("d", Now, MyTable![GameDate]) > 0 And NHL_TODAY_ONLY = True Then
                GoTo SKIPIT
            End If
        
        
            'Point to the Right Sport
            If MyTable![Sport] = "NFL" Then idx = NFL_TAB
            If MyTable![Sport] = "NBA" Then idx = NBA_TAB
            If MyTable![Sport] = "MLB" Then idx = MLB_TAB
            If MyTable![Sport] = "NHL" Then idx = NHL_TAB
        
            'Point to the Right Game in that Sport
            x = nGames(idx) * MAX_SCH_ITEMS
            
            'Fill in the Values for That Game
            sGames(idx, x + SCH_GAME_NBR) = MyTable![GameNumber]
            sGames(idx, x + SCH_DATE) = MyTable![GameDate]
            sGames(idx, x + SCH_ROTATION) = MyTable![RotationNumber]
            If Not IsNull(MyTable![RoadTeam]) Then sGames(idx, x + SCH_ROAD_TEAM) = MyTable![RoadTeam]
            If Not IsNull(MyTable![HomeTeam]) Then sGames(idx, x + SCH_HOME_TEAM) = MyTable![HomeTeam]
            
            'Only Fill in the Lines if the Game is Not Underway
            If DateDiff("s", Now, MyTable![GameDate]) > 0 Or DEPOSIT_SYSTEM = True Then
                If Not IsNull(MyTable![line]) Then sGames(idx, x + SCH_LINE) = MyTable![line]
                If Not IsNull(MyTable![total]) Then sGames(idx, x + SCH_TOTAL) = MyTable![total]
                If Not IsNull(MyTable![RoadPrice]) Then sGames(idx, x + SCH_ROAD_PRICE) = MyTable![RoadPrice]
                If Not IsNull(MyTable![HomePrice]) Then sGames(idx, x + SCH_HOME_PRICE) = MyTable![HomePrice]
                If Not IsNull(MyTable![OverPrice]) Then sGames(idx, x + SCH_OVER_PRICE) = MyTable![OverPrice]
                If Not IsNull(MyTable![UnderPrice]) Then sGames(idx, x + SCH_UNDER_PRICE) = MyTable![UnderPrice]
           
                If idx = NHL_TAB Then
                    If Len(sGames(idx, x + SCH_LINE)) > 0 Then
                        If Len(sGames(idx, x + SCH_ROAD_PRICE)) = 0 Then sGames(idx, x + SCH_ROAD_PRICE) = 100
                        If Len(sGames(idx, x + SCH_HOME_PRICE)) = 0 Then sGames(idx, x + SCH_HOME_PRICE) = 100
                    End If
                    If Len(sGames(idx, x + SCH_TOTAL)) > 0 Then
                        If Len(sGames(idx, x + SCH_OVER_PRICE)) = 0 Then sGames(idx, x + SCH_OVER_PRICE) = -110
                        If Len(sGames(idx, x + SCH_UNDER_PRICE)) = 0 Then sGames(idx, x + SCH_UNDER_PRICE) = -110
                    End If
                End If
            End If
            
            'Point to the Next Game for That Sport
            nGames(idx) = nGames(idx) + 1
        
        End If
SKIPIT:
        MyTable.MoveNext
    Loop

    'Close the Schedule Table
    MyTable.Close
    
End Sub






Public Sub DisplayScores()
        
    Dim lResult As Long
    Dim i As Long       'Local Temporary Counter
    Dim thedate As Date

    thedate = Now
    thedate = DateAdd("d", FormDate.sbDate.Value - 500, thedate)

    'Open the Schedule Table
    lResult = dbSelect("Select * From Schedule Order By RotationNumber")
    If lResult <> 0 Then Exit Sub
    
    i = 0
    Do Until MyTable.EOF
        'Only Look at Games of Selected Date
        If DateDiff("d", thedate, MyTable![GameDate]) = 0 Then
            i = i + 1
        End If
       MyTable.MoveNext
    Loop
    ReDim nSchNbr(i)
    
    FormScores.GridScore.Rows = (i * 2) + 1

    'set column headings
    FormScores.GridScore.Row = 0
    FormScores.GridScore.Col = 0
    FormScores.GridScore.Text = "Nbr"
    FormScores.GridScore.Col = 1
    FormScores.GridScore.Text = "Date/Time"
    FormScores.GridScore.Col = 2
    FormScores.GridScore.Text = "Teams"
    FormScores.GridScore.Col = 3
    FormScores.GridScore.Text = "Score"
    
    'set column widths
    FormScores.GridScore.ColWidth(0) = 600
    FormScores.GridScore.ColWidth(1) = 1100
    FormScores.GridScore.ColWidth(2) = 2830
    FormScores.GridScore.ColWidth(3) = 800
    
    MyTable.MoveFirst
    i = 0
    Do Until MyTable.EOF
        If DateDiff("d", thedate, MyTable![GameDate]) = 0 Then
            nSchNbr(i) = MyTable![GameNumber]
         
            FormScores.GridScore.Row = (i * 2) + 1
            FormScores.GridScore.Col = 0
            FormScores.GridScore.Text = MyTable![RotationNumber]
            FormScores.GridScore.Col = 1
            FormScores.GridScore.Text = Format(MyTable![GameDate], "ddd m/d")
            FormScores.GridScore.Col = 2
            FormScores.GridScore.Text = MyTable![RoadTeam]
            FormScores.GridScore.Col = 3
            If Not IsNull(MyTable![RoadScore]) Then FormScores.GridScore.Text = MyTable![RoadScore]
     
            FormScores.GridScore.Row = (i * 2) + 2
            FormScores.GridScore.Col = 0
            FormScores.GridScore.Text = MyTable![RotationNumber] + 1
            FormScores.GridScore.Col = 1
            FormScores.GridScore.Text = Format(MyTable![GameDate], "h:mm AMPM")
            FormScores.GridScore.Col = 2
            FormScores.GridScore.Text = MyTable![HomeTeam]
            FormScores.GridScore.Col = 3
            If Not IsNull(MyTable![HomeScore]) Then FormScores.GridScore.Text = MyTable![HomeScore]
         
            If Not IsNull(MyTable![Cancelled]) Then
                If MyTable![Cancelled] = "C" Then
                    FormScores.GridScore.Row = (i * 2) + 1
                    FormScores.GridScore.Col = 3
                    FormScores.GridScore.Text = "NoGame"
                    
                    FormScores.GridScore.Row = (i * 2) + 2
                    FormScores.GridScore.Col = 3
                    FormScores.GridScore.Text = "NoGame"
                End If
            End If
            
            i = i + 1
        End If
        MyTable.MoveNext
    Loop
 
    'set edit tag to off
    FormScores.GridScore.Tag = "0"
 
    'Close the Schedule Table
    MyTable.Close

End Sub

Public Sub ReCalcBalances()

    Dim i As Long, x As Long, nCount As Long, bal As Long
    Dim inA As Long, daysThisWeek As Long, lastWeek As Long, thisWeek As Long
    Dim accounts() As String
    Dim SchRecords() As String
    Dim theAccount As New Account
    Dim SQL As String, DayOfWeek As String
    Dim lResult As Long
    Dim CutOffDate As Date, thisWeekDate As Date, thedate As Date

    'Calculate The CutOffDate
    DayOfWeek = Format(Now(), "ddd")
    If DayOfWeek = "Tue" Then CutOffDate = DateAdd("d", -7, Now())
    If DayOfWeek = "Wed" Then CutOffDate = DateAdd("d", -8, Now())
    If DayOfWeek = "Thu" Then CutOffDate = DateAdd("d", -9, Now())
    If DayOfWeek = "Fri" Then CutOffDate = DateAdd("d", -10, Now())
    If DayOfWeek = "Sat" Then CutOffDate = DateAdd("d", -11, Now())
    If DayOfWeek = "Sun" Then CutOffDate = DateAdd("d", -12, Now())
    If DayOfWeek = "Mon" Then CutOffDate = DateAdd("d", -13, Now())

    'Calculate The CutOffDate
    If DayOfWeek = "Tue" Then daysThisWeek = 7
    If DayOfWeek = "Wed" Then daysThisWeek = 1
    If DayOfWeek = "Thu" Then daysThisWeek = 2
    If DayOfWeek = "Fri" Then daysThisWeek = 3
    If DayOfWeek = "Sat" Then daysThisWeek = 4
    If DayOfWeek = "Sun" Then daysThisWeek = 5
    If DayOfWeek = "Mon" Then daysThisWeek = 6
    thisWeekDate = DateAdd("d", 0 - daysThisWeek, Now())
    
    'Get Number of Accounts
    lResult = dbSelect("Select * from Accounts")
    If lResult <> 0 Then Exit Sub
    nCount = 0
    Do Until MyTable.EOF
        nCount = nCount + 1
        MyTable.MoveNext
    Loop
    ReDim accounts(nCount)
    
    'Save Account Numbers
    If nCount > 0 Then
        MyTable.MoveFirst
        i = 0
        Do Until MyTable.EOF
            accounts(i) = MyTable![Account]
            i = i + 1
            MyTable.MoveNext
        Loop
    End If
    
    'Close the Accounts Table
    MyTable.Close

    For i = 0 To nCount - 1
        sCurrentAccount = accounts(i)
        ReadAccount
        theAccount.Populate (sCurrentAccount)
        
        bal = theAccount.CurrentBalance
        SQL = "Update Accounts Set Balance = " & bal & " Where Account = '" & sCurrentAccount & "'"
        Call dbUpdate(SQL)
    
        inA = theAccount.InAction
        SQL = "Update Accounts Set InAction = " & inA & " Where Account = '" & sCurrentAccount & "'"
        Call dbUpdate(SQL)
    
        thisWeek = 0
        For x = 0 To daysThisWeek
            thedate = DateAdd("d", x, thisWeekDate)
            thisWeek = thisWeek + Val(theAccount.action(thedate))
        Next x
        lastWeek = bal - thisWeek
    
        SQL = "Update Accounts Set LastWeekBalance = " & lastWeek & " Where Account = '" & sCurrentAccount & "'"
        Call dbUpdate(SQL)
    
        'All Ledger Records Before The CutOffDate Are Deleted
        'The Balance Of These Records Becomes The Starting Balance
         theAccount.PurgeOldData (CutOffDate)
    Next i
    
    'Delete UnReferenced Schedule Records Before The CutOffDate
    'Build a List of Old Games
    i = 0
    ReDim SchRecords(0)
    lResult = dbSelect("Select * From Schedule")
    If lResult <> 0 Then Exit Sub
    Do Until MyTable.EOF
        'Only Look at Games Before the CutOffDate
        If DateDiff("d", CutOffDate, MyTable![GameDate]) < 0 Then
            i = i + 1
            ReDim Preserve SchRecords(i)
            SchRecords(i - 1) = MyTable![GameNumber]
        End If
        MyTable.MoveNext
    Loop
    MyTable.Close
    
    'Read Thru Remaining Ledgers to Make Sure These Games Are Not Referenced
    lResult = dbSelect("Select * From Ledger")
    If lResult <> 0 Then Exit Sub
    Do Until MyTable.EOF
        For x = 0 To i - 1
            If Not IsNull(MyTable![Game1]) Then
                If Val(SchRecords(x)) = Val(MyTable![Game1]) Then SchRecords(x) = "0"
            End If
            If Not IsNull(MyTable![Game2]) Then
                If Val(SchRecords(x)) = Val(MyTable![Game2]) Then SchRecords(x) = "0"
            End If
            If Not IsNull(MyTable![Game3]) Then
                If Val(SchRecords(x)) = Val(MyTable![Game3]) Then SchRecords(x) = "0"
            End If
            If Not IsNull(MyTable![Game4]) Then
                If Val(SchRecords(x)) = Val(MyTable![Game4]) Then SchRecords(x) = "0"
            End If
        Next x
        MyTable.MoveNext
    Loop
    MyTable.Close
   
    'Delete the Old Games
     For x = 0 To i - 1
        If Val(SchRecords(x)) > 0 Then
            Call dbDelete("Delete * From Schedule Where GameNumber = " & SchRecords(x))
        End If
     Next x

End Sub

Public Sub ReadAccount()
    
    Dim i As Long
    Dim lResult As Long
    
    'Clear the Current Account
    For i = 0 To MAX_ACT_ITEMS - 1
        AccountSettings(i) = ""
    Next i
    
    'Select the Account Record From the Accounts Table
    lResult = dbSelect("Select * From Accounts Where Account = '" & sCurrentAccount & "'")
    If lResult <> 0 Then
         'Do Nothing on an Error
         Exit Sub
    End If
    
    'Check if an Account was Selected
    If MyTable.EOF Then
        'The Selected Account Was Not Found
        Exit Sub
    End If

    AccountSettings(ACT_ACCOUNT) = MyTable![Account]
    AccountSettings(ACT_BALANCE) = MyTable![balance]
    AccountSettings(ACT_STARTING_BALANCE) = MyTable![StartingBalance]
    AccountSettings(ACT_PITCHERS) = MyTable![Pitchers]
    AccountSettings(ACT_4PT_TEASER) = MyTable![Teaser4Ties]
    AccountSettings(ACT_6PT_TEASER) = MyTable![Teaser6Ties]
    AccountSettings(ACT_7PT_TEASER) = MyTable![Teaser7Ties]
    AccountSettings(ACT_IN_ACTION) = MyTable![InAction]
    AccountSettings(ACT_LAST_WEEK) = MyTable![LastWeekBalance]
    
    If IsNull(MyTable![LIMIT]) Then
        AccountSettings(ACT_LIMIT) = DefaultSettings(ACT_LIMIT)
    Else
        AccountSettings(ACT_LIMIT) = MyTable![LIMIT]
    End If
    
    If IsNull(MyTable![StraightBet]) Then
        AccountSettings(ACT_STRAIGHT_BET) = DefaultSettings(ACT_STRAIGHT_BET)
    Else
        AccountSettings(ACT_STRAIGHT_BET) = MyTable![StraightBet]
    End If
    If AccountSettings(ACT_STRAIGHT_BET) = 0 Then
        AccountSettings(ACT_STRAIGHT_BET) = DefaultSettings(ACT_STRAIGHT_BET)
    End If
        
    If IsNull(MyTable![Teaser6Two]) Then
        AccountSettings(ACT_6PT_TEASER2) = DefaultSettings(ACT_6PT_TEASER2)
    Else
        AccountSettings(ACT_6PT_TEASER2) = MyTable![Teaser6Two]
    End If
    If AccountSettings(ACT_6PT_TEASER2) = 0 Then
        AccountSettings(ACT_6PT_TEASER2) = DefaultSettings(ACT_6PT_TEASER2)
    End If
        
    If IsNull(MyTable![Teaser7Two]) Then
        AccountSettings(ACT_7PT_TEASER2) = DefaultSettings(ACT_7PT_TEASER2)
    Else
        AccountSettings(ACT_7PT_TEASER2) = MyTable![Teaser7Two]
    End If
    If AccountSettings(ACT_7PT_TEASER2) = 0 Then
        AccountSettings(ACT_7PT_TEASER2) = DefaultSettings(ACT_7PT_TEASER2)
    End If
        
    AccountSettings(ACT_PARLAY2) = DefaultSettings(ACT_PARLAY2)
    AccountSettings(ACT_PARLAY3) = DefaultSettings(ACT_PARLAY3)
    AccountSettings(ACT_PARLAY4) = DefaultSettings(ACT_PARLAY4)
    AccountSettings(ACT_PARLAY_FEE) = DefaultSettings(ACT_PARLAY_FEE)
    AccountSettings(ACT_4PT_TEASER2) = DefaultSettings(ACT_4PT_TEASER2)
    AccountSettings(ACT_4PT_TEASER3) = DefaultSettings(ACT_4PT_TEASER3)
    AccountSettings(ACT_6PT_TEASER3) = DefaultSettings(ACT_6PT_TEASER3)
    AccountSettings(ACT_7PT_TEASER3) = DefaultSettings(ACT_7PT_TEASER3)
        
    MyTable.Close

End Sub

Public Sub DisplayAccountDetail()

    FormSelectAccount.MousePointer = vbHourglass
       
    Call FormDetail.theAccount.Populate(sCurrentAccount)
    Call FormDetail.theAccount.Display
 
    FormDetail.lBackground(0).Caption = "Account: " & sCurrentAccount
    FormDetail.tLimit = AccountSettings(ACT_LIMIT)
    FormDetail.tStraightBet = AccountSettings(ACT_STRAIGHT_BET)
    
    If Val(AccountSettings(ACT_6PT_TEASER2)) >= 0 Then
        FormDetail.tTeaser6Two = "+" & AccountSettings(ACT_6PT_TEASER2)
    Else
        FormDetail.tTeaser6Two = AccountSettings(ACT_6PT_TEASER2)
    End If
 
    If Val(AccountSettings(ACT_7PT_TEASER2)) >= 0 Then
        FormDetail.tTeaser7Two = "+" & AccountSettings(ACT_7PT_TEASER2)
    Else
        FormDetail.tTeaser7Two = AccountSettings(ACT_7PT_TEASER2)
    End If
 
    If Val(AccountSettings(ACT_PITCHERS)) = 1 Then
        FormDetail.rbPitchers(0).Value = True
    Else
        FormDetail.rbPitchers(1).Value = True
    End If
    
    If Val(AccountSettings(ACT_4PT_TEASER)) = 1 Then
        FormDetail.rb4ptteaser(0).Value = True
    Else
        FormDetail.rb4ptteaser(1).Value = True
    End If
 
    If Val(AccountSettings(ACT_6PT_TEASER)) = 1 Then
        FormDetail.rb6pt(0).Value = True
    Else
        FormDetail.rb6pt(1).Value = True
    End If
 
    If Val(AccountSettings(ACT_7PT_TEASER)) = 1 Then
        FormDetail.rb7pt(0).Value = True
    Else
        FormDetail.rb7pt(1).Value = True
    End If
    
    FormSelectAccount.MousePointer = vbDefault

End Sub

Public Sub DisplayAdjustments()

    Dim lResult As Long
    Dim i As Long       'Local Temporary Counter
    Dim thedate As Date
    Dim accounts() As String

    thedate = Now
    
    'Open the Accounts Table
    lResult = dbSelect("Select Account From Accounts Where Status = 'A' Order By Account")
    If lResult <> 0 Then Exit Sub
    
    i = 0
    Do Until MyTable.EOF
        i = i + 1
        MyTable.MoveNext
    Loop
    ReDim accounts(i)
    ReDim sAdjustments(i)
    FormAdjustments.GridAdjustments.Rows = i + 1

    'set column headings
    FormAdjustments.GridAdjustments.Row = 0
    FormAdjustments.GridAdjustments.Col = 0
    FormAdjustments.GridAdjustments.Text = "Account"
    FormAdjustments.GridAdjustments.Col = 1
    FormAdjustments.GridAdjustments.Text = "Adjustment"
    
    'set column widths
    FormAdjustments.GridAdjustments.ColWidth(0) = 1600
    FormAdjustments.GridAdjustments.ColWidth(1) = 1000
    
    MyTable.MoveFirst
    i = 0
    Do Until MyTable.EOF
        sAdjustments(i) = ""
        FormAdjustments.GridAdjustments.Row = i + 1
        FormAdjustments.GridAdjustments.Col = 0
        FormAdjustments.GridAdjustments.Text = MyTable![Account]
        i = i + 1
        MyTable.MoveNext
    Loop
 
    'set edit tag to off
    FormAdjustments.GridAdjustments.Tag = "0"
 
    'Close the Accounts Table
    MyTable.Close



    'Open the Ledger Table
    lResult = dbSelect("Select * From Ledger Where Type = " & BET_TYPE_ADJUSTMENT)
    If lResult <> 0 Then Exit Sub
    Do Until MyTable.EOF
        For i = 1 To FormAdjustments.GridAdjustments.Rows - 1
            FormAdjustments.GridAdjustments.Row = i
            FormAdjustments.GridAdjustments.Col = 0
            If FormAdjustments.GridAdjustments.Text = MyTable![Account] Then
                If DateDiff("d", thedate, MyTable![TimeStamp]) = -1 Then
                    FormAdjustments.GridAdjustments.Col = 1
                    FormAdjustments.GridAdjustments.Text = MyTable![amount]
                    sAdjustments(i) = MyTable![amount]
                End If
            End If
        Next i
        MyTable.MoveNext
    Loop
    MyTable.Close

End Sub

Public Sub ApplyAdjustments()
    Dim lResult As Long
    Dim i, bal As Long       'Local Temporary Counter
    Dim thedate As Date
    Dim theAccount As String, theAdj As String, theTrans As String
    Dim myAccount As New Account
    Dim SQL As String

    FormAdjustments.MousePointer = vbHourglass
    thedate = Now
    
    For i = 1 To FormAdjustments.GridAdjustments.Rows - 1
        FormAdjustments.GridAdjustments.Row = i
        FormAdjustments.GridAdjustments.Col = 0
        theAccount = FormAdjustments.GridAdjustments.Text
        FormAdjustments.GridAdjustments.Col = 1
        theAdj = FormAdjustments.GridAdjustments.Text
        
        If Val(theAdj) <> Val(sAdjustments(i)) Then
        
            'get the ledger number
            theTrans = ""
            lResult = dbSelect("Select * From Ledger Where Type = " & BET_TYPE_ADJUSTMENT & " And Account = '" & theAccount & "'")
            If lResult <> 0 Then Exit Sub
            Do Until MyTable.EOF
                If DateDiff("d", thedate, MyTable![TimeStamp]) = -1 Then
                    theTrans = MyTable![Transaction]
                    Exit Do
                End If
                MyTable.MoveNext
            Loop
            MyTable.Close
        
            If Val(theAdj) = 0 And Val(theTrans) > 0 Then
                'delete
                Call dbDelete("Delete From Ledger Where Transaction = " & theTrans)
            End If
            
            If Val(sAdjustments(i)) = 0 Then
                'insert
                Date = DateAdd("d", -1, Now())
                Call dbInsert("Insert Into Ledger(Account,Type,Amount) Values (" & "'" & theAccount & "'," & BET_TYPE_ADJUSTMENT & "," & theAdj & ")")
                Date = DateAdd("d", 1, Now())
            End If
                                
            If Val(theAdj) <> 0 And Val(sAdjustments(i)) <> 0 And Val(theTrans) > 0 Then
                'update
                Call dbUpdate("Update Ledger Set Amount = " & theAdj & " Where Transaction = " & theTrans)
            End If
        
            'ReCalc the account
            sCurrentAccount = theAccount
            ReadAccount
            myAccount.Populate (sCurrentAccount)
        
            bal = myAccount.CurrentBalance
            SQL = "Update Accounts Set Balance = " & bal & " Where Account = '" & sCurrentAccount & "'"
            Call dbUpdate(SQL)
        
        End If
    Next i
    FormAdjustments.MousePointer = vbDefault

End Sub

Public Sub DisplayPayments()

    Dim lResult As Long
    Dim i As Long       'Local Temporary Counter
    Dim thedate As Date
    Dim accounts() As String

    thedate = Now
    
    'Open the Accounts Table
    lResult = dbSelect("Select Account From Accounts Where Status = 'A' Order By Account")
    If lResult <> 0 Then Exit Sub
    
    i = 0
    Do Until MyTable.EOF
        i = i + 1
        MyTable.MoveNext
    Loop
    ReDim accounts(i)
    ReDim sPayments(i)
    FormPayments.GridPayments.Rows = i + 1

    'set column headings
    FormPayments.GridPayments.Row = 0
    FormPayments.GridPayments.Col = 0
    FormPayments.GridPayments.Text = "Account"
    FormPayments.GridPayments.Col = 1
    FormPayments.GridPayments.Text = "Amount"
    
    'set column widths
    FormPayments.GridPayments.ColWidth(0) = 1600
    FormPayments.GridPayments.ColWidth(1) = 1000
    
    MyTable.MoveFirst
    i = 0
    Do Until MyTable.EOF
        sPayments(i) = ""
        FormPayments.GridPayments.Row = i + 1
        FormPayments.GridPayments.Col = 0
        FormPayments.GridPayments.Text = MyTable![Account]
        i = i + 1
        MyTable.MoveNext
    Loop
 
    'set edit tag to off
    FormPayments.GridPayments.Tag = "0"
 
    'Close the Accounts Table
    MyTable.Close

    'Open the Ledger Table
    lResult = dbSelect("Select * From Ledger Where Type = " & BET_TYPE_PAYMENT)
    If lResult <> 0 Then Exit Sub
    Do Until MyTable.EOF
        For i = 1 To FormPayments.GridPayments.Rows - 1
            FormPayments.GridPayments.Row = i
            FormPayments.GridPayments.Col = 0
            If FormPayments.GridPayments.Text = MyTable![Account] Then
                If DateDiff("d", thedate, MyTable![TimeStamp]) = 0 Then
                    FormPayments.GridPayments.Col = 1
                    FormPayments.GridPayments.Text = MyTable![amount]
                    sPayments(i) = MyTable![amount]
                End If
            End If
        Next i
        MyTable.MoveNext
    Loop
    MyTable.Close

End Sub

Public Sub ApplyPayments()
   
    Dim lResult As Long
    Dim i, bal As Long       'Local Temporary Counter
    Dim thedate As Date
    Dim theAccount, theAdj, theTrans As String
    Dim myAccount As New Account
    Dim SQL As String

    FormPayments.MousePointer = vbHourglass
    thedate = Now
    
    For i = 1 To FormPayments.GridPayments.Rows - 1
        FormPayments.GridPayments.Row = i
        FormPayments.GridPayments.Col = 0
        theAccount = FormPayments.GridPayments.Text
        FormPayments.GridPayments.Col = 1
        theAdj = FormPayments.GridPayments.Text
        
        If Val(theAdj) <> Val(sPayments(i)) Then
        
            'get the ledger number
            theTrans = ""
            lResult = dbSelect("Select * From Ledger Where Type = " & BET_TYPE_PAYMENT & " And Account = '" & theAccount & "'")
            If lResult <> 0 Then Exit Sub
            Do Until MyTable.EOF
                If DateDiff("d", thedate, MyTable![TimeStamp]) = 0 Then
                    theTrans = MyTable![Transaction]
                    Exit Do
                End If
                MyTable.MoveNext
            Loop
            MyTable.Close
        
            If Val(theAdj) = 0 And Val(theTrans) > 0 Then
                'delete
                Call dbDelete("Delete From Ledger Where Transaction = " & theTrans)
            End If
            
            If Val(sPayments(i)) = 0 Then
                'insert
                Call dbInsert("Insert Into Ledger(Account,Type,Amount) Values (" & "'" & theAccount & "'," & BET_TYPE_PAYMENT & "," & theAdj & ")")
            End If
                                
            If Val(theAdj) <> 0 And Val(sPayments(i)) <> 0 And Val(theTrans) > 0 Then
                'update
                Call dbUpdate("Update Ledger Set Amount = " & theAdj & " Where Transaction = " & theTrans)
            End If
        
            'ReCalc the account
            sCurrentAccount = theAccount
            ReadAccount
            myAccount.Populate (sCurrentAccount)
        
            bal = myAccount.CurrentBalance
            SQL = "Update Accounts Set Balance = " & bal & " Where Account = '" & sCurrentAccount & "'"
            Call dbUpdate(SQL)
        
        End If
    Next i
    FormPayments.MousePointer = vbDefault
End Sub

Public Sub DisplayTotals()

    Dim lResult As Long
    Dim i, nCount As Long       'Local Temporary Counter
    Dim AccountNumbers() As String
    Dim AccountTotals() As Long
    Dim theAccount As New Account

    Form1.MousePointer = vbHourglass

    'Open the Accounts Table
    lResult = dbSelect("Select * From Accounts Order By Account")
    If lResult <> 0 Then Exit Sub
    
    ReDim AccountNumbers(0)
    ReDim AccountTotals(0)
    nCount = 0
    Do Until MyTable.EOF
        If MyTable![StartingBalance] <> 0 Or MyTable![balance] <> 0 Then
            nCount = nCount + 1
            ReDim Preserve AccountNumbers(nCount)
            ReDim Preserve AccountTotals(nCount)
            AccountNumbers(nCount - 1) = MyTable![Account]
            AccountTotals(nCount - 1) = MyTable![balance]
        End If
        MyTable.MoveNext
    Loop
    MyTable.Close
    
    FormTotals.GridTotals.Rows = nCount + 1

    'set column headings
    FormTotals.GridTotals.Row = 0
    FormTotals.GridTotals.Col = 0
    FormTotals.GridTotals.Text = "Account"
    FormTotals.GridTotals.Col = 1
    FormTotals.GridTotals.Text = "Total"
    
    'set column widths
    FormTotals.GridTotals.ColWidth(0) = 1400
    FormTotals.GridTotals.ColWidth(1) = 700
    
    For i = 0 To nCount - 1
        theAccount.Populate (AccountNumbers(i))
        AccountTotals(i) = AccountTotals(i) - theAccount.InAction()
        
        FormTotals.GridTotals.Row = i + 1
        FormTotals.GridTotals.Col = 0
        FormTotals.GridTotals.Text = AccountNumbers(i)
        FormTotals.GridTotals.Col = 1
        FormTotals.GridTotals.Text = AccountTotals(i)
    Next i

    Form1.MousePointer = vbDefault

End Sub

Public Function ExtractTeamName(Team As String) As String
    ExtractTeamName = Team
    
    If InStr(1, Team, " - ") > 0 Then
        ExtractTeamName = Mid(Team, 1, InStr(1, Team, " - ") - 1)
    End If
End Function


Public Sub DisplayPaymentReport()

    Dim lResult As Long
    Dim nDays, i, x, nCount, WeekBal, nRows, Activity As Long
    Dim PayToday, SubTotal, Columns As Long
    Dim AccountNumbers() As String
    Dim CutOffDate, thedate As Date
    Dim DayOfWeek As String

    Form1.MousePointer = vbHourglass

    'Calculate The CutOffDate
    DayOfWeek = Format(Now(), "ddd")
    If DayOfWeek = "Tue" Then Columns = 7
    If DayOfWeek = "Wed" Then Columns = 1
    If DayOfWeek = "Thu" Then Columns = 2
    If DayOfWeek = "Fri" Then Columns = 3
    If DayOfWeek = "Sat" Then Columns = 4
    If DayOfWeek = "Sun" Then Columns = 5
    If DayOfWeek = "Mon" Then Columns = 6
    CutOffDate = DateAdd("d", 0 - Columns, Now())
    PayToday = 0

    'Open the Accounts Table
    lResult = dbSelect("Select Account From Accounts Order By Account")
    If lResult <> 0 Then Exit Sub
    
    ReDim AccountNumbers(0)
    nCount = 0
    Do Until MyTable.EOF
        nCount = nCount + 1
        ReDim Preserve AccountNumbers(nCount)
        AccountNumbers(nCount - 1) = MyTable![Account]
        MyTable.MoveNext
    Loop
    MyTable.Close
    
    nRows = 0
    FormPaymentReport.GridPaymentReport.Rows = nRows + 2

    'set column headings
    FormPaymentReport.GridPaymentReport.Row = 0
    FormPaymentReport.GridPaymentReport.Col = 0
    FormPaymentReport.GridPaymentReport.Text = "Account"
    FormPaymentReport.GridPaymentReport.Col = Columns + 1
    FormPaymentReport.GridPaymentReport.Text = "Total"
    For i = 0 To Columns - 1
        FormPaymentReport.GridPaymentReport.Col = i + 1
        thedate = DateAdd("d", i, CutOffDate)
        FormPaymentReport.GridPaymentReport.Text = Format(thedate, "ddd")
    Next i
   
   
    'set column widths
    FormPaymentReport.GridPaymentReport.ColWidth(0) = 1620
    For i = 1 To 9
        FormPaymentReport.GridPaymentReport.ColWidth(i) = 600
    Next i
   
    'Populate the grid
    For i = 0 To nCount - 1
        'Populate the Account
        Activity = False
        FormPaymentReport.GridPaymentReport.Row = nRows + 1
        FormPaymentReport.GridPaymentReport.Col = 0
        FormPaymentReport.GridPaymentReport.Text = AccountNumbers(i)
        
        'Open the Ledger Table
        SubTotal = 0
        lResult = dbSelect("Select * from Ledger where Account = '" & AccountNumbers(i) & "' and Type = " & BET_TYPE_PAYMENT)
        If lResult <> 0 Then Exit Sub
        
        Do Until MyTable.EOF
            nDays = DateDiff("d", CutOffDate, MyTable![TimeStamp])
            If nDays >= 0 And nDays < Columns Then
                FormPaymentReport.GridPaymentReport.Col = nDays + 1
                FormPaymentReport.GridPaymentReport.Text = MyTable![amount]
                SubTotal = SubTotal + Val(MyTable![amount])
                FormPaymentReport.GridPaymentReport.Col = Columns + 1
                FormPaymentReport.GridPaymentReport.Text = SubTotal
                Activity = True
            Else
                If DateDiff("d", Now(), MyTable![TimeStamp]) = 0 Then
                    PayToday = PayToday + Val(MyTable![amount])
                End If
            End If
            MyTable.MoveNext
        Loop
        MyTable.Close
        
        FormPaymentReport.lToday.Caption = "Payments today:  " & PayToday
                
        If Activity = True Then
            nRows = nRows + 1
            FormPaymentReport.GridPaymentReport.Rows = nRows + 2
        End If
    Next i
    
    FormPaymentReport.GridPaymentReport.Row = nRows + 1
    FormPaymentReport.GridPaymentReport.Col = 0
    FormPaymentReport.GridPaymentReport.Text = "--- Totals ---"

    'Populate totals
    For x = 1 To Columns + 1
        SubTotal = 0
        FormPaymentReport.GridPaymentReport.Col = x
        For i = 0 To nRows - 1
            FormPaymentReport.GridPaymentReport.Row = i + 1
            SubTotal = SubTotal + Val(FormPaymentReport.GridPaymentReport.Text)
        Next i
        FormPaymentReport.GridPaymentReport.Row = i + 1
        FormPaymentReport.GridPaymentReport.Text = SubTotal
    Next x
 
    Form1.MousePointer = vbDefault

End Sub
Public Sub DisplayDaily()

    Dim lResult As Long
    Dim i, x, nCount, WeekBal, SubTotal, nRows, Activity As Long
    Dim Columns As Long
    Dim AccountNumbers() As String
    Dim AccountTotals() As Long
    Dim theAccount As New Account
    Dim CutOffDate As Date, thedate As Date
    Dim DayOfWeek As String

    Form1.MousePointer = vbHourglass

    'Calculate The CutOffDate
    DayOfWeek = Format(Now(), "ddd")
    If DayOfWeek = "Tue" Then Columns = 7
    If DayOfWeek = "Wed" Then Columns = 1
    If DayOfWeek = "Thu" Then Columns = 2
    If DayOfWeek = "Fri" Then Columns = 3
    If DayOfWeek = "Sat" Then Columns = 4
    If DayOfWeek = "Sun" Then Columns = 5
    If DayOfWeek = "Mon" Then Columns = 6
    CutOffDate = DateAdd("d", 0 - Columns, Now())

    'Open the Accounts Table
    lResult = dbSelect("Select * From Accounts Order By Account")
    If lResult <> 0 Then Exit Sub
    
    ReDim AccountNumbers(0)
    ReDim AccountTotals(0)
    nCount = 0
    Do Until MyTable.EOF
        nCount = nCount + 1
        ReDim Preserve AccountNumbers(nCount)
        ReDim Preserve AccountTotals(nCount)
        AccountNumbers(nCount - 1) = MyTable![Account]
        AccountTotals(nCount - 1) = MyTable![balance]
        MyTable.MoveNext
    Loop
    MyTable.Close
    
    nRows = 0
    FormDaily.GridDaily.Rows = nRows + 2

    'set column headings
    FormDaily.GridDaily.Row = 0
    FormDaily.GridDaily.Col = 0
    FormDaily.GridDaily.Text = "Account"
    FormDaily.GridDaily.Col = 1
    FormDaily.GridDaily.Text = "Debit"
    FormDaily.GridDaily.Col = Columns + 2
    FormDaily.GridDaily.Text = "Total"
    For i = 0 To Columns - 1
        FormDaily.GridDaily.Col = i + 2
        thedate = DateAdd("d", i, CutOffDate)
        FormDaily.GridDaily.Text = Format(thedate, "ddd")
    Next i
   
   
    'set column widths
    FormDaily.GridDaily.ColWidth(0) = 1620
    For i = 1 To 9
        FormDaily.GridDaily.ColWidth(i) = 600
    Next i
   
    'Populate the grid
    For i = 0 To nCount - 1
        'Populate the Account
        Activity = False
        FormDaily.GridDaily.Row = nRows + 1
        FormDaily.GridDaily.Col = 0
        FormDaily.GridDaily.Text = AccountNumbers(i)
        
        'Populate Current Balance
        FormDaily.GridDaily.Col = Columns + 2
        FormDaily.GridDaily.Text = AccountTotals(i)
        If AccountTotals(i) <> 0 Then Activity = True
                
        sCurrentAccount = AccountNumbers(i)
        ReadAccount
        theAccount.Populate (sCurrentAccount)
        
        'Populate daily numbers
        WeekBal = 0
        For x = 0 To Columns - 1
            thedate = DateAdd("d", x, CutOffDate)
            FormDaily.GridDaily.Col = x + 2
            FormDaily.GridDaily.Text = theAccount.action(thedate)
            WeekBal = WeekBal + Val(FormDaily.GridDaily.Text)
            If WeekBal <> 0 Then Activity = True
        Next x
        'Populate debit
        FormDaily.GridDaily.Col = 1
        FormDaily.GridDaily.Text = AccountTotals(i) - WeekBal
        
        If Activity = True Then
            nRows = nRows + 1
            FormDaily.GridDaily.Rows = nRows + 2
        End If
    Next i
    FormDaily.GridDaily.Row = nRows + 1
    FormDaily.GridDaily.Col = 0
    FormDaily.GridDaily.Text = "--- Totals ---"

    'Populate totals
    For x = 1 To Columns + 1
        SubTotal = 0
        FormDaily.GridDaily.Col = x
        For i = 0 To nRows - 1
            FormDaily.GridDaily.Row = i + 1
            SubTotal = SubTotal + Val(FormDaily.GridDaily.Text)
        Next i
        FormDaily.GridDaily.Row = i + 1
        FormDaily.GridDaily.Text = SubTotal
    Next x

    'Populate Last Total
    SubTotal = 0
    For x = 2 To Columns + 1
        FormDaily.GridDaily.Col = x
        SubTotal = SubTotal + Val(FormDaily.GridDaily.Text)
    Next x
    FormDaily.GridDaily.Col = Columns + 2
    FormDaily.GridDaily.Text = SubTotal

    Form1.MousePointer = vbDefault

End Sub
