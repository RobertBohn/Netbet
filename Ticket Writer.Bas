Attribute VB_Name = "Module1"
'All variables MUST be defined
Option Explicit

'Constants
Public Const GAMES_PER_TAB = 5     'Games Displayed Per Tab
Public Const MAX_WAGERS = 6        'Wagers Per Ticket
Public Const MAX_MONEY_BUTTONS = 10

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

Public Const WAGER_SPORT = 0
Public Const WAGER_GAME = 1
Public Const WAGER_TYPE = 2
Public Const WAGER_AMOUNT = 3

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
Public nGames(SPORTS) As Long          'Number of Scheduled Games by Sport
Public nLinePosition(2)                '0=Total 1=Line (X-Positions)
Public sWagers(MAX_WAGERS, 3)
Public sSaveWagers(MAX_WAGERS, 3)
Public sParlayAmount As String
Public sCurrentAccount As String
Public lastweek As Long
Public thisweek As Long
Public wagersinaction As Long

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
Public Const PS_SOLID = 0
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const SRCCOPY = &HCC0020   '(DWORD) dest = source
Public Const LF_FACESIZE = 32
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


Type POINTAPI
  x As Long
  Y As Long
End Type

Type Size
  cx As Long
  cy As Long
End Type

Type rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As rect, ByVal hBrush As Long) As Long
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Sub ReadAccounts()
    
    'Read in All Active Accounts
    Dim lResult As Long
    lResult = dbSelect("Select * from Accounts")
    If lResult <> 0 Then Exit Sub

    'Read Active Accounts into the cbAccount ComboBox
    Form1.cbAccount.Clear
    Do Until MyTable.EOF
        If MyTable![Status] = "A" Then Form1.cbAccount.AddItem MyTable![Account]
        MyTable.MoveNext
    Loop

    'Close the Accounts Table
    MyTable.Close

End Sub

Public Sub AccountSelected()
    Dim lResult As Long, i As Long
         
    If Len(Form1.cbAccount.Text) > 12 Then
        Form1.cbAccount.Text = Mid(Form1.cbAccount.Text, 1, 12)
    End If
    
    sCurrentAccount = Form1.cbAccount.Text
    
    'Only Process if Account Has Changed
    If Form1.cbAccount.Text = AccountSettings(ACT_ACCOUNT) Then Exit Sub
      
ReadAccount:
      
    'Set Hourglass Pointer
    Form1.MousePointer = vbHourglass
   
    'Clear the Current Account
    For i = 0 To MAX_ACT_ITEMS - 1
        AccountSettings(i) = ""
    Next i
    
    sParlayAmount = ""
    For i = 0 To MAX_WAGERS - 1
        sWagers(i, WAGER_SPORT) = ""
        sWagers(i, WAGER_GAME) = ""
        sWagers(i, WAGER_TYPE) = ""
        sWagers(i, WAGER_AMOUNT) = ""
    Next i
    
    'Select the Account Record From the Accounts Table
    lResult = dbSelect("Select * From Accounts Where Account = '" & Form1.cbAccount.Text & "'")
    If lResult <> 0 Then
         'Do Nothing on an Error
         Form1.MousePointer = vbDefault
         DisplayTicket
         Exit Sub
    End If
    
    'Check if an Account was Selected
    If MyTable.EOF Then
        'The Selected Account Was Not Found
        MyTable.Close
        Form1.lLimit.Caption = ""
        Form1.lBalance.Caption = ""
        DisplayTicket
        Form1.MousePointer = vbDefault
        
        lResult = MessageBox(0, "Account '" & Form1.cbAccount.Text & "' Was Not Found. Do You Want To Add This Account?", Form1.Caption, MB_YESNO)
        If lResult = IDNO Then Exit Sub
        
        lResult = dbInsert("Insert Into Accounts (Account,Status,Balance) Values (""" & Form1.cbAccount.Text & """, ""A"", 0)")
        If lResult <> 0 Then
            'Do Nothing on an Error
             Exit Sub
        End If
        
        FormBalance.Show 1
        
        Form1.cbAccount.AddItem Form1.cbAccount.Text
        GoTo ReadAccount
    End If

    AccountSettings(ACT_ACCOUNT) = MyTable![Account]
    AccountSettings(ACT_BALANCE) = MyTable![balance]
    lastweek = MyTable![StartingBalance]
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

    If DEPOSIT_SYSTEM = True Then
        AccountSettings(ACT_BALANCE) = Val(AccountSettings(ACT_BALANCE)) - Val(AccountSettings(ACT_IN_ACTION))
    End If
    
    lastweek = Val(AccountSettings(ACT_LAST_WEEK))
    thisweek = Val(AccountSettings(ACT_BALANCE)) - Val(AccountSettings(ACT_LAST_WEEK))
    wagersinaction = Val(AccountSettings(ACT_IN_ACTION))
        
    Form1.lLimit.Caption = "Limit:  $" & AccountSettings(ACT_LIMIT)
    Form1.lBalance.Caption = "Cur Balance: " & Val(AccountSettings(ACT_BALANCE))
    Form1.lLastWeek.Caption = "Last Balance: " & lastweek
    Form1.lThisWeek.Caption = "This Week: " & thisweek
         
    'Read in Schedule to get Line Changes
    'ReadSchedule
        
    DisplayTicket
    Form1.MousePointer = vbDefault
 
End Sub

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
 
    'Set up the Scroll Bars
    For i = 0 To SPORTS - 1
        If nGames(i) > 0 Then
            Form1.SchScroll(i).Max = nGames(i) - 1
            Form1.SchScroll(i).LargeChange = GAMES_PER_TAB - 1
        End If
    Next i
  
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
            If Not IsNull(MyTable![RoadScore]) Then sGames(idx, x + SCH_ROAD_SCORE) = MyTable![RoadScore]
            If Not IsNull(MyTable![HomeScore]) Then sGames(idx, x + SCH_HOME_SCORE) = MyTable![HomeScore]
          
            'Only Fill in the Lines if the Game is Not Underway
            If DateDiff("s", Now, MyTable![GameDate]) > 0 Or WAGER_AFTER_KICKOFF = True Then
                If Len(sGames(idx, x + SCH_ROAD_SCORE)) = 0 And Len(sGames(idx, x + SCH_HOME_SCORE)) = 0 Then
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
            End If
            'Point to the Next Game for That Sport
            nGames(idx) = nGames(idx) + 1
        End If
SKIPIT:
        MyTable.MoveNext
    Loop

    'Close the Schedule Table
    MyTable.Close
    PaintDisplayArea (Form1.SSTab1.Tab)
End Sub



Public Function GameSelected(ByVal Index As Long, ByVal x As Single, ByVal Y As Single) As Long

    Dim nGameHeight As Long  'Height of one Game
    Dim i As Long            'Temporary Counter
    Dim w As Long            'Temporary Counter
    Dim idx As Long       'The Schedule Index
    Dim nBetType As Long  '0=Total 1=Side
    Dim nSide As Long     '0=Road/Over 1=Home/Under
    Dim sType As String

    'Initialize Local Variables
    nGameHeight = ((Form1.DisplayArea(Index).Height - 1) / Screen.TwipsPerPixelY) / GAMES_PER_TAB

    'Convert Mouse X & Y to Screen Units
    x = x / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
      
    'Get Bet Type
    If x >= nLinePosition(1) Then
        'Total Wager
        nBetType = 0
    Else
        If x >= nLinePosition(0) Then
            'Sides Wager
            nBetType = 1
        Else
            'The Click Was Outside The Wager Area
            Exit Function
        End If
    End If
      
    'Work Thru the Games Displayed on the Tab
    idx = Form1.SchScroll(Index).Value
    For i = 0 To GAMES_PER_TAB - 1
        If idx < nGames(Index) And (i * nGameHeight) < Y And ((i + 1) * nGameHeight) > Y Then
            'Found the Game
            
            'Get the side (Top or bottom of game)
            If Y - (i * nGameHeight) < (nGameHeight / 2) Then
                nSide = 0
            Else
                nSide = 1
            End If

            If nBetType = 0 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)) = 0 Then
                'No Total for this game - exit
                Exit Function
            End If
            If nBetType = 1 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)) = 0 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)) = 0 Then
                'No Line for this game - exit
                Exit Function
            End If

            If nBetType = 0 Then
                If nSide = 0 Then
                    sType = "O"
                Else
                    sType = "U"
                End If
            Else
                If nSide = 0 Then
                    sType = "R"
                Else
                    sType = "H"
                End If
            End If
            
            For w = 0 To MAX_WAGERS - 1
                If Len(sWagers(w, WAGER_GAME)) > 0 Then
                    If sWagers(w, WAGER_SPORT) = Index And sWagers(w, WAGER_GAME) = idx Then
                        If ((sType = "O" Or sType = "U") And (sWagers(w, WAGER_TYPE) = "O" Or sWagers(w, WAGER_TYPE) = "U")) _
                        Or ((sType = "H" Or sType = "R") And (sWagers(w, WAGER_TYPE) = "H" Or sWagers(w, WAGER_TYPE) = "R")) Then
                            sWagers(w, WAGER_TYPE) = sType
                            DisplayTicket
                            Exit Function
                        End If
                    End If
                End If
            Next w
            
            For w = 0 To MAX_WAGERS - 1
                If Len(sWagers(w, WAGER_GAME)) = 0 Then
                    'Found 1st Open Wager
                    sWagers(w, WAGER_SPORT) = Index
                    sWagers(w, WAGER_GAME) = idx
                    sWagers(w, WAGER_TYPE) = sType
                    If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = True Then
                        Form1.tWagerAmount(w).Visible = True
                        Form1.tWagerAmount(w).SetFocus
                        
                        'Use Previous Straight Bet Amount
                        'If w > 0 Then
                        '    If Len(sWagers(w, WAGER_AMOUNT)) = 0 And Len(sWagers((w - 1), WAGER_AMOUNT)) > 0 Then
                        '        sWagers(w, WAGER_AMOUNT) = sWagers(w - 1, WAGER_AMOUNT)
                        '    End If
                        'End If
                        
                    End If
                    DisplayTicket
                    Exit Function
                End If
            Next w
            Exit Function
        End If
        idx = idx + 1
    Next i
       
End Function
Public Function DisplaySchedule(ByVal MyDc As Long, ByVal Index As Long) As Long

    'Note: VB Scale is in Twips, Windows GDI API Scale is in Pixels. Be Sure to Convert for Each API Call.
    
    Dim pt As POINTAPI       'MoveToEx Point Structure
    Dim i As Long            'Temporary counter
    Dim sz As Size           'GetTextExtentPoint Size Structure
    Dim nTextHeight As Long  'Height of Text Characters
    Dim nGameHeight As Long  'Height of one Game
    Dim nFreeSpace As Long   'Amount of Free Space in One Game Line
    Dim nLine1 As Long       'Y Position of Road Team Line
    Dim nLine2 As Long       'Y Position of Home Team Line
    Dim yCenter As Long      'Y Position of Center
    Dim idx As Long       'The Game Index
    Dim xDate As Long        'X-Position of Game Date/Time
    Dim xRotation As Long    'X-Position of Rotation Numbers
    Dim xTeamName As Long    'X-Position of Team Names
    Dim xTotalPrice As Long  'Ending X-Position of Total Price
    Dim xTotalLine As Long   'X-Position of Total Line
    Dim xOvUn As Long        'X-Position of OvUn text
    Dim xSidePrice As Long   'Ending X-Position of Side Price
    Dim xSide As Long        'X-Position of Side Line
    Dim xScore As Long       'X-Position of Final Scores
    Dim s As String          'Temporary Working String
    Dim hBrush As Long
    Dim rect As rect

    'Paint the Background
    Call GetClientRect(Form1.DisplayArea(Index).hwnd, rect)
    hBrush = CreateSolidBrush(GetBkColor(MyDc))
    Call FillRect(MyDc, rect, hBrush)
    Call DeleteObject(hBrush)
    
     'Get the Sizes of Various Objects
    Call GetTextExtentPoint(MyDc, "A", 1, sz)
    nGameHeight = ((Form1.DisplayArea(Index).Height - 1) / Screen.TwipsPerPixelY) / GAMES_PER_TAB
    nTextHeight = sz.cy
    nFreeSpace = nGameHeight - (nTextHeight * 2)

    'Calculate Y-Positions of Game Elements
    Call GetTextExtentPoint(MyDc, "X", 1, sz) 'Indent One Space
    xDate = sz.cx
    Call GetTextExtentPoint(MyDc, "XXXX 88/88 ", 11, sz) 'Max Date Size
    xRotation = xDate + sz.cx
    Call GetTextExtentPoint(MyDc, "8888 ", 5, sz) 'Max Rotation Size
    xTeamName = xRotation + sz.cx
    Call GetTextExtentPoint(MyDc, "X", 1, sz) 'Indent One Space From End
    xTotalPrice = (Form1.DisplayArea(Index).Width / Screen.TwipsPerPixelX) - 2 - sz.cx
    Call GetTextExtentPoint(MyDc, "8888 +888", 9, sz)
    xTotalLine = xTotalPrice - sz.cx
    Call GetTextExtentPoint(MyDc, "Ov ", 3, sz)
    xOvUn = xTotalLine - sz.cx
    Call GetTextExtentPoint(MyDc, "XXX", 3, sz)
    xSidePrice = xOvUn - sz.cx
    Call GetTextExtentPoint(MyDc, "888  +888", 9, sz)
    xSide = xSidePrice - sz.cx
    Call GetTextExtentPoint(MyDc, "Finalx8888", 10, sz)
    xScore = xSide + sz.cx
    
      
    nLinePosition(0) = xTeamName
    nLinePosition(1) = xOvUn
      
    'Draw a Box Around the Window (Picture) Area
    Call MoveToEx(MyDc, 0, 0, pt)
    Call LineTo(MyDc, (Form1.DisplayArea(Index).Width / Screen.TwipsPerPixelX) - 2, 0)
    Call LineTo(MyDc, (Form1.DisplayArea(Index).Width / Screen.TwipsPerPixelX) - 2, (Form1.DisplayArea(Index).Height / Screen.TwipsPerPixelY) - 1)
    Call LineTo(MyDc, 0, (Form1.DisplayArea(Index).Height / Screen.TwipsPerPixelY) - 1)
    Call LineTo(MyDc, 0, 0)

    'Draw the Game Seperator Lines
    For i = 1 To GAMES_PER_TAB - 1
        Call MoveToEx(MyDc, 0, 0, pt)
        Call LineTo(MyDc, 0, nGameHeight * i)
        Call LineTo(MyDc, (Form1.DisplayArea(Index).Width / Screen.TwipsPerPixelX) - 1, nGameHeight * i)
    Next i

    'Work Thru the Games Displayed on the Tab
    idx = Form1.SchScroll(Index).Value
    For i = 0 To GAMES_PER_TAB - 1
        If idx < nGames(Index) Then
                
            'Calculate Road/Home Team Line Positions
            nLine1 = (i * nGameHeight) + nFreeSpace / 3
            nLine2 = (i * nGameHeight) + ((nFreeSpace * 2) / 3) + nTextHeight
            yCenter = (i * nGameHeight) + ((nGameHeight - nTextHeight) / 2)
            
            'Display the Game Date
            s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_DATE)
            s = Format(s, "ddd m/d")
            Call TextOut(MyDc, xDate, nLine1, s, Len(s))
                
            'Display the Game Start Time
            s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_DATE)
            s = Format(s, "h:mm AMPM")
            Call TextOut(MyDc, xDate, nLine2, s, Len(s))
                
            'Display the Rotation Numbers
            s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROTATION)
            Call TextOut(MyDc, xRotation, nLine1, s, Len(s))
            s = s + 1
            Call TextOut(MyDc, xRotation, nLine2, s, Len(s))
                
            'Display the Teams
            If xSide > xTeamName Then
                'Be Sure Road Team Does Not OverWrite the Lines
                s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_TEAM)
                Call GetTextExtentPoint(MyDc, s, Len(s), sz)
                Do While sz.cx > xSide - xTeamName
                    s = Mid(s, 1, Len(s) - 1)
                    Call GetTextExtentPoint(MyDc, s, Len(s), sz)
                Loop
                Call TextOut(MyDc, xTeamName, nLine1, s, Len(s))
           
                'Be Sure Home Team Does Not OverWrite the Lines
                s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_TEAM)
                Call GetTextExtentPoint(MyDc, s, Len(s), sz)
                Do While sz.cx > xSide - xTeamName
                    s = Mid(s, 1, Len(s) - 1)
                    Call GetTextExtentPoint(MyDc, s, Len(s), sz)
                Loop
                Call TextOut(MyDc, xTeamName, nLine2, s, Len(s))
            
            End If
            
            If DateDiff("s", Now, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_DATE)) < 0 And WAGER_AFTER_KICKOFF = False Then
                sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE) = ""
                sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL) = ""
                sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE) = ""
                sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE) = ""
                sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_OVER_PRICE) = ""
                sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_UNDER_PRICE) = ""
            End If
                
            'Check If The Game Is Off The Board or Underway
            If Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)) = 0 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)) = 0 _
            And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)) = 0 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)) = 0 _
            And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)) = 0 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)) = 0 Then
                If Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_SCORE)) > 0 And Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_SCORE)) > 0 Then
                    Call TextOut(MyDc, xSide, yCenter, "Final", 5)
                    Call GetTextExtentPoint(MyDc, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_SCORE), Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_SCORE)), sz)
                    Call TextOut(MyDc, xScore - sz.cx, nLine1, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_SCORE), Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_SCORE)))
                    Call GetTextExtentPoint(MyDc, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_SCORE), Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_SCORE)), sz)
                    Call TextOut(MyDc, xScore - sz.cx, nLine2, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_SCORE), Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_SCORE)))
                Else
                    If DateDiff("s", Now, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_DATE)) < 0 And WAGER_AFTER_KICKOFF = False Then
                        Call TextOut(MyDc, xSide, yCenter, "Underway", 8)
                    Else
                        Call TextOut(MyDc, xSide, yCenter, "No Line", 7)
                    End If
                End If
            Else
                'Display Line
                Call FormatLines(Index, idx, " ")
            
                If strRoadLine = "Pick" Then
                    Call TextOut(MyDc, xSide, yCenter, strRoadLine, Len(strRoadLine))
                Else
                    Call TextOut(MyDc, xSide, nLine1, strRoadLine, Len(strRoadLine))
                    Call TextOut(MyDc, xSide, nLine2, strHomeLine, Len(strHomeLine))
                End If

                Call GetTextExtentPoint(MyDc, strRoadSidePrice, Len(strRoadSidePrice), sz)
                Call TextOut(MyDc, xSidePrice - sz.cx, nLine1, strRoadSidePrice, Len(strRoadSidePrice))
    
                Call GetTextExtentPoint(MyDc, strHomeSidePrice, Len(strHomeSidePrice), sz)
                Call TextOut(MyDc, xSidePrice - sz.cx, nLine2, strHomeSidePrice, Len(strHomeSidePrice))

                If Len(strGameTotal) > 0 Then
                    Call TextOut(MyDc, xOvUn, nLine1, "Ov", 2)
                    Call TextOut(MyDc, xOvUn, nLine2, "Un", 2)
                    Call TextOut(MyDc, xTotalLine, yCenter, strGameTotal, Len(strGameTotal))
                End If

                Call GetTextExtentPoint(MyDc, strOverPrice, Len(strOverPrice), sz)
                Call TextOut(MyDc, xTotalPrice - sz.cx, nLine1, strOverPrice, Len(strOverPrice))
    
                Call GetTextExtentPoint(MyDc, strUnderPrice, Len(strUnderPrice), sz)
                Call TextOut(MyDc, xTotalPrice - sz.cx, nLine2, strUnderPrice, Len(strUnderPrice))
            End If
           
            idx = idx + 1
        End If
    Next i

End Function


Public Function PaintDisplayArea(ByVal Index As Long)

    Dim hdc As Long          'VB's Window Device Context
    Dim lResult As Long      'Function Result Codes
    Dim hdcWin As Long       'Memory DC
    Dim hbmWin As Long       'BitMap
    Dim hbmWinOld As Long    'Old BitMap
    Dim x, Y As Long         'Window Size
    Dim lFont As LOGFONT     'Font Discription Structure
    Dim hTempFont As Long    'Temporary Font Handle
    Dim hVbFont As Long      'VB Current Font Handle
    Dim hMemFont As Long     'Memory DC Font Handle
    Dim hOldBkColor As Long  'Old Background Color
    Dim hOldFgColor As Long  'Old Foreground Color
    Dim nPen As Long         'Temporary Pen Object
    Dim nOldPen As Long      'Old Pen Object
     
    'Get the Height and Width of the screen in pixels
    x = Form1.DisplayArea(Index).Width / Screen.TwipsPerPixelX
    Y = Form1.DisplayArea(Index).Height / Screen.TwipsPerPixelY
    
    'Create a Temporary Font
    hTempFont = CreateFontIndirect(lFont)
    
    'Get VB's Window Device Context
    hdc = GetDC(Form1.DisplayArea(Index).hwnd)
    
    'Create a new Device Context Compatible with VB's Device Context
    hdcWin = CreateCompatibleDC(hdc)
    
    'Create a BitMap the same size as VB's window size
    hbmWin = CreateCompatibleBitmap(hdc, x, Y)
    
    'Use this new BitMap with the New Device Context
    hbmWinOld = SelectObject(hdcWin, hbmWin)
    
    'Switch the Temporary Font with VB's Current Font
    hVbFont = SelectObject(hdc, hTempFont)
    
    'Give VB's Font to the Memory Device Context
    hMemFont = SelectObject(hdcWin, hVbFont)
    
    'Set the Memory Device Context Back Ground Color
    hOldBkColor = SetBkColor(hdcWin, GetBkColor(hdc))

    'Set the Memory Device Context Fore Ground Color
    hOldFgColor = SetTextColor(hdcWin, GetTextColor(hdc))
 
    'Create a Pen the same Color as Fore Ground Text
    nPen = CreatePen(PS_SOLID, 0, GetTextColor(hdc))
    
    'Use This New Pen
     nOldPen = SelectObject(hdcWin, nPen)
    
    'Display the Schedule in the Memory Device Context
    lResult = DisplaySchedule(hdcWin, Index)
    
    'Blast the Memory Schedule Image to VB's Screen Display
    lResult = BitBlt(hdc, 0, 0, x, Y, hdcWin, 0, 0, SRCCOPY)
   
    'Reset Memory Device Context Old Pen
    lResult = SelectObject(hdcWin, nOldPen)
     
    'Destroy the Temporary Pen we Created
    lResult = DeleteObject(nPen)
   
    'Reset the Memory Device Context Fore Ground Color
    lResult = SetTextColor(hdcWin, hOldFgColor)
    
    'Reset the Memory Device Context Back Ground Color
    lResult = SetBkColor(hdcWin, hOldBkColor)
   
    'Reset the Memory Device Context's Original Font
    lResult = SelectObject(hdcWin, hMemFont)
    
    'Reset VB's Font
    lResult = SelectObject(hdc, hVbFont)
    
    'Reset the Memory Device's Original BitMap
    lResult = SelectObject(hdcWin, hbmWinOld)
    
    'Delete the Temporary Working BitMap
    lResult = DeleteObject(hbmWin)
    
    'Delete the Temporary Font
    lResult = DeleteObject(hTempFont)
    
    'Delete the Temporary Device Context
    lResult = DeleteDC(hdcWin)
    
    'Release VB's Device Context
    lResult = ReleaseDC(Form1.DisplayArea(Index).hwnd, hdc)

End Function

Public Function FormatLines(ByVal Index As Long, ByVal idx As Long, ByVal BetType As String) As Long

    Dim s As String
    Dim sLine As String
    Dim sTotal As String
    
    strGameTotal = ""
    strOverPrice = ""
    strUnderPrice = ""
    strRoadSidePrice = ""
    strHomeSidePrice = ""
    strRoadLine = ""
    strHomeLine = ""
    
    sLine = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)
    sTotal = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)
    
    'Adjust Ticket Display For Football Teasers
    If Index = NFL_TAB And BetType <> " " Then
        If Form1.oWagerType(BET_TYPE_6TEASER).Value = True Then
            If BetType = "H" Then sLine = Val(sLine) + 6
            If BetType = "R" Then sLine = Val(sLine) - 6
            If BetType = "O" Then sTotal = Val(sTotal) - 4
            If BetType = "U" Then sTotal = Val(sTotal) + 4
        End If
        If Form1.oWagerType(BET_TYPE_7TEASER).Value = True Then
            If BetType = "H" Then sLine = Val(sLine) + 7
            If BetType = "R" Then sLine = Val(sLine) - 7
            If BetType = "O" Then sTotal = Val(sTotal) - 4
            If BetType = "U" Then sTotal = Val(sTotal) + 4
        End If
    End If
    
    'Adjust Ticket Display For Basketball Teasers
    If Index = NBA_TAB And BetType <> " " Then
        If Form1.oWagerType(BET_TYPE_4TEASER).Value = True Then
            If BetType = "H" Then sLine = Val(sLine) + 4
            If BetType = "R" Then sLine = Val(sLine) - 4
        End If
    End If
    
    'NFL and NBA Lines
    s = sLine
    If Len(s) > 0 And (Index = NFL_TAB Or Index = NBA_TAB) Then
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
            If Val(sLine) < 0 Then
                strHomeLine = "-" & s
                strRoadLine = "+" & s
            Else
                strRoadLine = "-" & s
                strHomeLine = "+" & s
            End If
        End If
    End If
    
    'Baseball Home Line
    s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)
    If Len(s) > 0 And Index = MLB_TAB Then
        If Val(s) < 0 Then s = 0 - Val(s)
        If (Val(s) * 10) Mod 10 = 5 Then
             s = Val(s) - 0.5
             s = s & OneHalf
        End If
        If Val(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)) < 0 Then
            strHomeLine = "-" & s
        Else
            strHomeLine = "+" & s
        End If
    End If
    
    'Baseball Road Line
    s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)
    If Len(s) > 0 And Index = MLB_TAB Then
        If Val(s) < 0 Then s = 0 - Val(s)
        If (Val(s) * 10) Mod 10 = 5 Then
             s = Val(s) - 0.5
             s = s & OneHalf
        End If
        If Val(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)) < 0 Then
            strRoadLine = "-" & s
        Else
            strRoadLine = "+" & s
        End If
    End If
    
    'NHL Lines
    s = sLine
    If Len(s) > 0 And Index = NHL_TAB Then
        If Val(s) = 0 Then
            strHomeLine = "ev"
        Else
            If Val(s) < 0 Then s = 0 - Val(s)
            If (Val(s) * 10) Mod 10 = 5 Then
                 s = Val(s) - 0.5
                 If Val(s) = 0 Then s = ""
                 s = s & OneHalf
            End If
            If Val(sLine) < 0 Then
                strHomeLine = "-" & s
            Else
                strHomeLine = "+" & s
            End If
        End If
    
        s = 0 - Val(sLine) - 0.5
        If Val(s) = 0 Then
            strRoadLine = "ev"
        Else
            If Val(s) < 0 Then s = 0 - Val(s)
            If (Val(s) * 10) Mod 10 = 5 Then
                 s = Val(s) - 0.5
                 If Val(s) = 0 Then s = ""
                 s = s & OneHalf
            End If
            If (0 - Val(sLine) - 0.5) < 0 Then
                strRoadLine = "-" & s
            Else
                strRoadLine = "+" & s
            End If
        End If
    End If
    
    'Display Side Prices for NHL Only
    If Index = NHL_TAB Then
         If Val(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)) <> 100 _
         Or Val(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)) <> 100 Then
            'Display Road Side Price
            s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)
            If Len(s) Then
                If Val(s) > 0 Then s = "+" & s
                strRoadSidePrice = s
            End If
                
            'Display Home Side Price
            s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)
            If Len(s) Then
                If Val(s) > 0 Then s = "+" & s
                strHomeSidePrice = s
            End If
        End If
    End If
          
    'Display Totals
    s = sTotal
    If Len(s) > 0 Then
        'Display Total Line
        If (Val(s) * 10) Mod 10 = 5 Then
             s = Val(s) - 0.5
             s = s & OneHalf
        End If
        strGameTotal = s
        
        'Display Over Total Price
        s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_OVER_PRICE)
        If Len(s) And Val(s) <> -110 Then
            If Val(s) = 100 Then
                s = "even"
            Else
                If Val(s) > 0 Then s = "+" & s
            End If
            strOverPrice = s
        End If
                    
        'Display Under Total Price
        s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_UNDER_PRICE)
        If Len(s) And Val(s) <> -110 Then
            If Val(s) = 100 Then
                s = "even"
            Else
                If Val(s) > 0 Then s = "+" & s
            End If
            strUnderPrice = s
        End If
    End If
    
    If BetType = "O" Then _
        strBetDisplay = " " & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROTATION) & " " _
        & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_TEAM) _
        & " Over " & strGameTotal & " " & strOverPrice

    If BetType = "U" Then _
        strBetDisplay = " " & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROTATION) & " " _
        & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_TEAM) _
        & " Under " & strGameTotal & " " & strUnderPrice
   
    If BetType = "R" Then _
        strBetDisplay = " " & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROTATION) & " " _
        & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_TEAM) _
        & " " & strRoadLine & " " & strRoadSidePrice
                    
    If BetType = "H" Then _
        strBetDisplay = " " & Val(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROTATION)) + 1 & " " _
        & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_TEAM) _
        & " " & strHomeLine & " " & strHomeSidePrice
      
    'Do Not Display Dog Lines For NBA or NFL
    If InStr(1, strRoadLine, "+", 1) > 0 And (Index = NFL_TAB Or Index = NBA_TAB) Then strRoadLine = ""
    If InStr(1, strHomeLine, "+", 1) > 0 And (Index = NFL_TAB Or Index = NBA_TAB) Then strHomeLine = ""

End Function

Public Sub DisplayTicket()

    Dim i As Long, x As Long
    Dim s As String
    Dim GamesSelected As Long, action As Long
    
    'First Check Account - If No Account Allow Nothing
    If Len(AccountSettings(ACT_ACCOUNT)) = 0 Then
        Form1.pbWager.Enabled = False
        Form1.pbAccountDetail.Enabled = False
        Form1.lParlayName.Caption = ""
    
        For i = 0 To MAX_WAGERS - 1
            Form1.lWagerDescr(i).Caption = ""
            Form1.tWagerAmount(i).Text = ""
            Form1.lWagerDescr(i).Visible = False
            Form1.tWagerAmount(i).Visible = False
        Next i
        
        For i = 0 To MAX_MONEY_BUTTONS - 1
            Form1.pbMoney(i).Enabled = False
        Next i
        
        Form1.StatusBar1.SimpleText = "Select An Account"
        Exit Sub
    End If
    Form1.pbAccountDetail.Enabled = True

    'Describe Parlay Type
    If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = True Then Form1.lParlayName.Caption = ""
    If Form1.oWagerType(BET_TYPE_PARLAY).Value = True Then Form1.lParlayName.Caption = " Parlay"
    If Form1.oWagerType(BET_TYPE_4TEASER).Value = True Then Form1.lParlayName.Caption = " 4 Point Teaser"
    If Form1.oWagerType(BET_TYPE_6TEASER).Value = True Then Form1.lParlayName.Caption = " 6 Point Teaser"
    If Form1.oWagerType(BET_TYPE_7TEASER).Value = True Then Form1.lParlayName.Caption = " 7 Point Teaser"
    
    'Enable Parlay Amount For Parlays & Teasers
    If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = True Then
        Form1.tParlayAmount.Visible = False
    Else
        If Len(sParlayAmount) > 0 Then
            Form1.tParlayAmount.Text = "$" & sParlayAmount
        Else
            Form1.tParlayAmount.Text = ""
        End If
        Form1.tParlayAmount.Visible = True
    End If
    
    'Display Selected Games
    GamesSelected = 0
    For i = 0 To MAX_WAGERS - 1
        If Len(sWagers(i, WAGER_GAME)) > 0 Then
            GamesSelected = GamesSelected + 1
            Call FormatLines(sWagers(i, WAGER_SPORT), sWagers(i, WAGER_GAME), sWagers(i, WAGER_TYPE))
            
            'Replace "&" With "&&" or VB will think a shortcut was defined
            s = ""
            For x = 0 To Len(strBetDisplay) - 1
                If Mid(strBetDisplay, x + 1, 1) = "&" Then
                    s = s & "&&"
                Else
                    s = s & Mid(strBetDisplay, x + 1, 1)
                End If
            Next x
            Form1.lWagerDescr(i) = s
            
            Form1.lWagerDescr(i).Visible = True
            If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = True Then
                'Display Straight Bet Amounts
                If Len(sWagers(i, WAGER_AMOUNT)) > 0 Then
                    Form1.tWagerAmount(i).Text = "$" & sWagers(i, WAGER_AMOUNT)
                Else
                    Form1.tWagerAmount(i).Text = ""
                End If
                Form1.tWagerAmount(i).Visible = True
            Else
                Form1.tWagerAmount(i).Visible = False
            End If
        Else
            Form1.lWagerDescr(i).Visible = False
            Form1.tWagerAmount(i).Visible = False
        End If
    Next i
    
    'Manage Money Buttons
    If GamesSelected > 0 Or Form1.oWagerType(BET_TYPE_STRAIGHT).Value = False Then
        For i = 0 To MAX_MONEY_BUTTONS - 1
        '    If ((i + 1) * 50) <= Val(AccountSettings(ACT_LIMIT)) Then
                Form1.pbMoney(i).Enabled = True
        '    Else
        '        Form1.pbMoney(i).Enabled = False
        '    End If
        Next i
    Else
        For i = 0 To MAX_MONEY_BUTTONS - 1
            Form1.pbMoney(i).Enabled = False
        Next i
    End If
    
    'Manage Wager Button
    If GamesSelected > 0 Then
        If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = False And GamesSelected = 1 Then
            Form1.StatusBar1.SimpleText = "Select Games"
            Form1.pbWager.Enabled = False
            Exit Sub
        End If
    Else
        Form1.StatusBar1.SimpleText = "Select Games"
        Form1.pbWager.Enabled = False
        Exit Sub
    End If
   
    'Wager Amounts Are Required
    If Form1.tParlayAmount.Visible = True And Val(sParlayAmount) <= 0 Then
        Form1.StatusBar1.SimpleText = "Enter Wager Amount"
        Form1.pbWager.Enabled = False
        Exit Sub
    End If
    For i = 0 To MAX_WAGERS - 1
       If Form1.tWagerAmount(i).Visible = True And Val(sWagers(i, WAGER_AMOUNT)) <= 0 Then
            Form1.StatusBar1.SimpleText = "Enter Wager Amount"
            Form1.pbWager.Enabled = False
            Exit Sub
        End If
    Next i
  
    'Wager Amounts Can Not Exceed Betting Limit
    If Form1.tParlayAmount.Visible = True And Val(sParlayAmount) > Val(AccountSettings(ACT_LIMIT)) Then
        Form1.StatusBar1.SimpleText = "Can't Bet Over The Limit"
        Form1.pbWager.Enabled = False
        Exit Sub
    End If
    For i = 0 To MAX_WAGERS - 1
       If Form1.tWagerAmount(i).Visible = True And Val(sWagers(i, WAGER_AMOUNT)) > Val(AccountSettings(ACT_LIMIT)) Then
            Form1.StatusBar1.SimpleText = "Can't Bet Over The Limit"
            Form1.pbWager.Enabled = False
            Exit Sub
        End If
    Next i
  
    'Can Not Parlay More Than 4 Games
    If GamesSelected > 4 And Form1.oWagerType(BET_TYPE_PARLAY).Value = True Then
        Form1.StatusBar1.SimpleText = "Can't Parlay More Than 4 Games"
        Form1.pbWager.Enabled = False
        Exit Sub
    End If
  
    'Can Not Tease More Than 3 Games
    If GamesSelected > 3 And (Form1.oWagerType(BET_TYPE_6TEASER).Value = True Or Form1.oWagerType(BET_TYPE_7TEASER).Value = True Or Form1.oWagerType(BET_TYPE_4TEASER).Value = True) Then
        Form1.StatusBar1.SimpleText = "Can't Tease More Than 3 Games"
        Form1.pbWager.Enabled = False
        Exit Sub
    End If
    
    'Can Only 6-7 Tease Football
    If Form1.oWagerType(BET_TYPE_6TEASER).Value = True Or Form1.oWagerType(BET_TYPE_7TEASER).Value = True Then
        For i = 0 To MAX_WAGERS - 1
            If Len(sWagers(i, WAGER_GAME)) > 0 Then
                If sWagers(i, WAGER_SPORT) <> NFL_TAB Then
                    Form1.StatusBar1.SimpleText = "6-7 Point Teasers Are For Football Only"
                    Form1.pbWager.Enabled = False
                    Exit Sub
                End If
            End If
        Next i
    End If

    '6-7Pt Football Teaser Total Must Be With The Same Game
    If Form1.oWagerType(BET_TYPE_6TEASER).Value = True Or Form1.oWagerType(BET_TYPE_7TEASER).Value = True Then
        For i = 0 To MAX_WAGERS - 1
            If Len(sWagers(i, WAGER_GAME)) > 0 Then
                If sWagers(i, WAGER_TYPE) = "O" Or sWagers(i, WAGER_TYPE) = "U" Then
                    
                    'Must Be Two Games
                    If GamesSelected <> 2 Then
                        Form1.StatusBar1.SimpleText = "Football Totals Can Only Be Teased With The Side Of The Same Game."
                        Form1.pbWager.Enabled = False
                        Exit Sub
                    End If
                    
                    'Must Find Side Of Same Game
                    For x = 0 To MAX_WAGERS - 1
                        If Len(sWagers(x, WAGER_GAME)) > 0 And x <> i Then
                            If (sWagers(x, WAGER_TYPE) <> "H" And sWagers(x, WAGER_TYPE) <> "R") Or (sWagers(x, WAGER_GAME) <> sWagers(i, WAGER_GAME)) Then
                                Form1.StatusBar1.SimpleText = "Football Totals Can Only Be Teased With The Side Of The Same Game."
                                Form1.pbWager.Enabled = False
                                Exit Sub
                            End If
                        End If
                    Next x
                
                End If
            End If
        Next i
    End If

    'Can Only 4 Point Tease BasketBall Sides
    If Form1.oWagerType(BET_TYPE_4TEASER).Value = True Then
        For i = 0 To MAX_WAGERS - 1
            If Len(sWagers(i, WAGER_GAME)) > 0 Then
                If sWagers(i, WAGER_SPORT) <> NBA_TAB Or sWagers(i, WAGER_TYPE) = "O" Or sWagers(i, WAGER_TYPE) = "U" Then
                    Form1.StatusBar1.SimpleText = "4 Point Teasers Are For Basketball Sides Only"
                    Form1.pbWager.Enabled = False
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    'Can't bet more than your balance
    If DEPOSIT_SYSTEM = True Then
        action = 0
        If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = True Then
            For i = 0 To MAX_WAGERS - 1
                If Len(sWagers(i, WAGER_GAME)) > 0 Then
                    'Football and Basketball
                    If sWagers(i, WAGER_SPORT) = NFL_TAB Or sWagers(i, WAGER_SPORT) = NBA_TAB Then
                        action = action - ((Val(sWagers(i, WAGER_AMOUNT)) * (0 - Val(AccountSettings(ACT_STRAIGHT_BET)))) / 100)
                    End If
                    'Baseball Sides
                    If sWagers(i, WAGER_SPORT) = MLB_TAB Then
                        If sWagers(i, WAGER_TYPE) = "H" Then
                            If Val(sGames(MLB_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_HOME_PRICE)) > 0 Then
                                action = action - Val(sWagers(i, WAGER_AMOUNT))
                            Else
                                action = action + (Val(sWagers(i, WAGER_AMOUNT)) * Val(sGames(MLB_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_HOME_PRICE))) / 5
                            End If
                        End If
                        If sWagers(i, WAGER_TYPE) = "R" Then
                            If Val(sGames(MLB_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)) > 0 Then
                                action = action - Val(sWagers(i, WAGER_AMOUNT))
                            Else
                                action = action + (Val(sWagers(i, WAGER_AMOUNT)) * Val(sGames(MLB_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_ROAD_PRICE))) / 5
                            End If
                        End If
                    End If
                    'Baseball and Hockey Totals
                    If sWagers(i, WAGER_SPORT) = NHL_TAB Or sWagers(i, WAGER_SPORT) = MLB_TAB Then
                        If sWagers(i, WAGER_TYPE) = "O" Then
                            If Val(sGames(sWagers(i, WAGER_SPORT), (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_OVER_PRICE)) = 0 Then
                                action = action - ((Val(sWagers(i, WAGER_AMOUNT)) * (0 - Val(AccountSettings(ACT_STRAIGHT_BET)))) / 100)
                            Else
                                If Val(sGames(sWagers(i, WAGER_SPORT), (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_OVER_PRICE)) > 0 Then
                                    action = action - Val(sWagers(i, WAGER_AMOUNT))
                                Else
                                    action = action + (Val(sWagers(i, WAGER_AMOUNT)) * Val(sGames(sWagers(i, WAGER_SPORT), (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_OVER_PRICE))) / 100
                                End If
                            End If
                        End If
                        If sWagers(i, WAGER_TYPE) = "U" Then
                            If Val(sGames(sWagers(i, WAGER_SPORT), (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_UNDER_PRICE)) = 0 Then
                                action = action - ((Val(sWagers(i, WAGER_AMOUNT)) * (0 - Val(AccountSettings(ACT_STRAIGHT_BET)))) / 100)
                            Else
                                If Val(sGames(sWagers(i, WAGER_SPORT), (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_UNDER_PRICE)) > 0 Then
                                    action = action - Val(sWagers(i, WAGER_AMOUNT))
                                Else
                                    action = action + (Val(sWagers(i, WAGER_AMOUNT)) * Val(sGames(sWagers(i, WAGER_SPORT), (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_UNDER_PRICE))) / 100
                                End If
                            End If
                        End If
                    End If
                    'Hockey Sides
                    If sWagers(i, WAGER_SPORT) = NHL_TAB Then
                        If sWagers(i, WAGER_TYPE) = "H" Then
                            If Val(sGames(NHL_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_HOME_PRICE)) >= 0 Then
                                action = action - Val(sWagers(i, WAGER_AMOUNT))
                            Else
                                action = action + (Val(sWagers(i, WAGER_AMOUNT)) * Val(sGames(NHL_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_HOME_PRICE))) / 100
                           End If
                        End If
                        If sWagers(i, WAGER_TYPE) = "R" Then
                            If Val(sGames(NHL_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)) >= 0 Then
                                action = action - Val(sWagers(i, WAGER_AMOUNT))
                            Else
                                action = action + (Val(sWagers(i, WAGER_AMOUNT)) * Val(sGames(NHL_TAB, (sWagers(i, WAGER_GAME) * MAX_SCH_ITEMS) + SCH_ROAD_PRICE))) / 100
                           End If
                        End If
                    End If
                End If
            Next i
            If (0 - action) > Val(AccountSettings(ACT_BALANCE)) Then
                Form1.StatusBar1.SimpleText = "Can't Bet More Than The Account Balance"
                Form1.pbWager.Enabled = False
                Exit Sub
            End If
        End If
        If Form1.oWagerType(BET_TYPE_PARLAY).Value = True Then
            action = action - Val(sParlayAmount)
        End If
        If Form1.oWagerType(BET_TYPE_4TEASER).Value = True Then
            If GamesSelected = 2 And Val(AccountSettings(ACT_4PT_TEASER2)) < 0 Then
                action = action - ((Val(sParlayAmount) * (0 - Val(AccountSettings(ACT_4PT_TEASER2)))) / 100)
            Else
                action = action - Val(sParlayAmount)
            End If
        End If
        If Form1.oWagerType(BET_TYPE_6TEASER).Value = True Then
            If GamesSelected = 2 And Val(AccountSettings(ACT_6PT_TEASER2)) < 0 Then
                action = action - ((Val(sParlayAmount) * (0 - Val(AccountSettings(ACT_6PT_TEASER2)))) / 100)
            Else
                action = action - Val(sParlayAmount)
            End If
        End If
        If Form1.oWagerType(BET_TYPE_7TEASER).Value = True Then
            If GamesSelected = 2 And Val(AccountSettings(ACT_7PT_TEASER2)) < 0 Then
                action = action - ((Val(sParlayAmount) * (0 - Val(AccountSettings(ACT_7PT_TEASER2)))) / 100)
            Else
                action = action - Val(sParlayAmount)
            End If
        End If
        If (0 - action) > Val(AccountSettings(ACT_BALANCE)) Then
            Form1.StatusBar1.SimpleText = "Can't Bet More Than The Account Balance"
            Form1.pbWager.Enabled = False
            Exit Sub
        End If
    End If
    
    'Describe Parlay Type
    If Form1.oWagerType(BET_TYPE_STRAIGHT).Value = False Then
        Form1.lParlayName.Caption = " " & GamesSelected & " Team" & Form1.lParlayName.Caption
    End If
    
    'Looks Good, Accept The Wager
    Form1.StatusBar1.SimpleText = ""
    Form1.pbWager.Enabled = True

End Sub

Public Sub DisplayAccountDetail()

    Form1.MousePointer = vbHourglass
       
    Call FormDetail.theAccount.Populate(Form1.cbAccount.Text)
    Call FormDetail.theAccount.Display
 
    FormDetail.lBackground(0).Caption = "Account: " & Form1.cbAccount.Text
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
 
    Form1.MousePointer = vbDefault

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
    DisplayTicket
End Sub

Public Function LineSelected(ByVal Index As Long, ByVal x As Single, ByVal Y As Single) As Long
    Dim nGameHeight As Long  'Height of one Game
    Dim i As Long            'Temporary Counter
    Dim w As Long            'Temporary Counter
    Dim idx As Long       'The Schedule Index
    Dim nBetType As Long  '0=Total 1=Side
    Dim nSide As Long     '0=Road/Over 1=Home/Under
    Dim sType As String
    Dim s As String

    'Initialize Local Variables
    nGameHeight = ((Form1.DisplayArea(Index).Height - 1) / Screen.TwipsPerPixelY) / GAMES_PER_TAB

    'Convert Mouse X & Y to Screen Units
    x = x / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
      
    'Get Bet Type
    If x >= nLinePosition(1) Then
        'Total Wager
        nBetType = 0
    Else
        If x >= nLinePosition(0) Then
            'Sides Wager
            nBetType = 1
        Else
            'The Click Was Outside The Wager Area
            Exit Function
        End If
    End If
      
    'Work Thru the Games Displayed on the Tab
    idx = Form1.SchScroll(Index).Value
    For i = 0 To GAMES_PER_TAB - 1
        If idx < nGames(Index) And (i * nGameHeight) < Y And ((i + 1) * nGameHeight) > Y Then
            'Found the Game
            
            'Get the side (Top or bottom of game)
            If Y - (i * nGameHeight) < (nGameHeight / 2) Then
                nSide = 0
            Else
                nSide = 1
            End If
            
            If WAGER_AFTER_KICKOFF = False Then
                If DateDiff("s", Now, sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_DATE)) < 0 Then
                    Exit Function
                End If
                If Len(sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_SCORE)) > 0 Then
                    Exit Function
                End If
            End If
            
            If Index = NFL_TAB Or Index = NBA_TAB Then
                Load FormEdit
                If nBetType = 0 Then
                    FormEdit.Text1.Text = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)
                    FormEdit.Top = Form1.Top + Form1.SSTab1.Top + Form1.DisplayArea(Index).Top + ((i * nGameHeight) * Screen.TwipsPerPixelY) + (FormEdit.Text1.Height / 2) + ((nGameHeight * Screen.TwipsPerPixelY) / 2)
                    FormEdit.Left = Form1.Left + Form1.SSTab1.Left + Form1.DisplayArea(Index).Left + (nLinePosition(1) * Screen.TwipsPerPixelX)
                Else
                    FormEdit.Top = Form1.Top + Form1.SSTab1.Top + Form1.DisplayArea(Index).Top + ((i * nGameHeight) * Screen.TwipsPerPixelY) + FormEdit.Text1.Height + ((nGameHeight * Screen.TwipsPerPixelY) / 2)
                    FormEdit.Left = Form1.Left + Form1.SSTab1.Left + Form1.DisplayArea(Index).Left + (nLinePosition(1) * Screen.TwipsPerPixelX) - (2 * FormEdit.Text1.Width)
                    FormEdit.Text1.Text = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)
                End If
                FormEdit.Text1.SelStart = 0
                FormEdit.Text1.SelLength = Len(FormEdit.Text1.Text)
                FormEdit.GameIndex = Index
                FormEdit.GameOffset = idx
                FormEdit.BetType = nBetType
                FormEdit.Show
            End If
            Exit Function
         End If
        idx = idx + 1
    Next i
 
End Function
