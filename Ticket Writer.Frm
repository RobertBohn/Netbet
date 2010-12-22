VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket Writer - For Entertainment Purposes"
   ClientHeight    =   6285
   ClientLeft      =   135
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "Ticket Writer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6285
   ScaleWidth      =   8175
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   6600
      TabIndex        =   54
      Top             =   3840
      Width           =   1455
      Begin VB.CommandButton pbExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   0
         TabIndex        =   58
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton pbDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   0
         TabIndex        =   57
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton pbChangeLines 
         Caption         =   "&Line Changes"
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton pbAccountDetail 
         Caption         =   "&Account Detail"
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FrameTicket 
      Caption         =   "Ticket"
      Height          =   2655
      Left            =   2040
      TabIndex        =   35
      Top             =   0
      Width           =   4455
      Begin VB.TextBox tParlayAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3720
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tWagerAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   3720
         TabIndex        =   47
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tWagerAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   3720
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tWagerAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   3720
         TabIndex        =   45
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tWagerAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tWagerAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   43
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tWagerAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   37
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton pbWager 
         Caption         =   "&Wager"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lParlayName 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lWagerDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   2280
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lWagerDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   2040
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lWagerDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lWagerDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lWagerDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lWagerDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.Frame FrameWagerAmount 
      Caption         =   "Amount"
      Height          =   2175
      Left            =   6600
      TabIndex        =   24
      Top             =   1680
      Width           =   1455
      Begin VB.CommandButton pbMoney 
         Caption         =   "$10"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$100"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$20"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$200"
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$30"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$300"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$40"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$400"
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$50"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton pbMoney 
         Caption         =   "$500"
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.Frame FrameWagetType 
      Caption         =   "Type"
      Height          =   1575
      Left            =   6600
      TabIndex        =   19
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton oWagerType 
         Caption         =   "7 pt Teaser"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton oWagerType 
         Caption         =   "6 pt Teaser"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton oWagerType 
         Caption         =   "4 pt Teaser"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton oWagerType 
         Caption         =   "Parlay"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton oWagerType 
         Caption         =   "Straight"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame FrameFindGame 
      Caption         =   "Find Game"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1815
      Begin VB.TextBox tNumber 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox tTeam 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lByNumber 
         Caption         =   "By Number"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lByTeam 
         BackStyle       =   0  'Transparent
         Caption         =   "By Team"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame FrameAccount 
      Caption         =   "Account"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   1815
      Begin VB.ComboBox cbAccount 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lThisWeek 
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lLastWeek 
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lLimit 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lBalance 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "&Football"
      TabPicture(0)   =   "Ticket Writer.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DisplayArea(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SchScroll(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Basketball"
      TabPicture(1)   =   "Ticket Writer.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DisplayArea(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SchScroll(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ba&seball"
      TabPicture(2)   =   "Ticket Writer.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DisplayArea(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SchScroll(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Hockey"
      TabPicture(3)   =   "Ticket Writer.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SchScroll(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "DisplayArea(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.VScrollBar SchScroll 
         Height          =   2655
         Index           =   3
         Left            =   -69000
         Max             =   0
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox DisplayArea 
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2655
         Index           =   3
         Left            =   -74880
         ScaleHeight     =   2655
         ScaleWidth      =   5895
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
      Begin VB.VScrollBar SchScroll 
         Height          =   2655
         Index           =   2
         Left            =   -69000
         Max             =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox DisplayArea 
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2655
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   2655
         ScaleWidth      =   5895
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
      Begin VB.VScrollBar SchScroll 
         Height          =   2655
         Index           =   1
         Left            =   -69000
         Max             =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox DisplayArea 
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2655
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   2655
         ScaleWidth      =   5895
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
      Begin VB.VScrollBar SchScroll 
         Height          =   2655
         Index           =   0
         Left            =   6000
         Max             =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox DisplayArea 
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2655
         Index           =   0
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   5895
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   50
      Top             =   6000
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit


Private Sub cbAccount_Click()
    AccountSelected
End Sub


Private Sub cbAccount_KeyPress(KeyAscii As Integer)
    
    'Convert to uppercase
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = KeyAscii - Asc("a") + Asc("A")
    
    'Select an Account when Enter is pressed
    If KeyAscii = vbKeyReturn Then AccountSelected

End Sub




Private Sub cbAccount_LostFocus()
    AccountSelected
End Sub

Private Sub DisplayArea_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  
    'Do Not Let The DisplayArea Keep The Focus
    SSTab1.SetFocus
     
    'Process a Game Wager Left Mouse Button Click
    If Button = vbLeftButton Then Call GameSelected(Index, x, Y)
    If Button = vbRightButton Then Call LineSelected(Index, x, Y)

End Sub


Private Sub DisplayArea_Paint(Index As Integer)
    PaintDisplayArea (Index)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim x As Long


    If KeyCode = vbKeyF2 Then
        For i = 0 To MAX_WAGERS
            For x = 0 To 3
                sWagers(i, x) = sSaveWagers(i, x)
            Next x
        Next i
        DisplayTicket
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    
    'Initialize
    DEPOSIT_SYSTEM = GetSetting("Ticket Writer", "Settings", "DepositSystem", 0)
    ALLOW_DATE_TO_BE_SET_BACK = GetSetting("Ticket Writer", "Settings", "SetBackDate", 0)
    DELETE_ALL_BUTTON = GetSetting("Ticket Writer", "Settings", "DeleteAllButton", 1)
    NBA_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "NBATodayOnly", 0)
    NHL_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "NHLTodayOnly", 0)
    MLB_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "MLBTodayOnly", 0)
    NFL_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "NFLTodayOnly", 0)
    ALL_TIES_PUSH = GetSetting("Ticket Writer", "Settings", "AllTiesPush", 0)
    STOP_WORKING_DATE = GetSetting("Ticket Writer", "Settings", "StopWorkingDate", "1/1/95")
    MY_PASSWORD = GetSetting("Ticket Writer", "Settings", "Password", "")
    WAGER_AFTER_KICKOFF = GetSetting("Ticket Writer", "Settings", "WagersAfterKickoff", 0)
    MY_PASSWORD = GetSetting("Ticket Writer", "Settings", "Password", 0)
    DATABASE = App.Path & "\Ticket Writer.mdb"
    TOTALS_SHEET = App.Path & "\Totals.xls"
    DAILY_SHEET = App.Path & "\Daily.xls"
    PAYMENT_SHEET = App.Path & "\Payment.xls"
    LINES_SHEET = App.Path & "\Print Lines.xls"
    DETAIL_SHEET = App.Path & "\Print Account.xls"
    ADJUSTMENTS_SHEET = App.Path & "\Adjustments.xls"
    
    'Open the Database
    If dbOpen <> 0 Then
        Unload Me
        Exit Sub
    End If
    
    'Display Ticket Area
    DisplayTicket
    
    'Read in Schedule
    ReadSchedule
    
    'Get System Default Values
    ReadDefaults
    
    'Populate cbAccounts Combo Box With Active Accounts
    ReadAccounts

    If DELETE_ALL_BUTTON = False Then Form1.pbDelete.Visible = False

End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    'Close the Database
    On Error Resume Next
    dbClose
    End
    
End Sub






Private Sub lWagerDescr_DblClick(Index As Integer)
    sWagers(Index, WAGER_GAME) = ""
    sWagers(Index, WAGER_TYPE) = ""
    sWagers(Index, WAGER_AMOUNT) = ""
    DisplayTicket
End Sub


Private Sub oWagerType_Click(Index As Integer)
    DisplayTicket
End Sub

Private Sub pbAccountDetail_Click()
    DisplayAccountDetail
    FormDetail.Show 1

    wagersinaction = Val(AccountSettings(ACT_IN_ACTION))
    Call dbUpdate("Update Accounts Set InAction = " & wagersinaction & " Where Account = '" & sCurrentAccount & "'")
    Form1.lBalance.Caption = "Cur Balance: " & Val(AccountSettings(ACT_BALANCE))
    thisweek = Val(AccountSettings(ACT_BALANCE)) - lastweek
    Form1.lLastWeek.Caption = "Last Balance: " & lastweek
    Form1.lThisWeek.Caption = "This Week: " & thisweek
    
End Sub

Private Sub pbChangeLines_Click()
    Me.MousePointer = vbHourglass
    ReadSchedule
    Me.MousePointer = vbDefault
    
    DisplayLineChanges
    FormLines.Show 1
End Sub

Private Sub pbDelete_Click()
    FormDelete.Show 1
End Sub

Private Sub pbExit_Click()
    Unload Form1
End Sub

Private Sub pbMoney_Click(Index As Integer)
    Dim i As Long
    Dim howmuch As Long
    
    If Index = 0 Then howmuch = 10
    If Index = 1 Then howmuch = 100
    If Index = 2 Then howmuch = 20
    If Index = 3 Then howmuch = 200
    If Index = 4 Then howmuch = 30
    If Index = 5 Then howmuch = 300
    If Index = 6 Then howmuch = 40
    If Index = 7 Then howmuch = 400
    If Index = 8 Then howmuch = 50
    If Index = 9 Then howmuch = 500
    
    If oWagerType(BET_TYPE_STRAIGHT).Value = False Then
        'Amount Must Apply To a Parlay Bet
        sParlayAmount = Val(sParlayAmount) + howmuch
        DisplayTicket
    Else
        'Find First Staright With No Amount
        For i = 0 To MAX_WAGERS
            If Len(sWagers(i, WAGER_GAME)) > 0 And Len(sWagers(i, WAGER_AMOUNT)) = 0 Then
                sWagers(i, WAGER_AMOUNT) = Val(sWagers(i, WAGER_AMOUNT)) + howmuch
                DisplayTicket
                Exit Sub
            End If
        Next i
        
        'Find Last Straight Bet
        For i = 0 To MAX_WAGERS
            If Len(sWagers(MAX_WAGERS - i, WAGER_GAME)) > 0 Then
                sWagers(MAX_WAGERS - i, WAGER_AMOUNT) = Val(sWagers(MAX_WAGERS - i, WAGER_AMOUNT)) + howmuch
                DisplayTicket
                Exit Sub
            End If
        Next i
        
    End If
    
End Sub

Private Sub pbWager_Click()
    
    Dim i As Long
    Dim x As Long
    Dim lResult As Long
    Dim SQL As String
    Dim s As String
    Dim Index, idx, action, teams As Long
    
    
    For i = 0 To MAX_WAGERS
        For x = 0 To 3
            sSaveWagers(i, x) = sWagers(i, x)
        Next x
    Next i
    
    'Insert One Ledger Row For Each Straight Bet
    If oWagerType(BET_TYPE_STRAIGHT).Value = True Then
        For i = 0 To MAX_WAGERS - 1
            If Len(sWagers(i, WAGER_GAME)) > 0 Then
            
                Index = sWagers(i, WAGER_SPORT)
                idx = sWagers(i, WAGER_GAME)
            
                SQL = "Insert Into Ledger(Account,Type,Amount,Odds,Game1,Side1,Line1,Price1) Values ("
                SQL = SQL & """" & Form1.cbAccount.Text & """"
                SQL = SQL & "," & BET_TYPE_STRAIGHT
                SQL = SQL & "," & sWagers(i, WAGER_AMOUNT)
                SQL = SQL & "," & AccountSettings(ACT_STRAIGHT_BET)
                SQL = SQL & "," & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                SQL = SQL & "," & """" & sWagers(i, WAGER_TYPE) & """"
                
                If sWagers(i, WAGER_TYPE) = "H" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)
                If sWagers(i, WAGER_TYPE) = "R" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)
                If sWagers(i, WAGER_TYPE) = "O" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)
                If sWagers(i, WAGER_TYPE) = "U" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)
                If Len(s) > 0 Then
                    SQL = SQL & "," & s
                Else
                    SQL = SQL & "," & """"""
                End If
                
                If sWagers(i, WAGER_TYPE) = "H" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)
                If sWagers(i, WAGER_TYPE) = "R" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)
                If sWagers(i, WAGER_TYPE) = "O" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_OVER_PRICE)
                If sWagers(i, WAGER_TYPE) = "U" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_UNDER_PRICE)
                If Len(s) > 0 Then
                    SQL = SQL & "," & s
                Else
                    SQL = SQL & "," & """"""
                End If
                  
                SQL = SQL & ")"
    
                lResult = dbInsert(SQL)
                If lResult <> 0 Then Exit Sub
            End If
        Next i
    Else
        'Insert One Parlay/Teaser Ledger Row
        'First Get Number of Games
        lResult = 0
        teams = 0
        For i = 0 To MAX_WAGERS - 1
            If Len(sWagers(i, WAGER_GAME)) > 0 Then
                lResult = lResult + 1
                teams = teams + 1
            End If
        Next i

        SQL = "Insert Into Ledger(Account,Type,Amount,Odds,Game1,Side1,Line1,Price1,Game2,Side2,Line2,Price2"
        If lResult > 2 Then SQL = SQL & ",Game3,Side3,Line3,Price3"
        If lResult > 3 Then SQL = SQL & ",Game4,Side4,Line4,Price4"
        SQL = SQL & ") Values ("
     
        'Account Number
        SQL = SQL & """" & Form1.cbAccount.Text & """"
                
        'Wager Type
        If oWagerType(BET_TYPE_PARLAY).Value = True Then SQL = SQL & "," & BET_TYPE_PARLAY
        If oWagerType(BET_TYPE_4TEASER).Value = True Then SQL = SQL & "," & BET_TYPE_4TEASER
        If oWagerType(BET_TYPE_6TEASER).Value = True Then SQL = SQL & "," & BET_TYPE_6TEASER
        If oWagerType(BET_TYPE_7TEASER).Value = True Then SQL = SQL & "," & BET_TYPE_7TEASER
                
        'Wager Amount
        SQL = SQL & "," & sParlayAmount
                
        'Payoff Odds
        SQL = SQL & "," & AccountSettings(ACT_STRAIGHT_BET)
     
        'Get Game Information
        For i = 0 To MAX_WAGERS - 1
            If Len(sWagers(i, WAGER_GAME)) > 0 Then
            
                Index = sWagers(i, WAGER_SPORT)
                idx = sWagers(i, WAGER_GAME)
                
                'Game Number
                SQL = SQL & "," & sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_GAME_NBR)
                
                'Wager Type
                SQL = SQL & "," & """" & sWagers(i, WAGER_TYPE) & """"
                
                'Line
                If sWagers(i, WAGER_TYPE) = "H" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)
                If sWagers(i, WAGER_TYPE) = "R" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_LINE)
                If sWagers(i, WAGER_TYPE) = "O" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)
                If sWagers(i, WAGER_TYPE) = "U" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_TOTAL)
                If Len(s) > 0 Then
                    SQL = SQL & "," & s
                Else
                    SQL = SQL & "," & """"""
                End If
                
                'Price
                If sWagers(i, WAGER_TYPE) = "H" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_HOME_PRICE)
                If sWagers(i, WAGER_TYPE) = "R" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_ROAD_PRICE)
                If sWagers(i, WAGER_TYPE) = "O" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_OVER_PRICE)
                If sWagers(i, WAGER_TYPE) = "U" Then s = sGames(Index, (idx * MAX_SCH_ITEMS) + SCH_UNDER_PRICE)
                If Len(s) > 0 Then
                    SQL = SQL & "," & s
                Else
                    SQL = SQL & "," & """"""
                End If
            End If
        Next i
    
        SQL = SQL & ")"
    
        lResult = dbInsert(SQL)
        If lResult <> 0 Then Exit Sub
    End If
    
    'Adjust Balance
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
        End If
        If Form1.oWagerType(BET_TYPE_PARLAY).Value = True Then
            action = action - Val(sParlayAmount)
        End If
        If Form1.oWagerType(BET_TYPE_4TEASER).Value = True Then
            If teams = 2 And Val(AccountSettings(ACT_4PT_TEASER2)) < 0 Then
                action = action - ((Val(sParlayAmount) * (0 - Val(AccountSettings(ACT_4PT_TEASER2)))) / 100)
            Else
                action = action - Val(sParlayAmount)
            End If
        End If
        If Form1.oWagerType(BET_TYPE_6TEASER).Value = True Then
            If teams = 2 And Val(AccountSettings(ACT_6PT_TEASER2)) < 0 Then
                action = action - ((Val(sParlayAmount) * (0 - Val(AccountSettings(ACT_6PT_TEASER2)))) / 100)
            Else
                action = action - Val(sParlayAmount)
            End If
        End If
        If Form1.oWagerType(BET_TYPE_7TEASER).Value = True Then
            If teams = 2 And Val(AccountSettings(ACT_7PT_TEASER2)) < 0 Then
                action = action - ((Val(sParlayAmount) * (0 - Val(AccountSettings(ACT_7PT_TEASER2)))) / 100)
            Else
                action = action - Val(sParlayAmount)
            End If
        End If

        wagersinaction = wagersinaction - action
        AccountSettings(ACT_BALANCE) = Val(AccountSettings(ACT_BALANCE)) + action
        Call dbUpdate("Update Accounts Set InAction = " & wagersinaction & " Where Account = '" & sCurrentAccount & "'")
        Form1.lBalance.Caption = "Cur Balance: " & Val(AccountSettings(ACT_BALANCE))
        Form1.lThisWeek.Caption = "This Week: " & thisweek - wagersinaction
    End If
    
    'Clear the betting ticket form
    sParlayAmount = ""
    For i = 0 To MAX_WAGERS - 1
        sWagers(i, WAGER_SPORT) = ""
        sWagers(i, WAGER_GAME) = ""
        sWagers(i, WAGER_TYPE) = ""
        sWagers(i, WAGER_AMOUNT) = ""
    Next i
    
    DisplayTicket
    
End Sub

Private Sub SchScroll_Change(Index As Integer)
    
    'ReDisplay Schedule When the Scroll Position has Changed
    PaintDisplayArea (Index)

End Sub



Private Sub SchScroll_Scroll(Index As Integer)
   
   'ReDisplay Schedule When the Scroll Position has Changed
    PaintDisplayArea (Index)

End Sub


Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim idx As Long  'Sport Index
    Dim i As Long 'Scroll Value
    
    'Point to the Correct Sport Depending on the Active Tab
    idx = SSTab1.Tab
  
    'Only Process Keys on Sports Tabs

    i = SchScroll(idx).Value
    If KeyCode = vbKeyDown Then i = i + 1
    If KeyCode = vbKeyUp Then i = i - 1
    If KeyCode = vbKeyPageDown Then i = i + (GAMES_PER_TAB - 1)
    If KeyCode = vbKeyPageUp Then i = i - (GAMES_PER_TAB - 1)
    If KeyCode = vbKeyHome Then i = SchScroll(idx).Min
    If KeyCode = vbKeyEnd Then i = SchScroll(idx).Max
        
    'Avoid Over-Shooting the Scroll Bar Positions
    If i > SchScroll(idx).Max Then i = SchScroll(idx).Max
    If i < SchScroll(idx).Min Then i = SchScroll(idx).Min
        
    'Adjust the Scroll Bar Position
    SchScroll(idx).Value = i

End Sub

Private Sub tNumber_Change()

    Dim idx As Long
    Dim gm As Long
    Dim sRoad As String
    Dim sHome As String
    
    For idx = 0 To SPORTS - 1
        For gm = 0 To nGames(idx) - 1
            sRoad = sGames(idx, (gm * MAX_SCH_ITEMS) + SCH_ROTATION)
            sHome = sGames(idx, (gm * MAX_SCH_ITEMS) + SCH_ROTATION) + 1
            If tNumber.Text = Mid(sRoad, 1, Len(tNumber.Text)) Or tNumber.Text = Mid(sHome, 1, Len(tNumber.Text)) Then
                Form1.SSTab1.Tab = idx
                Form1.SchScroll(idx).Value = gm
                Exit Sub
            End If
        Next gm
    Next idx
    
End Sub


Private Sub tNumber_GotFocus()
    tNumber.SelStart = 0
    tNumber.SelLength = Len(tNumber.Text)
End Sub


Private Sub tParlayAmount_Change()
    Dim s As String
    s = tParlayAmount.Text
    
    If Len(s) = 0 Then
        sParlayAmount = ""
    Else
        If Mid(s, 1, 1) = "$" Then s = Mid(s, 2, Len(s) - 1)
        sParlayAmount = Val(s)
    End If
End Sub

Private Sub tParlayAmount_GotFocus()
    tParlayAmount.SelStart = 0
    tParlayAmount.SelLength = Len(tParlayAmount.Text)
End Sub


Private Sub tParlayAmount_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then DisplayTicket
End Sub


Private Sub tParlayAmount_LostFocus()
    DisplayTicket
End Sub


Private Sub tTeam_Change()

    Dim idx As Long
    Dim gm As Long
    
    For idx = 0 To SPORTS - 1
        For gm = 0 To nGames(idx) - 1
             If InStr(1, sGames(idx, (gm * MAX_SCH_ITEMS) + SCH_ROAD_TEAM), tTeam, 1) > 0 Or InStr(1, sGames(idx, (gm * MAX_SCH_ITEMS) + SCH_HOME_TEAM), tTeam, 1) > 0 Then
                Form1.SSTab1.Tab = idx
                Form1.SchScroll(idx).Value = gm
                Exit Sub
            End If
        Next gm
    Next idx
    
End Sub


Private Sub tTeam_GotFocus()
    tTeam.SelStart = 0
    tTeam.SelLength = Len(tTeam.Text)
End Sub


Private Sub tWagerAmount_Change(Index As Integer)
    Dim s As String
    s = tWagerAmount(Index).Text
    
    If Len(s) = 0 Then
        sWagers(Index, WAGER_AMOUNT) = ""
    Else
        If Mid(s, 1, 1) = "$" Then s = Mid(s, 2, Len(s) - 1)
        sWagers(Index, WAGER_AMOUNT) = Val(s)
    End If

End Sub


Private Sub tWagerAmount_GotFocus(Index As Integer)
    tWagerAmount(Index).SelStart = 0
    tWagerAmount(Index).SelLength = Len(tWagerAmount(Index).Text)
End Sub


Private Sub tWagerAmount_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then DisplayTicket
End Sub


Private Sub tWagerAmount_LostFocus(Index As Integer)
    DisplayTicket
End Sub


