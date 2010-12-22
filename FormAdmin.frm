VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administration - For Entertainment Purposes"
   ClientHeight    =   4830
   ClientLeft      =   525
   ClientTop       =   1500
   ClientWidth     =   5580
   Icon            =   "FormAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4830
   ScaleWidth      =   5580
   Begin VB.Frame Frame3 
      Caption         =   "Reports"
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
      Begin VB.CommandButton pbPaymentsReport 
         Caption         =   "Payments Report"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton pbTotals 
         Caption         =   "Totals Report"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton pbDaily 
         Caption         =   "Daily Report"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Accounts"
      Height          =   1935
      Left            =   2880
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton pbAdjustments 
         Caption         =   " Adjustments"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton pbPC 
         Caption         =   "Payments && Collections"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton pbAccountDetail 
         Caption         =   "Account Detail"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Games"
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton pbImportScores 
         Caption         =   "Import Scores"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton pbScores 
         Caption         =   "Final Scores"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton pbLines 
         Caption         =   "Line Changes"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton pbSchedule 
         Caption         =   "Import Schedule"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton pbExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lLimit 
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit



Private Sub Form_Load()
   
    'Initialize
    DEPOSIT_SYSTEM = GetSetting("Ticket Writer", "Settings", "DepositSystem", 0)
    ALLOW_DATE_TO_BE_SET_BACK = GetSetting("Ticket Writer", "Settings", "SetBackDate", 0)
    DELETE_ALL_BUTTON = GetSetting("Ticket Writer", "Settings", "DeleteAllButton", 1)
    ALL_TIES_PUSH = GetSetting("Ticket Writer", "Settings", "AllTiesPush", 0)
    NBA_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "NBATodayOnly", 0)
    NHL_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "NHLTodayOnly", 0)
    MLB_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "MLBTodayOnly", 0)
    NFL_TODAY_ONLY = GetSetting("Ticket Writer", "Settings", "NFLTodayOnly", 0)
    STOP_WORKING_DATE = GetSetting("Ticket Writer", "Settings", "StopWorkingDate", "1/1/95")
    MY_PASSWORD = GetSetting("Ticket Writer", "Settings", "Password", "")
    WAGER_AFTER_KICKOFF = GetSetting("Ticket Writer", "Settings", "WagersAfterKickoff", 0)
    MY_PASSWORD = GetSetting("Ticket Writer", "Settings", "Password", 0)
    IMPORT_FILE_PATH = GetSetting("Ticket Writer", "Settings", "ImportFilePath", App.Path)
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
      
    'Read Default System Settings
    ReadDefaults
      
    'Read in Schedule
    ReadSchedule
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    'Close the Database
    On Error Resume Next
    dbClose
    End

End Sub


Private Sub pbAccountDetail_Click()

    Dim lResult As Long
    lResult = dbSelect("Select * from Accounts")
    If lResult <> 0 Then Exit Sub

    'Read Active Accounts into the cbAccount ComboBox
    FormSelectAccount.cbAccount.Clear
    Do Until MyTable.EOF
        If MyTable![Status] = "A" Then FormSelectAccount.cbAccount.AddItem MyTable![Account]
        MyTable.MoveNext
    Loop

    'Close the Accounts Table
    MyTable.Close
            
    If FormSelectAccount.cbAccount.ListCount > 0 Then
        FormSelectAccount.cbAccount.ListIndex = 0
    End If
    
    FormSelectAccount.Show 1
End Sub

Private Sub pbAdjustments_Click()
    DisplayAdjustments
    FormAdjustments.Show 1
End Sub

Private Sub pbDaily_Click()
    DisplayDaily
    FormDaily.Show 1
End Sub

Private Sub pbExit_Click()
    Unload Form1
End Sub


Private Sub pbImportScores_Click()
    Dim nRow As Long
    Dim MyXL As Object                  'Variable to hold reference to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   'Flag for final release.
    Dim xSport As String, xDate As String, xRoad As String, xHome As String
    Dim sSQL As String, xRoadScore As String, xHomeScore As String
    Dim xRotation As String
    Dim lGameNbr As Long, lResult As Long
    Dim iUpdated As Long, iProcessed As Long
    Dim sPath As String
    Dim sOriginalRoad As String, sOriginalHome As String
   
    FormOpenExcel.CommonDialog1.filename = ""
    FormOpenExcel.CommonDialog1.InitDir = IMPORT_FILE_PATH
    FormOpenExcel.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNFileMustExist
    FormOpenExcel.CommonDialog1.ShowOpen
    sPath = FormOpenExcel.CommonDialog1.filename
    Unload FormOpenExcel
    If Len(sPath) = 0 Then
        Exit Sub
    End If
    Form1.MousePointer = vbHourglass

    IMPORT_FILE_PATH = sPath
    Do
        If Len(IMPORT_FILE_PATH) = 0 Then Exit Do
        If Mid(IMPORT_FILE_PATH, Len(IMPORT_FILE_PATH), 1) = "\" Then
            IMPORT_FILE_PATH = Mid(IMPORT_FILE_PATH, 1, Len(IMPORT_FILE_PATH) - 1)
            Call SaveSetting("Ticket Writer", "Settings", "ImportFilePath", IMPORT_FILE_PATH)
            Exit Do
        Else
            IMPORT_FILE_PATH = Mid(IMPORT_FILE_PATH, 1, Len(IMPORT_FILE_PATH) - 1)
        End If
    Loop

    ' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    ' GetObject function called without the first argument returns a
    ' reference to an instance of the application. If the application isn't
    ' running, an  error occurs. Note the comma used as the first argument placeholder.
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

    ' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject(sPath)
    'MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True

    iUpdated = 0
    iProcessed = 0
    nRow = 1
    Do
        'Get Data From Spreadsheet
        Form1.MousePointer = vbHourglass
        xSport = MyXL.Application.Cells(nRow, 1).Value
        xDate = MyXL.Application.Cells(nRow, 2).Value
        xRotation = MyXL.Application.Cells(nRow, 3).Value
        xRoad = MyXL.Application.Cells(nRow, 4).Value
        xHome = MyXL.Application.Cells(nRow + 1, 4).Value
        xRoadScore = MyXL.Application.Cells(nRow, 9).Value
        xHomeScore = MyXL.Application.Cells(nRow + 1, 9).Value
        Form1.MousePointer = vbHourglass
  
        'Check for End of Spreadsheet
        If Len(xRoad) = 0 And Len(xHome) = 0 Then Exit Do
        iProcessed = iProcessed + 1
  
        'Check for Valid Sport
        If xSport = "CFB" Then xSport = "NFL"
        If xSport = "CBB" Then xSport = "NBA"
        If xSport <> "NBA" And xSport <> "NFL" And xSport <> "NHL" And xSport <> "MLB" Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD SPORT"
            GoTo NextGame
        End If
  
        'Check Date
        If Len(xDate) = 0 Or IsDate(xDate) = False Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD DATE"
            GoTo NextGame
        End If
        
        'Check Scores
        If Len(xHomeScore) = 0 Or Len(xRoadScore) = 0 Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "MISSING SCORE"
            iUpdated = iUpdated + 1
            GoTo NextGame
        End If
 
        'Find the Game
        lGameNbr = -1
        
        If xSport = "MLB" Then
            sSQL = "Select * From Schedule Where Sport = '" & xSport & "' And RotationNumber = " & xRotation
        Else
            sSQL = "Select * From Schedule Where Sport = '" & xSport & "' and RoadTeam = '" & xRoad & "' and HomeTeam = '" & xHome & "'"
        End If
        lResult = dbSelect(sSQL)
        If lResult <> 0 Then
            Exit Sub
        End If
    
        Do Until MyTable.EOF
            If DateDiff("d", xDate, MyTable![GameDate]) = 0 Then
                If xSport = "MLB" Then
                    sOriginalRoad = MyTable![RoadTeam]
                    sOriginalHome = MyTable![HomeTeam]
                    If ExtractTeamName(sOriginalRoad) = ExtractTeamName(xRoad) And ExtractTeamName(sOriginalHome) = ExtractTeamName(xHome) Then
                        lGameNbr = MyTable![GameNumber]
                        Exit Do
                    End If
                Else
                    lGameNbr = MyTable![GameNumber]
                    Exit Do
                End If
            End If
            MyTable.MoveNext
        Loop
        MyTable.Close

        If lGameNbr > 0 Then
            'Update
            iUpdated = iUpdated + 1
          
            sSQL = "Update Schedule Set RoadScore = " & xRoadScore & " Where GameNumber = " & lGameNbr
            Call dbUpdate(sSQL)
          
            sSQL = "Update Schedule Set HomeScore = " & xHomeScore & " Where GameNumber = " & lGameNbr
            Call dbUpdate(sSQL)
            
            MyXL.Application.Cells(nRow + 1, 11).Value = "UPDATED"
        End If
        
NextGame:
        nRow = nRow + 2
    Loop
    
    MyXL.Application.ActiveWorkbook.Save
    If ExcelWasNotRunning = True Then MyXL.Application.Quit
    Set MyXL = Nothing  ' Release reference to the application and spreadsheet.
    
    Form1.MousePointer = vbHourglass
    ReCalcBalances
    Form1.MousePointer = vbDefault
    
    If iProcessed = 0 Then
        sSQL = "No games were found in " & sPath
    Else
        If iProcessed = iUpdated Then
            sSQL = iProcessed & " games processed."
        Else
            sSQL = (iProcessed - iUpdated) & " games had errors. Examine column K in " & sPath
        End If
    End If
    MsgBox sSQL


End Sub

Private Sub pbLines_Click()
    DisplayLineChanges
    FormLines.Show 1
End Sub


Private Sub pbPaymentsReport_Click()
    DisplayPaymentReport
    FormPaymentReport.Show 1
End Sub

Private Sub pbPC_Click()
    DisplayPayments
    FormPayments.Show 1
End Sub

Private Sub pbSchedule_Click()
    
    Dim nRow As Long
    Dim MyXL As Object                  'Variable to hold reference to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   'Flag for final release.
    Dim xSport As String, xDate As String, xTime As String, xRoad As String, xHome As String, xRot1 As String, xRot2 As String
    Dim xRoadLine As String, xHomeLine As String, xRoadPrice As String, xHomePrice As String, xTotal As String, xOverPrice As String, xUnderPrice As String
'    Dim xRoadScore, xHomeScore As String
    Dim sSQL As String, sOriginalRoad As String, sOriginalHome As String
    Dim lGameNbr As Long, lResult As Long
    Dim iSkipped As Long, iAdded As Long, iUpdated As Long, iProcessed As Long
    Dim sPath As String
   
    FormOpenExcel.CommonDialog1.filename = ""
    FormOpenExcel.CommonDialog1.InitDir = IMPORT_FILE_PATH
    FormOpenExcel.CommonDialog1.DefaultExt = "xls"
    FormOpenExcel.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNFileMustExist
    FormOpenExcel.CommonDialog1.ShowOpen
    sPath = FormOpenExcel.CommonDialog1.filename
    Unload FormOpenExcel
    
    If Len(sPath) = 0 Then
        Exit Sub
    End If
    
    IMPORT_FILE_PATH = sPath
    Do
        If Len(IMPORT_FILE_PATH) = 0 Then Exit Do
        If Mid(IMPORT_FILE_PATH, Len(IMPORT_FILE_PATH), 1) = "\" Then
            IMPORT_FILE_PATH = Mid(IMPORT_FILE_PATH, 1, Len(IMPORT_FILE_PATH) - 1)
            Call SaveSetting("Ticket Writer", "Settings", "ImportFilePath", IMPORT_FILE_PATH)
            Exit Do
        Else
            IMPORT_FILE_PATH = Mid(IMPORT_FILE_PATH, 1, Len(IMPORT_FILE_PATH) - 1)
        End If
    Loop
    
    Form1.MousePointer = vbHourglass

    ' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    ' GetObject function called without the first argument returns a
    ' reference to an instance of the application. If the application isn't
    ' running, an  error occurs. Note the comma used as the first argument placeholder.
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

    ' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject(sPath)
    'MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True

    iSkipped = 0
    iAdded = 0
    iUpdated = 0
    iProcessed = 0
    nRow = 1
    Do
        'Get Data From Spreadsheet
        Form1.MousePointer = vbHourglass
        xSport = MyXL.Application.Cells(nRow, 1).Value
        xDate = MyXL.Application.Cells(nRow, 2).Value
        xTime = MyXL.Application.Cells(nRow + 1, 2).Value
        xRoad = MyXL.Application.Cells(nRow, 4).Value
        xHome = MyXL.Application.Cells(nRow + 1, 4).Value
        xRot1 = MyXL.Application.Cells(nRow, 3).Value
        xRot2 = MyXL.Application.Cells(nRow + 1, 3).Value
        xRoadLine = MyXL.Application.Cells(nRow, 5).Value
        xHomeLine = MyXL.Application.Cells(nRow + 1, 5).Value
        xRoadPrice = MyXL.Application.Cells(nRow, 6).Value
        xHomePrice = MyXL.Application.Cells(nRow + 1, 6).Value
        xTotal = MyXL.Application.Cells(nRow + 1, 7).Value
        xOverPrice = MyXL.Application.Cells(nRow, 8).Value
        xUnderPrice = MyXL.Application.Cells(nRow + 1, 8).Value
'        xRoadScore = MyXL.Application.Cells(nRow, 9).Value
'        xHomeScore = MyXL.Application.Cells(nRow + 1, 9).Value
        Form1.MousePointer = vbHourglass
  
        'Check for End of Spreadsheet
        If Len(xRoad) = 0 And Len(xHome) = 0 And Len(xSport) = 0 And Len(xDate) = 0 And Len(xTime) = 0 Then Exit Do
        
        iProcessed = iProcessed + 1
  
        If Len(xRoad) = 0 And Len(xHome) = 0 Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "SKIPPED"
            iSkipped = iSkipped + 1
            GoTo NextGame
        End If
  
        'Check for Valid Sport
        If xSport = "CFB" Then xSport = "NFL"
        If xSport = "CBB" Then xSport = "NBA"
        If xSport <> "NBA" And xSport <> "NFL" And xSport <> "NHL" And xSport <> "MLB" Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD SPORT"
            GoTo NextGame
        End If
        
        'Check Date
        If Len(xDate) = 0 Or IsDate(xDate) = False Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD DATE"
            GoTo NextGame
        End If
  
        'Check Time
        If Len(xTime) = 0 Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD TIME"
            GoTo NextGame
        End If
        xDate = Format(xDate, "m/d/yyyy") & " " & Format(xTime, "h:mm AMPM")
        If IsDate(xDate) = False Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD TIME"
            GoTo NextGame
        End If
   
        'Check Rotation Numbers
        If Val(xRot1) < 1 Or (Val(xRot1) Mod 2) = 0 Or Val(xRot1) <> (Val(xRot2) - 1) Then
            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD ROTATION NUMBER"
            GoTo NextGame
        End If
  
        'Check NFL & NBA Lines
        If xSport = "NFL" Or xSport = "NBA" Then
            If Len(xRoadLine) > 0 And Len(xHomeLine) = 0 Then
                xHomeLine = 0 - Val(xRoadLine)
                xRoadLine = ""
            End If
        End If
  
        'Check Hockey Lines
        If xSport = "NHL" Then
            If (Len(xRoadLine) > 0 And Len(xHomeLine) = 0) _
            Or (Len(xRoadLine) = 0 And Len(xHomeLine) > 0) Then
                MyXL.Application.Cells(nRow + 1, 11).Value = "BAD LINE"
                GoTo NextGame
            End If
            If Len(xRoadLine) > 0 Then
                If Len(xRoadPrice) = 0 Then xRoadPrice = "-110"
                If Len(xHomePrice) = 0 Then xHomePrice = "-110"
            Else
                xRoadPrice = ""
                xHomePrice = ""
            End If
        End If
        
        'Check Totals Money Line
        If xSport = "NHL" Or xSport = "MLB" Then
            If Len(xTotal) > 0 Then
                If Len(xOverPrice) = 0 Then xOverPrice = "-110"
                If Len(xUnderPrice) = 0 Then xUnderPrice = "-110"
            End If
        Else
            xOverPrice = ""
            xUnderPrice = ""
        End If
        
        'Check Scores
'        If (Len(xRoadScore) > 0 And Len(xHomeScore) = 0) _
'        Or (Len(xRoadScore) = 0 And Len(xHomeScore) > 0) Then
'            MyXL.Application.Cells(nRow + 1, 11).Value = "BAD SCORE"
'            GoTo NextGame
'        End If
         
        'Check MLB Line
        If xSport = "MLB" Then
            If (Len(xRoadLine) > 0 And Len(xHomeLine) = 0) _
            Or (Len(xRoadLine) = 0 And Len(xHomeLine) > 0) Then
                MyXL.Application.Cells(nRow + 1, 11).Value = "BAD LINE"
                GoTo NextGame
            End If
            If Len(xRoadLine) > 0 Then
                xRoadPrice = xRoadLine
                xHomePrice = xHomeLine
                xRoadLine = ""
                xHomeLine = ""
            End If
        End If

        'Find the Game
        lGameNbr = -1
        sSQL = "Select * From Schedule Where Sport = '" & xSport & "' and RotationNumber = " & xRot1
        lResult = dbSelect(sSQL)
        If lResult <> 0 Then
            Exit Sub
        End If
    
        Do Until MyTable.EOF
            If DateDiff("d", xDate, MyTable![GameDate]) = 0 Then
                lGameNbr = MyTable![GameNumber]
                sOriginalRoad = MyTable![RoadTeam]
                sOriginalHome = MyTable![HomeTeam]
                Exit Do
            End If
            MyTable.MoveNext
        Loop
        MyTable.Close

        If lGameNbr > 0 Then
            'Update
            If xSport = "MLB" Then
                If ExtractTeamName(sOriginalRoad) <> ExtractTeamName(xRoad) Or ExtractTeamName(sOriginalHome) <> ExtractTeamName(xHome) Then
                    MyXL.Application.Cells(nRow + 1, 11).Value = "BAD TEAM NAME OR ROTATION NUMBER"
                    GoTo NextGame
                End If
            Else
                If sOriginalRoad <> xRoad Or sOriginalHome <> xHome Then
                    MyXL.Application.Cells(nRow + 1, 11).Value = "BAD TEAM NAME OR ROTATION NUMBER"
                    GoTo NextGame
                End If
            End If
            iUpdated = iUpdated + 1
            
            sSQL = "Update Schedule Set GameDate = '" & xDate & "' Where GameNumber = " & lGameNbr
            Call dbUpdate(sSQL)
            
            If Len(xHome) > Len(sOriginalHome) Then
                sSQL = "Update Schedule Set HomeTeam = '" & NoQuotes(xHome) & "' Where GameNumber = " & lGameNbr
                Call dbUpdate(sSQL)
            End If
            
            If Len(xRoad) > Len(sOriginalRoad) Then
                sSQL = "Update Schedule Set RoadTeam = '" & NoQuotes(xRoad) & "' Where GameNumber = " & lGameNbr
                Call dbUpdate(sSQL)
            End If
            
            If Len(xHomeLine) > 0 Then
                sSQL = "Update Schedule Set Line = " & xHomeLine & " Where GameNumber = " & lGameNbr
            Else
                sSQL = "Update Schedule Set Line = '' Where GameNumber = " & lGameNbr
            End If
            Call dbUpdate(sSQL)
            
            If Len(xTotal) > 0 Then
                sSQL = "Update Schedule Set Total = " & xTotal & " Where GameNumber = " & lGameNbr
            Else
                sSQL = "Update Schedule Set Total = '' Where GameNumber = " & lGameNbr
            End If
            Call dbUpdate(sSQL)
            
            If Len(xRoadPrice) > 0 Then
                sSQL = "Update Schedule Set RoadPrice = " & xRoadPrice & " Where GameNumber = " & lGameNbr
            Else
                sSQL = "Update Schedule Set RoadPrice = '' Where GameNumber = " & lGameNbr
            End If
            Call dbUpdate(sSQL)
            
            If Len(xHomePrice) > 0 Then
                sSQL = "Update Schedule Set HomePrice = " & xHomePrice & " Where GameNumber = " & lGameNbr
            Else
                sSQL = "Update Schedule Set HomePrice = '' Where GameNumber = " & lGameNbr
            End If
            Call dbUpdate(sSQL)
            
            If Len(xOverPrice) > 0 Then
                sSQL = "Update Schedule Set OverPrice = " & xOverPrice & " Where GameNumber = " & lGameNbr
            Else
                sSQL = "Update Schedule Set OverPrice = '' Where GameNumber = " & lGameNbr
            End If
            Call dbUpdate(sSQL)
           
            If Len(xUnderPrice) > 0 Then
                sSQL = "Update Schedule Set UnderPrice = " & xUnderPrice & " Where GameNumber = " & lGameNbr
            Else
                sSQL = "Update Schedule Set UnderPrice = '' Where GameNumber = " & lGameNbr
            End If
            Call dbUpdate(sSQL)
           
 '           If Len(xRoadScore) > 0 Then
 '               sSQL = "Update Schedule Set RoadScore = " & xRoadScore & " Where GameNumber = " & lGameNbr
 '           Else
 '               sSQL = "Update Schedule Set RoadScore = '' Where GameNumber = " & lGameNbr
 '           End If
 '           Call dbUpdate(sSQL)
          
 '           If Len(xHomeScore) > 0 Then
 '               sSQL = "Update Schedule Set HomeScore = " & xHomeScore & " Where GameNumber = " & lGameNbr
 '           Else
 '               sSQL = "Update Schedule Set HomeScore = '' Where GameNumber = " & lGameNbr
 '           End If
 '           Call dbUpdate(sSQL)
            
            MyXL.Application.Cells(nRow + 1, 11).Value = "UPDATED"
        Else
            'Add
            iAdded = iAdded + 1
            sSQL = "Insert Into Schedule(GameDate,RotationNumber,Sport,RoadTeam,HomeTeam,Line,Total,RoadPrice,HomePrice,OverPrice,UnderPrice) Values("
            sSQL = sSQL & "'" & xDate & "',"
            sSQL = sSQL & Val(xRot1) & ","
            sSQL = sSQL & "'" & xSport & "',"
            sSQL = sSQL & "'" & xRoad & "',"
            sSQL = sSQL & "'" & xHome & "',"
             
            If Len(xHomeLine) > 0 Then
                sSQL = sSQL & xHomeLine & ","
            Else
                sSQL = sSQL & "'',"
            End If
             
            If Len(xTotal) > 0 Then
                sSQL = sSQL & xTotal & ","
            Else
                sSQL = sSQL & "'',"
            End If
            
            If Len(xRoadPrice) > 0 Then
                sSQL = sSQL & xRoadPrice & ","
            Else
                sSQL = sSQL & "'',"
            End If
            
            If Len(xHomePrice) > 0 Then
                sSQL = sSQL & xHomePrice & ","
            Else
                sSQL = sSQL & "'',"
            End If
            
            If Len(xOverPrice) > 0 Then
                sSQL = sSQL & xOverPrice & ","
            Else
                sSQL = sSQL & "'',"
            End If
            
            If Len(xUnderPrice) > 0 Then
                sSQL = sSQL & xUnderPrice
            Else
                sSQL = sSQL & "''"
            End If
            
 '           If Len(xHomeScore) > 0 Then
 '               sSQL = sSQL & xHomeScore & ","
 '           Else
 '               sSQL = sSQL & "'',"
 '           End If
           
 '           If Len(xRoadScore) > 0 Then
 '               sSQL = sSQL & xRoadScore
 '           Else
 '               sSQL = sSQL & "''"
 '           End If
            
            sSQL = sSQL & ")"
            Call dbInsert(sSQL)
            
            MyXL.Application.Cells(nRow + 1, 11).Value = "ADDED"
        End If
  
NextGame:
        nRow = nRow + 2
    Loop
    
    MyXL.Application.ActiveWorkbook.Save
    If ExcelWasNotRunning = True Then MyXL.Application.Quit
    Set MyXL = Nothing  ' Release reference to the application and spreadsheet.
    
'    Form1.MousePointer = vbHourglass
'    ReCalcBalances
    ReadSchedule
    Form1.MousePointer = vbDefault
    
    If iProcessed = 0 Then
        sSQL = "No games were found in " & sPath
    Else
        If iProcessed = (iAdded + iUpdated + iSkipped) Then
            sSQL = iProcessed & " games processed."
        Else
            sSQL = (iProcessed - iAdded - iUpdated - iSkipped) & " games had errors. Examine column K in " & sPath
        End If
    End If
    MsgBox sSQL

End Sub
Private Sub pbScores_Click()
    FormDate.lDate = Format(Now, "m/d/yyyy ddd")
    FormDate.sbDate.Max = 1000
    FormDate.sbDate.Value = 500
    FormDate.Show 1
End Sub


Private Sub pbTotals_Click()
    DisplayTotals
    FormTotals.Show 1
End Sub


