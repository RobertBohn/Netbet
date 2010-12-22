VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form FormDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Detail"
   ClientHeight    =   6285
   ClientLeft      =   180
   ClientTop       =   390
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6285
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox tTeaser7Two 
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "7pt Teasers Ties:"
      Height          =   615
      Left            =   6000
      TabIndex        =   20
      Top             =   4920
      Width           =   1815
      Begin VB.OptionButton rb7pt 
         Caption         =   "Lose"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton rb7pt 
         Caption         =   "Push"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "6pt Teasers Ties:"
      Height          =   615
      Left            =   6000
      TabIndex        =   19
      Top             =   4200
      Width           =   1815
      Begin VB.OptionButton rb6pt 
         Caption         =   "Lose"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton rb6pt 
         Caption         =   "Push"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "4pt Teasers Ties:"
      Height          =   615
      Left            =   6000
      TabIndex        =   18
      Top             =   3480
      Width           =   1815
      Begin VB.OptionButton rb4ptteaser 
         Caption         =   "Push"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton rb4ptteaser 
         Caption         =   "Lose"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fPitchers 
      Caption         =   "Pitchers Must Start:"
      Height          =   615
      Left            =   6000
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
      Begin VB.OptionButton rbPitchers 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton rbPitchers 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox tTeaser6Two 
      Height          =   285
      Left            =   6000
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox tStraightBet 
      Height          =   285
      Left            =   6960
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox tLimit 
      Height          =   285
      Left            =   6960
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.ListBox DayList 
      Height          =   840
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton pbExit 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
   Begin MSOutl.Outline Outline1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   10610
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin VB.Label lBackground 
      Caption         =   "2-Team 7-Point Teasers:"
      Height          =   255
      Index           =   4
      Left            =   6000
      TabIndex        =   23
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lBackground 
      Caption         =   "2-Team 6-Point Teasers:"
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   16
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lBackground 
      Caption         =   "Straight Bet:"
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lBackground 
      Caption         =   "Wager Limit:"
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lBackground 
      Caption         =   "Account:"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FormDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit

Public theAccount As New Account



Private Sub Form_Unload(Cancel As Integer)

    'Update Limit If Changed
    If Val(FormDetail.tLimit) <> Val(AccountSettings(ACT_LIMIT)) Then
        AccountSettings(ACT_LIMIT) = Val(FormDetail.tLimit)
        Call dbUpdate("Update Accounts Set Limit = " & Val(FormDetail.tLimit) & " Where Account = '" & sCurrentAccount & "'")
        Form1.lLimit.Caption = "Limit:  $" & AccountSettings(ACT_LIMIT)
    End If

    'Update StraightBet Odds If Changed
    If Val(FormDetail.tStraightBet) <> Val(AccountSettings(ACT_STRAIGHT_BET)) Then
        AccountSettings(ACT_STRAIGHT_BET) = Val(FormDetail.tStraightBet)
        Call dbUpdate("Update Accounts Set StraightBet = " & Val(FormDetail.tStraightBet) & " Where Account = '" & sCurrentAccount & "'")
    End If

    'Update 6pt Teaser Odds If Changed
    If Val(FormDetail.tTeaser6Two) <> Val(AccountSettings(ACT_6PT_TEASER2)) Then
        AccountSettings(ACT_6PT_TEASER2) = Val(FormDetail.tTeaser6Two)
        Call dbUpdate("Update Accounts Set Teaser6Two = " & Val(FormDetail.tTeaser6Two) & " Where Account = '" & sCurrentAccount & "'")
    End If

    'Update 7pt Teaser Odds If Changed
    If Val(FormDetail.tTeaser7Two) <> Val(AccountSettings(ACT_7PT_TEASER2)) Then
        AccountSettings(ACT_7PT_TEASER2) = Val(FormDetail.tTeaser7Two)
        Call dbUpdate("Update Accounts Set Teaser7Two = " & Val(FormDetail.tTeaser7Two) & " Where Account = '" & sCurrentAccount & "'")
    End If

    'Update Pitchers
    If FormDetail.rbPitchers(0).Value = True Then
        AccountSettings(ACT_PITCHERS) = "1"
        Call dbUpdate("Update Accounts Set Pitchers = 1 Where Account = '" & sCurrentAccount & "'")
    Else
        AccountSettings(ACT_PITCHERS) = "0"
        Call dbUpdate("Update Accounts Set Pitchers = 0 Where Account = '" & sCurrentAccount & "'")
    End If

    'Update 4pt Teaser Ties
    If FormDetail.rb4ptteaser(0).Value = True Then
        AccountSettings(ACT_4PT_TEASER) = "1"
        Call dbUpdate("Update Accounts Set Teaser4Ties = 1 Where Account = '" & sCurrentAccount & "'")
    Else
        AccountSettings(ACT_4PT_TEASER) = "0"
        Call dbUpdate("Update Accounts Set Teaser4Ties = 0 Where Account = '" & sCurrentAccount & "'")
    End If

    'Update 6pt Teaser Ties
    If FormDetail.rb6pt(0).Value = True Then
        AccountSettings(ACT_6PT_TEASER) = "1"
        Call dbUpdate("Update Accounts Set Teaser6Ties = 1 Where Account = '" & sCurrentAccount & "'")
    Else
        AccountSettings(ACT_6PT_TEASER) = "0"
        Call dbUpdate("Update Accounts Set Teaser6Ties = 0 Where Account = '" & sCurrentAccount & "'")
    End If

    'Update 7pt Teaser Ties
    If FormDetail.rb7pt(0).Value = True Then
        AccountSettings(ACT_7PT_TEASER) = "1"
        Call dbUpdate("Update Accounts Set Teaser7Ties = 1 Where Account = '" & sCurrentAccount & "'")
    Else
        AccountSettings(ACT_7PT_TEASER) = "0"
        Call dbUpdate("Update Accounts Set Teaser7Ties = 0 Where Account = '" & sCurrentAccount & "'")
    End If

End Sub


Private Sub Outline1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim result As Long
    
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal Then
        FormDetail.MousePointer = vbHourglass
        result = FormDetail.theAccount.DeleteItem
        If result < 0 Then
            Call FormDetail.theAccount.Populate(sCurrentAccount)
            Call FormDetail.theAccount.Display
            
            If DEPOSIT_SYSTEM = True Then
                If dbSelect("Select * From Accounts Where Account = '" & sCurrentAccount & "'") = 0 Then
                    If Not MyTable.EOF Then
                        AccountSettings(ACT_BALANCE) = MyTable![balance]
                        MyTable.Close
                        AccountSettings(ACT_IN_ACTION) = FormDetail.theAccount.InAction()
                        AccountSettings(ACT_BALANCE) = Val(AccountSettings(ACT_BALANCE)) - AccountSettings(ACT_IN_ACTION)
                    Else
                        MyTable.Close
                    End If
                End If
            End If
            
        End If
        FormDetail.MousePointer = vbDefault
    End If
    
End Sub


Private Sub Outline1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        Call FormDetail.theAccount.DisplayItem
    End If
End Sub


Private Sub pbExit_Click()
    Unload FormDetail
End Sub

Private Sub pbPrint_Click()
    Dim nRow, nCol As Long
    Dim MyXL As Object                  'Variable to hold reference to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   'Flag for final release.

    FormDetail.MousePointer = vbHourglass

    ' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    ' GetObject function called without the first argument returns a
    ' reference to an instance of the application. If the application isn't
    ' running, an  error occurs. Note the comma used as the first argument placeholder.
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

    ' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject(DETAIL_SHEET)
    'MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True


    MyXL.Application.Cells.Select
    MyXL.Application.Selection.ClearContents
    MyXL.Application.Range("A1").Select
  
    MyXL.Application.Cells(1, 1).Value = "Account: " & sCurrentAccount
    For nRow = 0 To FormDetail.Outline1.ListCount - 1
        FormDetail.MousePointer = vbHourglass
        MyXL.Application.Cells(nRow + 3, FormDetail.Outline1.Indent(nRow)).Value = FormDetail.Outline1.List(nRow)
    Next nRow
    
    MyXL.Application.ActiveWorkbook.PrintOut Copies:=1
    MyXL.Application.ActiveWorkbook.Save
    If ExcelWasNotRunning = True Then MyXL.Application.Quit
    Set MyXL = Nothing  ' Release reference to the application and spreadsheet.

    FormDetail.MousePointer = vbDefault

End Sub

Private Sub tLimit_GotFocus()
    tLimit.SelStart = 0
    tLimit.SelLength = Len(tLimit.Text)
End Sub


Private Sub tStraightBet_GotFocus()
    tStraightBet.SelStart = 0
    tStraightBet.SelLength = Len(tStraightBet.Text)
End Sub


Private Sub tTeaser6Two_GotFocus()
    tTeaser6Two.SelStart = 0
    tTeaser6Two.SelLength = Len(tTeaser6Two.Text)
End Sub


Private Sub tTeaser7Two_GotFocus()
    tTeaser7Two.SelStart = 0
    tTeaser7Two.SelLength = Len(tTeaser7Two.Text)
End Sub


