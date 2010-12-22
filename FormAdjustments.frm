VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormAdjustments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjustments"
   ClientHeight    =   5505
   ClientLeft      =   1260
   ClientTop       =   1035
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5505
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton pbOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin MSGrid.Grid GridAdjustments 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   8493
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   30
      ScrollBars      =   2
      HighLight       =   0   'False
   End
End
Attribute VB_Name = "FormAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GridAdjustments_GotFocus()
    FormAdjustments.GridAdjustments.Tag = "0"
End Sub

Private Sub GridAdjustments_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not (FormAdjustments.GridAdjustments.Col = 1) Then Exit Sub
    
    'check if a valid key
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal _
    Or (KeyCode >= vbKey0 And KeyCode <= vbKey9) _
    Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) _
    Or KeyCode = vbKeySubtract _
    Or KeyCode = vbKeyBack _
    Or KeyCode = 189 _
    Then

        'Apply the key
        
        If Val(FormAdjustments.GridAdjustments.Tag) = 0 Then
            FormAdjustments.GridAdjustments.Tag = "1"
            FormAdjustments.GridAdjustments.Text = ""
        End If
        
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal Then
            FormAdjustments.GridAdjustments.Text = ""
        End If
    
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
           FormAdjustments.GridAdjustments.Text = FormAdjustments.GridAdjustments.Text & (KeyCode - vbKey0)
        End If
        
        If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
            FormAdjustments.GridAdjustments.Text = FormAdjustments.GridAdjustments.Text & (KeyCode - vbKeyNumpad0)
        End If
    
        If KeyCode = vbKeySubtract Then
            FormAdjustments.GridAdjustments.Text = FormAdjustments.GridAdjustments.Text & "-"
        End If
    
        If KeyCode = vbKeyBack And Len(FormAdjustments.GridAdjustments.Text) > 0 Then
            FormAdjustments.GridAdjustments.Text = Mid(FormAdjustments.GridAdjustments.Text, 1, Len(FormAdjustments.GridAdjustments.Text) - 1)
        End If
        
         If KeyCode = 189 Then
            FormAdjustments.GridAdjustments.Text = FormAdjustments.GridAdjustments.Text & "-"
        End If
        
    End If

End Sub


Private Sub GridAdjustments_LostFocus()
    FormAdjustments.GridAdjustments.Tag = "0"
End Sub

Private Sub GridAdjustments_RowColChange()
    FormAdjustments.GridAdjustments.Tag = "0"
End Sub

Private Sub GridAdjustments_SelChange()
    FormAdjustments.GridAdjustments.Tag = "0"
End Sub


Private Sub pbCancel_Click()
    Unload FormAdjustments
End Sub


Private Sub pbOk_Click()
    ApplyAdjustments
    Unload FormAdjustments
End Sub


Private Sub pbPrint_Click()
    Dim lResult As Long, i As Long, daysThisWeek As Long, x As Long, count As Long, flg As Long
    Dim s() As String
    Dim buf As String, DayOfWeek As String
    Dim thisWeekDate As Date, theDate As Date
    Dim XL As Object
        
    'Calculate The CutOffDate
    DayOfWeek = Format(Now(), "ddd")
    If DayOfWeek = "Tue" Then daysThisWeek = 7
    If DayOfWeek = "Wed" Then daysThisWeek = 1
    If DayOfWeek = "Thu" Then daysThisWeek = 2
    If DayOfWeek = "Fri" Then daysThisWeek = 3
    If DayOfWeek = "Sat" Then daysThisWeek = 4
    If DayOfWeek = "Sun" Then daysThisWeek = 5
    If DayOfWeek = "Mon" Then daysThisWeek = 6
    thisWeekDate = DateAdd("d", 0 - daysThisWeek, Now())

    'Open the Ledger Table
    count = 0
    lResult = dbSelect("Select * From Ledger Where Type = " & BET_TYPE_ADJUSTMENT)
    If lResult <> 0 Then Exit Sub
    Do Until MyTable.EOF
        ReDim Preserve s(3, count + 1)
        s(0, count) = MyTable![Account]
        s(1, count) = MyTable![TimeStamp]
        s(2, count) = MyTable![amount]
        count = count + 1
        MyTable.MoveNext
    Loop
    MyTable.Close

    buf = ""
    For x = 0 To daysThisWeek
        theDate = DateAdd("d", x, thisWeekDate)
        buf = buf & Format(theDate, "dddd mm/dd/yyyy") & vbCr & vbLf
        flg = 0
        For i = 0 To count - 1
            If DateDiff("d", theDate, s(1, i)) = 0 Then
                flg = 1
                buf = buf & vbTab & "Account:" & s(0, i) & "  " & s(2, i) & vbCr & vbLf
            End If
        Next i
        If flg = 0 Then buf = buf & vbTab & "None" & vbCr & vbLf
    Next x

    On Error Resume Next
    Set XL = CreateObject("excel.application")
    If Not XL Is Nothing Then
        XL.Parent.Windows(1).Visible = True
        XL.Workbooks.Open filename:=ADJUSTMENTS_SHEET
    
        XL.Application.Cells.Select
        XL.Application.Selection.ClearContents
        
        Clipboard.SetText buf
        XL.Application.Range("A1").Select
        XL.Application.ActiveSheet.Paste
        XL.Application.Range("A1").Select
        
        XL.Application.ActiveWorkbook.PrintOut Copies:=1
        
        XL.Application.ActiveWorkbook.Save
        XL.Application.Quit
        Set XL = Nothing
    End If

End Sub


