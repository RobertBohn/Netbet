VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormScores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Final Scores"
   ClientHeight    =   5490
   ClientLeft      =   375
   ClientTop       =   660
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin MSGrid.Grid GridScore 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
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
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
End
Attribute VB_Name = "FormScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GridScore_GotFocus()
    FormScores.GridScore.Tag = "0"
End Sub

Private Sub GridScore_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not FormScores.GridScore.Col = 3 Then Exit Sub

    'check if a valid key
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal _
    Or (KeyCode >= vbKey0 And KeyCode <= vbKey9) _
    Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) _
    Or KeyCode = vbKeyBack Or KeyCode = vbKeyP _
    Then
        If Val(FormScores.GridScore.Tag) = 0 Then
            FormScores.GridScore.Tag = "1"
            FormScores.GridScore.Text = ""
        End If
        
        If KeyCode = vbKeyDelete Then
            FormScores.GridScore.Text = ""
        End If
    
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
           FormScores.GridScore.Text = FormScores.GridScore.Text & (KeyCode - vbKey0)
        End If
        
        If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
            FormScores.GridScore.Text = FormScores.GridScore.Text & (KeyCode - vbKeyNumpad0)
        End If
    
        If KeyCode = vbKeyBack And Len(FormScores.GridScore.Text) > 0 Then
            FormScores.GridScore.Text = Mid(FormScores.GridScore.Text, 1, Len(FormScores.GridScore.Text) - 1)
        End If
        
        If KeyCode = vbKeyP Then
            FormScores.GridScore.Text = "NoGame"
        End If
        
    End If

End Sub


Private Sub GridScore_LostFocus()
    FormScores.GridScore.Tag = "0"
End Sub

Private Sub GridScore_RowColChange()
    FormScores.GridScore.Tag = "0"
End Sub

Private Sub GridScore_SelChange()
    FormScores.GridScore.Tag = "0"
End Sub


Private Sub pbCancel_Click()
    Unload FormScores
End Sub




Private Sub pbOk_Click()
    
    Dim i As Long
    Dim SQL As String
    Dim lResult As Long
    Dim Cancelled As String
    
    If FormScores.GridScore.Rows < 2 Then
        Unload FormScores
        Exit Sub
    End If
    
    FormScores.MousePointer = vbHourglass
    
    For i = 0 To ((FormScores.GridScore.Rows - 1) / 2) - 1
        Cancelled = " "
        
        FormScores.GridScore.Row = (i * 2) + 1
        FormScores.GridScore.Col = 3
        If FormScores.GridScore.Text = "NoGame" Then Cancelled = "C"
        If Len(FormScores.GridScore.Text) > 0 Then
            SQL = "Update Schedule Set RoadScore = " & Val(FormScores.GridScore.Text) & " Where GameNumber = " & nSchNbr(i)
        Else
            SQL = "Update Schedule Set RoadScore = '' Where GameNumber = " & nSchNbr(i)
        End If
        lResult = dbUpdate(SQL)
        If lResult <> 0 Then Exit Sub
           
        FormScores.GridScore.Row = (i * 2) + 2
        If FormScores.GridScore.Text = "NoGame" Then Cancelled = "C"
        If Len(FormScores.GridScore.Text) > 0 Then
            SQL = "Update Schedule Set HomeScore = " & Val(FormScores.GridScore.Text) & " Where GameNumber = " & nSchNbr(i)
        Else
            SQL = "Update Schedule Set HomeScore = '' Where GameNumber = " & nSchNbr(i)
        End If
        lResult = dbUpdate(SQL)
        If lResult <> 0 Then Exit Sub
        
        SQL = "Update Schedule Set Cancelled = '" & Cancelled & "' Where GameNumber = " & nSchNbr(i)
        lResult = dbUpdate(SQL)
        If lResult <> 0 Then Exit Sub
 
    Next i
    ReCalcBalances
    FormScores.MousePointer = vbDefault
    Unload FormScores
End Sub


