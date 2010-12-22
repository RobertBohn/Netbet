VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormPayments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments/ Collections"
   ClientHeight    =   5505
   ClientLeft      =   1290
   ClientTop       =   1080
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "FormPayments"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5505
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin MSGrid.Grid GridPayments 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
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
Attribute VB_Name = "FormPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GridPayments_GotFocus()
    FormPayments.GridPayments.Tag = "0"
End Sub

Private Sub GridPayments_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (FormPayments.GridPayments.Col = 1) Then Exit Sub
    
    'check if a valid key
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal _
    Or (KeyCode >= vbKey0 And KeyCode <= vbKey9) _
    Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) _
    Or KeyCode = vbKeySubtract _
    Or KeyCode = vbKeyBack _
    Or KeyCode = 189 _
    Then

        'Apply the key
        
        If Val(FormPayments.GridPayments.Tag) = 0 Then
            FormPayments.GridPayments.Tag = "1"
            FormPayments.GridPayments.Text = ""
        End If
        
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal Then
            FormPayments.GridPayments.Text = ""
        End If
    
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
           FormPayments.GridPayments.Text = FormPayments.GridPayments.Text & (KeyCode - vbKey0)
        End If
        
        If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
            FormPayments.GridPayments.Text = FormPayments.GridPayments.Text & (KeyCode - vbKeyNumpad0)
        End If
    
        If KeyCode = vbKeySubtract Then
            FormPayments.GridPayments.Text = FormPayments.GridPayments.Text & "-"
        End If
    
        If KeyCode = vbKeyBack And Len(FormPayments.GridPayments.Text) > 0 Then
            FormPayments.GridPayments.Text = Mid(FormPayments.GridPayments.Text, 1, Len(FormPayments.GridPayments.Text) - 1)
        End If
        
         If KeyCode = 189 Then
            FormPayments.GridPayments.Text = FormPayments.GridPayments.Text & "-"
        End If
        
    End If

End Sub

Private Sub GridPayments_LostFocus()
    FormPayments.GridPayments.Tag = "0"
End Sub


Private Sub GridPayments_RowColChange()
    FormPayments.GridPayments.Tag = "0"
End Sub


Private Sub GridPayments_SelChange()
    FormPayments.GridPayments.Tag = "0"
End Sub


Private Sub pbCancel_Click()
    Unload FormPayments
End Sub


Private Sub pbOk_Click()
    ApplyPayments
    Unload FormPayments
End Sub


