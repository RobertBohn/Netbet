VERSION 5.00
Begin VB.Form FormEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2865
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "FormEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit

Public GameIndex As Long
Public GameOffset As Long
Public BetType As Long
Private Sub Form_Load()
    FormEdit.Width = FormEdit.Text1.Width
    FormEdit.Height = FormEdit.Text1.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    PaintDisplayArea (Form1.SSTab1.Tab)
    DisplayTicket
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyReturn Then
        'Validate Number Format
        If Len(FormEdit.Text1.Text) > 0 Then
            FormEdit.Text1.Text = Format(Val(FormEdit.Text1.Text), "0.0")
            If Mid(FormEdit.Text1.Text, Len(FormEdit.Text1.Text), 1) <> "0" And Mid(FormEdit.Text1.Text, Len(FormEdit.Text1.Text), 1) <> "5" Then FormEdit.Text1.Text = Format(Val(FormEdit.Text1.Text), "0")
            FormEdit.Text1.Text = Val(FormEdit.Text1.Text)
        End If

        'Update Schedule Record
        If BetType = 0 Then
            If Len(FormEdit.Text1.Text) > 0 Then
                Call dbUpdate("Update Schedule Set Total = " & FormEdit.Text1.Text & " Where GameNumber = " & sGames(GameIndex, (GameOffset * MAX_SCH_ITEMS) + SCH_GAME_NBR))
            Else
                Call dbUpdate("Update Schedule Set Total = '' Where GameNumber = " & sGames(GameIndex, (GameOffset * MAX_SCH_ITEMS) + SCH_GAME_NBR))
            End If
        Else
            If Len(FormEdit.Text1.Text) > 0 Then
                Call dbUpdate("Update Schedule Set Line = " & FormEdit.Text1.Text & " Where GameNumber = " & sGames(GameIndex, (GameOffset * MAX_SCH_ITEMS) + SCH_GAME_NBR))
            Else
                Call dbUpdate("Update Schedule Set Line = '' Where GameNumber = " & sGames(GameIndex, (GameOffset * MAX_SCH_ITEMS) + SCH_GAME_NBR))
            End If
        End If
            
        If BetType = 0 Then
            sGames(GameIndex, (GameOffset * MAX_SCH_ITEMS) + SCH_TOTAL) = FormEdit.Text1.Text
        Else
            sGames(GameIndex, (GameOffset * MAX_SCH_ITEMS) + SCH_LINE) = FormEdit.Text1.Text
        End If
        Unload Me
    End If
End Sub


Private Sub Text1_LostFocus()
    Unload Me
End Sub


