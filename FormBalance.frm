VERSION 5.00
Begin VB.Form FormBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Starting Balance"
   ClientHeight    =   1890
   ClientLeft      =   1260
   ClientTop       =   1845
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtStartingBalance 
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the starting balance for this account"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "FormBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Dim newbal As Long
    
    On Error Resume Next
    
    newbal = Val(Me.txtStartingBalance.Text)
    Call dbUpdate("Update Accounts Set Balance = " & newbal & " Where Account = '" & sCurrentAccount & "'")
    Call dbUpdate("Update Accounts Set StartingBalance = " & newbal & " Where Account = '" & sCurrentAccount & "'")
    Unload Me
End Sub


Private Sub txtStartingBalance_GotFocus()
    Me.txtStartingBalance.SelStart = 0
    Me.txtStartingBalance.SelLength = Len(Me.txtStartingBalance.Text)
End Sub


