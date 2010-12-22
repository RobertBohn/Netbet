VERSION 5.00
Begin VB.Form FormPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Required"
   ClientHeight    =   1620
   ClientLeft      =   4200
   ClientTop       =   2760
   ClientWidth     =   3000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter database password:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OK As Boolean
Private Sub cmdCancel_Click()
    OK = False
    Hide
End Sub


Private Sub cmdOk_Click()
    OK = True
    Hide
End Sub


