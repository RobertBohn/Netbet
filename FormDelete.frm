VERSION 5.00
Begin VB.Form FormDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete All"
   ClientHeight    =   2850
   ClientLeft      =   2715
   ClientTop       =   2010
   ClientWidth     =   3465
   ControlBox      =   0   'False
   Icon            =   "FormDelete.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbDelete 
      Caption         =   "&Delete All"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   2640
      Picture         =   "FormDelete.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "FormDelete.frx":058C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Clicking 'Delete All' will permanently delete all Ticket Writer data. You won't be able to undo this change."
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FormDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pbCancel_Click()
    Unload FormDelete
End Sub


Private Sub pbDelete_Click()
    On Error Resume Next
    dbClose
    Kill DATABASE
    Kill TOTALS_SHEET
    Kill DAILY_SHEET
    Kill PAYMENT_SHEET
    Kill LINES_SHEET
    Kill DETAIL_SHEET
    Kill ADJUSTMENTS_SHEET
    Unload FormDelete
    Unload Form1
End Sub


