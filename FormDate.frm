VERSION 5.00
Begin VB.Form FormDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date"
   ClientHeight    =   1740
   ClientLeft      =   1035
   ClientTop       =   2190
   ClientWidth     =   2760
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1740
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbOk 
      Caption         =   "&Display"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.VScrollBar sbDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon 1/1/1998"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "FormDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables MUST be defined
Option Explicit


Private Sub pbCancel_Click()
    Unload FormDate
End Sub

Private Sub pbOk_Click()
    Hide
    DisplayScores
    FormScores.Show 1
    Unload FormDate
End Sub


Private Sub sbDate_Change()

    Dim thedate As Date

    thedate = Now
    FormDate.lDate = Format(DateAdd("d", FormDate.sbDate.Value - 500, thedate), "m/d/yyyy ddd")

End Sub


