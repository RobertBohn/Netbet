VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form FormOpenExcel 
   Caption         =   "FormOpenExcel"
   ClientHeight    =   5190
   ClientLeft      =   615
   ClientTop       =   930
   ClientWidth     =   6915
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5190
   ScaleWidth      =   6915
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DialogTitle     =   "Select Excel Spreadsheet"
      Filter          =   "Microsoft Excel Files|*.xls"
   End
End
Attribute VB_Name = "FormOpenExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
