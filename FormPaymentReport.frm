VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormPaymentReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments Report"
   ClientHeight    =   5100
   ClientLeft      =   495
   ClientTop       =   795
   ClientWidth     =   7650
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton pbPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin MSGrid.Grid GridPaymentReport 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   7646
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
      Cols            =   10
      FixedCols       =   0
      HighLight       =   0   'False
   End
   Begin VB.Label lToday 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   4095
   End
End
Attribute VB_Name = "FormPaymentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pbClose_Click()
    Unload FormPaymentReport
End Sub


Private Sub pbPrint_Click()
    Dim nRow, nCol As Long
    Dim MyXL As Object                  'Variable to hold reference to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   'Flag for final release.

    FormPaymentReport.MousePointer = vbHourglass

    ' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    ' GetObject function called without the first argument returns a
    ' reference to an instance of the application. If the application isn't
    ' running, an  error occurs. Note the comma used as the first argument placeholder.
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

    ' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject(PAYMENT_SHEET)
    'MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True

    MyXL.Application.Rows("4:500").Select
    MyXL.Application.Selection.ClearContents
    MyXL.Application.Range("A1").Select

    MyXL.Application.Cells(1, 1).Value = Format(Now(), "m/d/yyyy")
 
    For nRow = 0 To FormPaymentReport.GridPaymentReport.Rows - 1
        For nCol = 0 To 9
            FormPaymentReport.MousePointer = vbHourglass
            FormPaymentReport.GridPaymentReport.Row = nRow
            FormPaymentReport.GridPaymentReport.Col = nCol
            MyXL.Application.Cells(nRow + 3, nCol + 1).Value = FormPaymentReport.GridPaymentReport.Text
        Next nCol
    Next nRow
    
    MyXL.Application.ActiveWorkbook.PrintOut Copies:=1
    MyXL.Application.ActiveWorkbook.Save
    If ExcelWasNotRunning = True Then MyXL.Application.Quit
    Set MyXL = Nothing  ' Release reference to the application and spreadsheet.

    FormPaymentReport.MousePointer = vbDefault

End Sub


