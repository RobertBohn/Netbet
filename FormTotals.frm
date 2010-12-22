VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormTotals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totals Report"
   ClientHeight    =   5355
   ClientLeft      =   1170
   ClientTop       =   705
   ClientWidth     =   2655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton pbPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin MSGrid.Grid GridTotals 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   8070
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
      FixedCols       =   0
      HighLight       =   0   'False
   End
End
Attribute VB_Name = "FormTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pbClose_Click()
    Unload FormTotals
End Sub


Private Sub pbPrint_Click()
    Dim i As Long
    Dim count(4) As Long
    Dim nRow As Long
    Dim nCol As Long
    Dim MyXL As Object                  'Variable to hold reference to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   'Flag for final release.

    FormTotals.MousePointer = vbHourglass

    ' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    ' GetObject function called without the first argument returns a
    ' reference to an instance of the application. If the application isn't
    ' running, an  error occurs. Note the comma used as the first argument placeholder.
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

    ' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject(TOTALS_SHEET)
    'MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True

    MyXL.Application.Rows("5:500").Select
    MyXL.Application.Selection.ClearContents
    MyXL.Application.Range("A1").Select

    count(0) = Int((FormTotals.GridTotals.Rows - 1) / 4)
    count(1) = Int((FormTotals.GridTotals.Rows - 1) / 4)
    count(2) = Int((FormTotals.GridTotals.Rows - 1) / 4)
    count(3) = Int((FormTotals.GridTotals.Rows - 1) / 4)
    
    If (FormTotals.GridTotals.Rows - 1) Mod 4 >= 1 Then count(0) = count(0) + 1
    If (FormTotals.GridTotals.Rows - 1) Mod 4 >= 2 Then count(1) = count(1) + 1
    If (FormTotals.GridTotals.Rows - 1) Mod 4 >= 3 Then count(2) = count(2) + 1
     
    MyXL.Application.Cells(1, 1).Value = Format(Now(), "m/d/yyyy")
    For i = 1 To FormTotals.GridTotals.Rows - 1
        FormTotals.MousePointer = vbHourglass
            
        If i <= count(0) Then
            nCol = 1
            nRow = i + 4
        End If
       
        If i > count(0) And i <= (count(0) + count(1)) Then
            nCol = 4
            nRow = i - count(0) + 4
        End If
       
        If i > (count(0) + count(1)) And i <= (count(0) + count(1) + count(2)) Then
            nCol = 7
            nRow = i - count(0) - count(1) + 4
        End If
       
        If i > (count(0) + count(1) + count(2)) Then
            nCol = 10
            nRow = i - count(0) - count(1) - count(2) + 4
        End If
        
        
'        nCol = 1
'        nRow = i + 4
        
'        If i > (FormTotals.GridTotals.Rows + 1) / 4 Then
'            nRow = i + 4 - (FormTotals.GridTotals.Rows / 4)
'            nCol = 4
'        End If
            
'        If i > ((FormTotals.GridTotals.Rows + 1) * 2) / 4 Then
'            nRow = i + 4 - ((FormTotals.GridTotals.Rows * 2) / 4)
'            nCol = 7
'        End If
        
'        If i > ((FormTotals.GridTotals.Rows + 1) * 3) / 4 Then
'            nRow = i + 4 - ((FormTotals.GridTotals.Rows * 3) / 4)
'            nCol = 10
'        End If
        
        FormTotals.GridTotals.Row = i
        FormTotals.GridTotals.Col = 0
        MyXL.Application.Cells(nRow, nCol).Value = FormTotals.GridTotals.Text
        FormTotals.GridTotals.Col = 1
        MyXL.Application.Cells(nRow, nCol + 1).Value = FormTotals.GridTotals.Text
    Next i

    MyXL.Application.ActiveWorkbook.PrintOut Copies:=1
    MyXL.Application.ActiveWorkbook.Save
    If ExcelWasNotRunning = True Then MyXL.Application.Quit
    Set MyXL = Nothing  ' Release reference to the application and spreadsheet.

    FormTotals.MousePointer = vbDefault

End Sub

