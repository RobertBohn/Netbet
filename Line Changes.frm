VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormLines 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Line Changes"
   ClientHeight    =   5970
   ClientLeft      =   225
   ClientTop       =   630
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton pbOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin MSGrid.Grid GridLines 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   9340
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
      Cols            =   9
      FixedCols       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.Label lblTotals 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lblSides 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   5520
      Width           =   615
   End
End
Attribute VB_Name = "FormLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'All variables MUST be defined
Option Explicit

Private Sub Label1_Click()

End Sub


Private Sub Grid1_RowColChange()

End Sub


Private Sub GridLines_GotFocus()
    FormLines.GridLines.Tag = "0"
End Sub

Private Sub GridLines_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i, x, cur As Long

    'check if a valid key
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyDecimal _
    Or (KeyCode >= vbKey0 And KeyCode <= vbKey9) _
    Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) _
    Or KeyCode = vbKeySubtract Or KeyCode = vbKeyAdd _
    Or KeyCode = vbKeyBack _
    Or KeyCode = 189 Or KeyCode = 190 Or KeyCode = 187 _
    Then

        'get the current game
         cur = 0
         For i = 0 To SPORTS - 1
            For x = 0 To nGames(i) - 1
                
                If (cur * 2) + 1 = FormLines.GridLines.Row Then
                    'on road team line
                    If i = NFL_TAB Or i = NBA_TAB Then Exit Sub
                    If Not (FormLines.GridLines.Col = 6 Or FormLines.GridLines.Col = 8) Then Exit Sub
                End If
                
                If (cur * 2) + 2 = FormLines.GridLines.Row Then
                    'on home team line
                    If (i = NFL_TAB Or i = NBA_TAB) And Not (FormLines.GridLines.Col = 5 Or FormLines.GridLines.Col = 7) Then Exit Sub
                    If (i = MLB_TAB) And Not (FormLines.GridLines.Col = 6 Or FormLines.GridLines.Col = 7 Or FormLines.GridLines.Col = 8) Then Exit Sub
                    If (i = NHL_TAB) And Not (FormLines.GridLines.Col = 5 Or FormLines.GridLines.Col = 6 Or FormLines.GridLines.Col = 7 Or FormLines.GridLines.Col = 8) Then Exit Sub
                End If
                
                cur = cur + 1
            Next x
        Next i

        'Apply the key
        
        If Val(FormLines.GridLines.Tag) = 0 Then
            FormLines.GridLines.Tag = "1"
            FormLines.GridLines.Text = ""
        End If
        
        If KeyCode = vbKeyDelete Then
            FormLines.GridLines.Text = ""
        End If
    
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
           FormLines.GridLines.Text = FormLines.GridLines.Text & (KeyCode - vbKey0)
        End If
        
        If KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & (KeyCode - vbKeyNumpad0)
        End If
    
        If KeyCode = vbKeyDecimal Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & "."
        End If

        If KeyCode = vbKeySubtract Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & "-"
        End If
    
        If KeyCode = vbKeyAdd Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & "+"
        End If
        
        If KeyCode = vbKeyBack And Len(FormLines.GridLines.Text) > 0 Then
            FormLines.GridLines.Text = Mid(FormLines.GridLines.Text, 1, Len(FormLines.GridLines.Text) - 1)
        End If
        
        If KeyCode = 187 Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & "+"
        End If
        
        If KeyCode = 189 Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & "-"
        End If
        
        If KeyCode = 190 Then
            FormLines.GridLines.Text = FormLines.GridLines.Text & "."
        End If
        
    End If

End Sub
Private Sub GridLines_LostFocus()
    FormLines.GridLines.Tag = "0"
End Sub

Private Sub GridLines_RowColChange()
    FormLines.GridLines.Tag = "0"
End Sub

Private Sub GridLines_SelChange()
    FormLines.GridLines.Tag = "0"
End Sub


Private Sub pbExit_Click()
    ApplyLineChanges
    Unload FormLines
End Sub




Private Sub pbCancel_Click()
    Unload FormLines
End Sub

Private Sub pbOk_Click()
    FormLines.MousePointer = vbHourglass
    ApplyLineChanges
    FormLines.MousePointer = vbDefault
    Unload FormLines
End Sub


Private Sub pbPrint_Click()
    Dim nRow, nCol As Long
    Dim MyXL As Object                  'Variable to hold reference to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean   'Flag for final release.

    FormLines.MousePointer = vbHourglass

    ' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.

    ' GetObject function called without the first argument returns a
    ' reference to an instance of the application. If the application isn't
    ' running, an  error occurs. Note the comma used as the first argument placeholder.
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    Err.Clear   ' Clear Err object in case error occurred.

    ' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject(LINES_SHEET)
    'MyXL.Application.Visible = True
    MyXL.Parent.Windows(1).Visible = True

    MyXL.Application.Rows("1:500").Select
    MyXL.Application.Selection.ClearContents
    MyXL.Application.Range("A1").Select
  
    For nRow = 0 To FormLines.GridLines.Rows - 1
        For nCol = 0 To 8
            FormLines.MousePointer = vbHourglass
            FormLines.GridLines.Row = nRow
            FormLines.GridLines.Col = nCol
            MyXL.Application.Cells(nRow + 1, nCol + 1).Value = FormLines.GridLines.Text
        Next nCol
    Next nRow
    
    MyXL.Application.Cells(nRow + 3, 1).Value = "Sides Handle " & FormLines.lblSides.Caption
    MyXL.Application.Cells(nRow + 4, 1).Value = "Totals Handle " & FormLines.lblTotals.Caption
    
    MyXL.Application.ActiveWorkbook.PrintOut Copies:=1
    MyXL.Application.ActiveWorkbook.Save
    If ExcelWasNotRunning = True Then MyXL.Application.Quit
    Set MyXL = Nothing  ' Release reference to the application and spreadsheet.

    FormLines.MousePointer = vbDefault

End Sub


