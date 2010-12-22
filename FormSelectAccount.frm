VERSION 5.00
Begin VB.Form FormSelectAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Account"
   ClientHeight    =   1755
   ClientLeft      =   900
   ClientTop       =   2355
   ClientWidth     =   3060
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1755
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton pbOk 
      Caption         =   "&Display"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton pbClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cbAccount 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "cbAccount"
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "FormSelectAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAccount_KeyPress(KeyAscii As Integer)
    
    Dim lResult As Long
   
    'Convert to uppercase
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = KeyAscii - Asc("a") + Asc("A")
    
    'Select an Account when Enter is pressed
    If KeyAscii = vbKeyReturn And Len(FormSelectAccount.cbAccount) > 0 Then
    
ReadAgain:
        'Select the Account Record From the Accounts Table
        sCurrentAccount = FormSelectAccount.cbAccount.Text
        lResult = dbSelect("Select * From Accounts Where Account = '" & sCurrentAccount & "'")
        If lResult <> 0 Then Exit Sub
    
        'Check if an Account was Selected
        If MyTable.EOF Then   'The Selected Account Was Not Found
            MyTable.Close
            lResult = MessageBox(0, "Account '" & sCurrentAccount & "' Was Not Found. Do You Want To Add This Account?", Form1.Caption, MB_YESNO)
            If lResult = IDNO Then Exit Sub
        
            lResult = dbInsert("Insert Into Accounts (Account,Status,Balance) Values (""" & sCurrentAccount & """, ""A"", 0)")
            If lResult <> 0 Then Exit Sub
            FormBalance.Show 1
            FormSelectAccount.cbAccount.AddItem sCurrentAccount
            GoTo ReadAgain
        Else
            MyTable.Close
        End If
    
        sCurrentAccount = FormSelectAccount.cbAccount.Text
        ReadAccount
        DisplayAccountDetail
        FormDetail.Show 1
    End If
    
    
    
    
    
    
    
    
    
    
     


  
    
    
    
    
    
    
    
    
End Sub


Private Sub pbClose_Click()
    Unload FormSelectAccount
End Sub


Private Sub pbOk_Click()
    Dim lResult As Long

    If Len(FormSelectAccount.cbAccount.Text) > 12 Then
        FormSelectAccount.cbAccount.Text = Mid(FormSelectAccount.cbAccount.Text, 1, 12)
    End If

ReadAgain:
    'Select the Account Record From the Accounts Table
    sCurrentAccount = FormSelectAccount.cbAccount.Text
    lResult = dbSelect("Select * From Accounts Where Account = '" & sCurrentAccount & "'")
    If lResult <> 0 Then Exit Sub
    
    'Check if an Account was Selected
    If MyTable.EOF Then   'The Selected Account Was Not Found
        MyTable.Close
        lResult = MessageBox(0, "Account '" & sCurrentAccount & "' Was Not Found. Do You Want To Add This Account?", Form1.Caption, MB_YESNO)
        If lResult = IDNO Then Exit Sub
        
        lResult = dbInsert("Insert Into Accounts (Account,Status,Balance) Values (""" & sCurrentAccount & """, ""A"", 0)")
        If lResult <> 0 Then Exit Sub
        FormBalance.Show 1
        FormSelectAccount.cbAccount.AddItem sCurrentAccount
        GoTo ReadAgain
    Else
        MyTable.Close
    End If
    
    sCurrentAccount = FormSelectAccount.cbAccount.Text
    ReadAccount
    DisplayAccountDetail
    FormDetail.Show 1
End Sub


