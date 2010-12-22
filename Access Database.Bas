Attribute VB_Name = "Module2"
'All variables MUST be defined
Option Explicit

'System Settings
Public STOP_WORKING_DATE As String
Public MY_PASSWORD As String
Public ALLOW_DATE_TO_BE_SET_BACK As Boolean
Public DEPOSIT_SYSTEM As Boolean
Public DELETE_ALL_BUTTON As Boolean
Public WAGER_AFTER_KICKOFF As Boolean
Public ALL_TIES_PUSH As Boolean
Public NBA_TODAY_ONLY As Boolean
Public NHL_TODAY_ONLY As Boolean
Public MLB_TODAY_ONLY As Boolean
Public NFL_TODAY_ONLY As Boolean
Public DATABASE As String
Public TOTALS_SHEET As String
Public DAILY_SHEET As String
Public PAYMENT_SHEET As String
Public LINES_SHEET As String
Public DETAIL_SHEET As String
Public ADJUSTMENTS_SHEET As String
Public IMPORT_FILE_PATH As String

Public Function dbOpen() As Long
    Dim sDate, sCode As String
    Dim i, crc  As Integer
    
    'Assume Trouble
    dbOpen = 1
    On Error GoTo dbOpenErrorHandler
    
    'date  02/01/1998
    'codw  ac ab bjji
    'mid   4 9 / 12 18 / 21 28 35 44
    '
    '         1         2         3         4         5
    '12345678901234567890123456789012345678901234567890
    'm93af3SlcDoAm39WnblOBqX6wMlJqWNPb5j9awMlWq8iEpleWQnP7
GoTo NoBadDate
    sCode = STOP_WORKING_DATE
    crc = 0
    For i = 1 To Len(sCode)
        crc = crc + Asc(Mid(sCode, i, 1))
    Next i
    If crc <> 5000 Then GoTo badDate
    sCode = Format(sCode, ">")
    If Len(sCode) < 50 Then GoTo badDate
    sDate = Asc(Mid(sCode, 4, 1)) - Asc("A")
    sDate = sDate & (Asc(Mid(sCode, 9, 1)) - Asc("A")) & "/"
    sDate = sDate & (Asc(Mid(sCode, 12, 1)) - Asc("A"))
    sDate = sDate & (Asc(Mid(sCode, 18, 1)) - Asc("A")) & "/"
    sDate = sDate & (Asc(Mid(sCode, 21, 1)) - Asc("A"))
    sDate = sDate & (Asc(Mid(sCode, 28, 1)) - Asc("A"))
    sDate = sDate & (Asc(Mid(sCode, 35, 1)) - Asc("A"))
    sDate = sDate & (Asc(Mid(sCode, 44, 1)) - Asc("A"))
    If IsDate(sDate) = False Then GoTo badDate
    If DateDiff("d", sDate, Now) >= 0 Then GoTo badDate
    If DateDiff("d", sDate, "01/02/2010") <= 0 Then GoTo badDate
    If DateDiff("d", sDate, "01/01/1997") >= 0 Then GoTo badDate
    
    'Be sure the system date has not been set back
    If DateDiff("s", FileDateTime(DATABASE), Now) < 0 Then
        If ALLOW_DATE_TO_BE_SET_BACK = False Then
            Call MessageBox(0, "This Program Will Not Run If The System Date Has Been Set Back.", "System Date Set Back", MB_OK)
            Exit Function
        End If
    End If
NoBadDate:
    'Open the Database
    Set MyDb = DBEngine.OpenDatabase(DATABASE, False, False, ";pwd=" & MY_PASSWORD)
        
    dbOpen = 0
    Exit Function

badDate:
    Call MessageBox(0, "The License For This Program Has Expired", "License Experation", MB_OK)
    Exit Function

'Error Handler
dbOpenErrorHandler:
    Select Case Err
        Case 3031   'Password.
            Load FormPassword
            FormPassword.Show vbModal
            If FormPassword.OK = True Then
                Call SaveSetting("Ticket Writer", "Settings", "Password", FormPassword.txtPassword.Text)
                MY_PASSWORD = FormPassword.txtPassword.Text
                Unload FormPassword
                Resume
            Else
                Unload FormPassword
            End If
            
        Case Else   'Misc Open Error
            MsgBox "Error in dbOpen. " & Error$ & " " & Str$(Err)
    End Select
    dbOpen = 1
End Function

Public Sub dbClose()
    
    'Set up Basic Error Handler
    On Error GoTo dbCloseErrorHandler
    
    'Close the Database
    If Not MyDb Is Nothing Then
        MyDb.Close
    End If
    Exit Sub

'Error Handler
dbCloseErrorHandler:
    MsgBox "Error in dbClose. " & Error$ & " " & Str$(Err)
    Err.Clear
    
End Sub


Public Function dbSelect(ByVal Statement As String) As Long
    
    'Set up Basic Error Handler
    On Error GoTo dbSelectErrorHandler
    
    'Perform the Select Statement
    Set MyTable = MyDb.OpenRecordset(Statement, dbOpenDynaset)
    
    'Set Return Value to 0
    dbSelect = 0
    
    Exit Function
 
dbSelectErrorHandler:
    MsgBox "Error in dbSelect. " & Error$ & " " & Str$(Err)
    Err.Clear
    dbSelect = 1
     
End Function

Public Function dbInsert(ByVal Statement As String) As Long
    
    'Set up Basic Error Handler
    On Error GoTo dbInsertErrorHandler
    
    'Perform the Select Statement
    MyDb.Execute Statement
    
    'Set Return Value to 0
    dbInsert = 0
    
    Exit Function
 
dbInsertErrorHandler:
    MsgBox "Error in dbInsert. " & Error$ & " " & Str$(Err)
    Err.Clear
    dbInsert = 1
 
End Function



Public Function dbUpdate(ByVal Statement As String) As Long
    
    'Set up Basic Error Handler
    On Error GoTo dbUpdateErrorHandler
    
    'Perform the Update Statement
    MyDb.Execute Statement
    
    'Set Return Value to 0
    dbUpdate = 0
    
    Exit Function
 
dbUpdateErrorHandler:
    MsgBox "Error in dbUpdate. " & Error$ & " " & Str$(Err)
    Err.Clear
    dbUpdate = 1
    
End Function

Public Function dbDelete(ByVal Statement As String) As Long
    
    'Set up Basic Error Handler
    On Error GoTo dbDeleteErrorHandler
    
    'Perform the Select Statement
    MyDb.Execute Statement
    
    'Set Return Value to 0
    dbDelete = 0
    
    Exit Function
 
dbDeleteErrorHandler:
    MsgBox "Error in dbDelete. " & Error$ & " " & Str$(Err)
    Err.Clear
    dbDelete = 1
     
End Function
