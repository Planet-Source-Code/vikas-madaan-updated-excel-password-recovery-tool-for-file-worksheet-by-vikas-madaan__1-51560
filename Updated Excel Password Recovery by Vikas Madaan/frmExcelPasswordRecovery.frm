VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExcelPasswordRecovery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excel Password Recovery Tool"
   ClientHeight    =   6345
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   7305
   ClipControls    =   0   'False
   Icon            =   "frmExcelPasswordRecovery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMax 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   18
      Text            =   "1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.VScrollBar ScrlMax 
      Height          =   495
      Left            =   6960
      Max             =   25
      Min             =   1
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5040
      Value           =   1
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtMin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      Text            =   "1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.VScrollBar ScrlMin 
      Height          =   495
      Left            =   6960
      Max             =   25
      Min             =   1
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Value           =   1
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkSheet 
      Caption         =   "Recover Sheet Password Also"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   5280
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.OptionButton optBF 
      Caption         =   "Brute Force Attack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   4800
      Width           =   2055
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Dictionary Attack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   4440
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4005
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   1800
   End
   Begin VB.CommandButton cmdRecover 
      Caption         =   "Recover Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1245
      TabIndex        =   3
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox FName 
      Height          =   375
      Left            =   304
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
   End
   Begin VB.CommandButton cmdSelFile 
      Caption         =   "Select File"
      Height          =   375
      Left            =   5824
      TabIndex        =   1
      Top             =   1055
      Width           =   1335
   End
   Begin VB.TextBox Password 
      Height          =   2055
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2280
      Width           =   6855
   End
   Begin MSComDlg.CommonDialog OpenFile 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Xls"
      DialogTitle     =   "Choose Excel File to Find Password"
   End
   Begin VB.Label lblMax 
      AutoSize        =   -1  'True
      Caption         =   "Max Password Value :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1350
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMin 
      AutoSize        =   -1  'True
      Caption         =   "Min Password Value :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Recovery Method :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4560
      Width           =   2070
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status : Stopped....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   1830
   End
   Begin VB.Label Label2 
      Caption         =   "File Name :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Passwords  :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Speed 
      Height          =   495
      Left            =   315
      TabIndex        =   4
      Top             =   480
      Width           =   6855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmExcelPasswordRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'PROGRAM    :   Updated Excel Password Recovery Tool
'AUTHOR     :   Vikas Madaan
'  __         __        ___      ___
'  \ \       / /       |   \    /   |
'   \ \     / /        | |\ \  / /| |
'    \ \   / /         | | \ \/ / | |
'     \ \_/ /    __    | |  \__/  | |
'      \___/    (__)   |_|        |_|
'
'DATE       :   Feburary 07, 2004.
'
'COMMENTS   :   This is an Excel File Password Recovery Tool Update.
'           It is used to recover password from the Excel File.
'           It also Recover the Password of Sheets within that File.
'           If no Password is set then it shows the relative
'           message at the strarting of checking.
'           It show the usage of Dictionary Attack &
'           Brute Force Attack from 1 to 25 Character Length
'           But you can increase it to any length.
'           when U modify this code & add New Features
'           then please also send me the copy of that.
'           USE FOR EDUCATIONAL & HELPING PURPOSES ONLY!!!
'           If you need support or to give suggestions to improve,
'           you can email me at vikasmadaan25@hotmail.com
'           or thru yahoo messenger vikasmadaan25@yahoo.com
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Option Explicit
Dim Char(62) As String 'Character for Brute Force
Dim CharSet As String 'Include all Characters in Char()
Dim tm As Date  'For Total Time
Dim PCount As Long ' To Check Total Password Checked
Dim PLast As Long 'To Check Last Total
Dim ExcelApp As Object 'Object of Excel
Dim wb 'As excel.Workbook  'For Excel Workbook
Dim ws 'As Worksheet 'For Excel Worksheet
Dim Pass As String 'Hold the Current Password Applied
Dim FPass As String 'Hold the File Password
Dim Find As Boolean 'Contanin True if the Password Found
Dim RecoveryStop As Boolean 'If true the Recovery Process will be Stoped

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This is the main function that checks for the password of
'Excel file it Returns True if Password Found.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Function FindPassword(ByVal Pass As String) As Boolean
On Error GoTo NotFound
PCount = PCount + 1
DoEvents
Set wb = ExcelApp.Workbooks.Open(FName.Text, , True, , Pass)
wb.Close False
FindPassword = True
Exit Function

NotFound:
FindPassword = False
End Function

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This is the main function that checks for the password of
'WorkSheet/Sheet of Excel File, It Returns True if Password Found.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Function FindPasswordSheet(ByVal Pass As String) As Boolean
On Error GoTo NotFound
PCount = PCount + 1
DoEvents
ws.Unprotect Pass
FindPasswordSheet = True
Exit Function

NotFound:
FindPasswordSheet = False
End Function

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function checks Whether the Excel File or
'WorkSheet/Sheet is Password Protected or Not.
'It Returns True if the File is Password Protected.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Function CheckPasswordSet(ByVal CheckBookPass As Boolean) As Boolean
Dim Find As Boolean

If CheckBookPass Then
 Find = FindPassword("")
 If Find Then
  'MsgBox "No Password Set For The File" & vbCrLf & "You Can Open File Without Any Password", vbExclamation, "Excel File Password Recovery"
  Password.Text = Password.Text & FName.Text & vbTab & ":" & vbTab & "No Password Set for the File, You can open File without any Password." & vbCrLf
  FPass = ""
  CheckPasswordSet = False
  Exit Function
 End If
Else
 Find = FindPasswordSheet("")
 If Find Then
  'MsgBox "No Password Set For The Sheet : " & ws.Name & vbCrLf & "You Can Open File Without Any Password", vbExclamation, "Excel File Password Recovery"
  Password.Text = Password.Text & FName.Text & " -> " & ws.Name & vbTab & ":" & vbTab & "No Password Set for the Sheet, You can Make Changes in Sheet without any Password." & vbCrLf
  CheckPasswordSet = False
  Exit Function
 End If
End If
CheckPasswordSet = True
End Function

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function Apply the Dictionary Attack Method on
'Excel File or WorkSheet/Sheet to Recover the Password.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub DictionaryAttack(ByVal CheckBookPass As Boolean)
'Dictionary Attack for File & Sheet
On Error GoTo ErrOccur
Dim Find As Boolean

Open App.Path & "\" & "English.dic" For Input As #1
PCount = 0
PLast = 0
Timer1.Enabled = True
tm = Now

Do Until EOF(1)
  DoEvents
  Line Input #1, Pass
  If CheckBookPass Then
   Find = FindPassword(Pass)
  Else
   Find = FindPasswordSheet(Pass)
  End If
  If Find Or RecoveryStop Then Exit Do
Loop

Timer1_Timer
If RecoveryStop Then
     Password.Text = Password.Text & "Recovery Process Stopped By User....." & vbCrLf
ElseIf Find Then
  'MsgBox "Password Found" & vbCrLf & vbCrLf & "Password=""" & Pass & """", , "Excel File Password Recovery"
  If CheckBookPass Then
   Password.Text = Password.Text & FName.Text & vbTab & ":" & vbTab & Pass & vbCrLf
   FPass = Pass
  Else
   Password.Text = Password.Text & FName.Text & " -> " & ws.Name & vbTab & ":" & vbTab & Pass & vbCrLf
  End If
Else
  'MsgBox "Sorry! Password Not Found", , "Excel File Password Recovery"
  If CheckBookPass Then
   Password.Text = Password.Text & FName.Text & vbTab & ":" & vbTab & "Sorry, Password Not Found." & vbCrLf
   FPass = "File Password Not Found"
  Else
   Password.Text = Password.Text & FName.Text & " -> " & ws.Name & vbTab & ":" & vbTab & "Sorry, Password Not Found." & vbCrLf
  End If
End If

ErrOccur:
Timer1.Enabled = False
Close #1
If Err Then
 MsgBox Err.Description, vbCritical, "Excel File Password Recovery - Error"
End If
End Sub

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function Apply the Brute Force Attack Method on
'Excel File or WorkSheet/Sheet to Recover the Password.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub BruteForceAttack(ByVal CheckBookPass As Boolean)
'Brute Force Attack for File & Sheet
Dim Find As Boolean

PCount = 0
PLast = 0
Timer1.Enabled = True
tm = Now

RecoverPassword CheckBookPass

Timer1.Enabled = False
End Sub

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function help in Recover Password for
'Brute Force Attack.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub RecoverPassword(ByVal CheckBookPass As Boolean)
    Dim MinLen As Integer, MaxLen As Integer
 
    MinLen = Val(txtMin.Text)
    MaxLen = Val(txtMax.Text)
    Find = False
    RecoveryStop = False
    
    ' Continue Generate Passwords Until the Following Happen:
    ' > Password Found.
    ' > Password Length Exceed Max Length.
    ' > User Stop.
    
    'Start from Min Length to Max Length
    Do
        GenPass MinLen, CheckBookPass
        MinLen = MinLen + 1
    Loop Until Find Or MinLen > MaxLen Or RecoveryStop
    
    Timer1_Timer
    ' Determine why password generation stopped.
    If RecoveryStop Then
     Password.Text = Password.Text & "Recovery Process Stopped By User....." & vbCrLf
    ElseIf Find Then
     'MsgBox "Password Found" & vbCrLf & vbCrLf & "Password=""" & Pass & """", , "Excel File Password Recovery"
     If CheckBookPass Then
      Password.Text = Password.Text & FName.Text & vbTab & ":" & vbTab & Pass & vbCrLf
      FPass = Pass
     Else
      Password.Text = Password.Text & FName.Text & " -> " & ws.Name & vbTab & ":" & vbTab & Pass & vbCrLf
     End If
    Else
     'MsgBox "Sorry! Password Not Found", , "Excel File Password Recovery"
     If CheckBookPass Then
      Password.Text = Password.Text & FName.Text & vbTab & ":" & vbTab & "Sorry, Password Not Found." & vbCrLf
      FPass = "File Password Not Found"
     Else
      Password.Text = Password.Text & FName.Text & " -> " & ws.Name & vbTab & ":" & vbTab & "Sorry, Password Not Found." & vbCrLf
     End If
    End If
End Sub

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function Generate Password for
'Brute Force Attack.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub GenPass(ByVal PLen As Integer, ByVal CheckBookPass As Boolean)

Dim i As Double 'for Loop
Dim TotalChar As Integer 'Hold Total Number of Char for Password
Dim InvTotalChar As Single 'Hold the Inverse of Total Char for Password
Dim MaxPass As Double 'Hold the Total Password to Generate
Dim Pos As Integer 'Hold the Current Char Position for Password
Dim Tmp As Double

    TotalChar = UBound(Char)
    InvTotalChar = 1 / TotalChar
    
    ' Calculate Total Passwords to Generate
    MaxPass = TotalChar ^ PLen - 1
    
    Pass = String$(PLen, Left$(CharSet, 1))
    
    For i = 0 To MaxPass
    
        Tmp = i
        Pos = PLen
      
        Do
            Mid$(Pass, Pos, 1) = Char(Tmp Mod TotalChar)
            Pos = Pos - 1
            'Get the Next Char Pos to Change
            Tmp = Int(Tmp * InvTotalChar)
        Loop Until Tmp = 0
        
        DoEvents
        If CheckBookPass Then
         Find = FindPassword(Pass)
        Else
         Find = FindPasswordSheet(Pass)
        End If
        
        If Find Then
            Exit Sub
        ' If user cancels the Process.
        ElseIf RecoveryStop Then
            Exit Sub
        End If
    Next
    
End Sub

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function Disable All Controls when
'Recovery is in Progress.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub DisableAll()
cmdSelFile.Enabled = False
optBF.Enabled = False
optDict.Enabled = False
txtMax.Enabled = False
txtMin.Enabled = False
ScrlMax.Enabled = False
ScrlMin.Enabled = False
chkSheet.Enabled = False
cmdRecover.Caption = "Cancel"
End Sub

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'This Function Enables All Controls when
'Recovery is not in Progress.
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

Private Sub EnableAll()
cmdSelFile.Enabled = True
optBF.Enabled = True
optDict.Enabled = True
txtMax.Enabled = True
txtMin.Enabled = True
ScrlMax.Enabled = True
ScrlMin.Enabled = True
chkSheet.Enabled = True
cmdRecover.Caption = "Recover Password"
End Sub

Private Sub cmdRecover_Click()
If cmdRecover.Caption = "Cancel" Then
 RecoveryStop = True
 GoTo Complete
End If

'Check For File Selected
If Len(FName.Text) = 0 Then
 MsgBox "No File Selected....." & vbCrLf & "Select The File First.....", vbCritical, "Excel File Password Recovery"
 Exit Sub
End If

Dim i As Long

Set ExcelApp = CreateObject("Excel.Application")
'Set wb = CreateObject("Excel.Workbook")
DisableAll
RecoveryStop = False
If optDict.Value Then
 lblStatus = "Status : Checking File For Password Protection....."
 Password.Text = Password.Text & vbCrLf & "Recovering Password Using Dictionary Attack....." & vbCrLf & _
                 "Recovery Process Started At " & Now & vbCrLf & vbCrLf
 'Check for the file is password protected or not
 If CheckPasswordSet(True) Then
  lblStatus = "Status : File Password Recovery in Progress using Dictionary Attack....."
  DictionaryAttack True
 End If
 
 If chkSheet.Value = 1 And FPass <> "File Password Not Found" And RecoveryStop = False Then
  Set wb = ExcelApp.Workbooks.Open(FName.Text, , , , FPass)
  lblStatus = "Status : Sheet Password Recovery in Progress using Dictionary Attack....."
  For i = 1 To wb.Worksheets.Count
   Set ws = wb.Worksheets(i)
   If CheckPasswordSet(False) Then
     DictionaryAttack False
   End If
  Next
 End If

Else
 
 lblStatus = "Status : Checking File For Password Protection....."
 Password.Text = Password.Text & vbCrLf & "Recovering Password Using Brute Force Attack....." & vbCrLf & _
                 "Recovery Process Started At " & Now & vbCrLf & vbCrLf
 'Check for the file is password protected or not
 If CheckPasswordSet(True) Then
  lblStatus = "Status : File Password Recovery in Progress using Brute Force Attack....."
  BruteForceAttack True
 End If
 
 If chkSheet.Value = 1 And FPass <> "File Password Not Found" And RecoveryStop = False Then
  Set wb = ExcelApp.Workbooks.Open(FName.Text, , , , FPass)
  lblStatus = "Status : Sheet Password Recovery in Progress using Brute Force Attack....."
  For i = 1 To wb.Worksheets.Count
   Set ws = wb.Worksheets(i)
   If CheckPasswordSet(False) Then
     BruteForceAttack False
   End If
  Next
 End If
End If

Password.Text = Password.Text & vbCrLf & "Recovery Process Completed At " & Now & vbCrLf

Complete:
On Error Resume Next
lblStatus = "Status : Stopped....."
EnableAll
Set ws = Nothing
wb.Close False
Set wb = Nothing
ExcelApp.Quit
Set ExcelApp = Nothing
End Sub

Private Sub cmdSelFile_Click()
On Error GoTo Cancel
OpenFile.FileName = ""
OpenFile.Filter = "Excel Files (*.Xls)|*.Xls"
OpenFile.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
OpenFile.ShowOpen
FName.Text = OpenFile.FileName
Exit Sub
Cancel:
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
'U can also add any number of Characters
Dim i As Integer, j As Integer
j = 0
For i = Asc("a") To Asc("z")
 Char(j) = Chr(i)
 j = j + 1
Next i
For i = Asc("A") To Asc("Z")
 Char(j) = Chr(i)
 j = j + 1
Next i
For i = Asc("0") To Asc("9")
 Char(j) = Chr(i)
 j = j + 1
Next i
For i = 0 To UBound(Char)
 CharSet = CharSet & Char(i)
Next
RecoveryStop = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If cmdRecover.Caption = "Cancel" Then
 cmdExit_Click
End If
Set wb = Nothing
Set ws = Nothing
End
End Sub

Private Sub optBF_Click()
lblMax.Visible = True
lblMin.Visible = True
txtMax.Visible = True
txtMin.Visible = True
ScrlMax.Visible = True
ScrlMin.Visible = True
ScrlMax.Value = 1
ScrlMin.Value = 1
End Sub

Private Sub optDict_Click()
lblMax.Visible = False
lblMin.Visible = False
txtMax.Visible = False
txtMin.Visible = False
ScrlMax.Visible = False
ScrlMin.Visible = False
End Sub

Private Sub ScrlMax_Change()
txtMax = ScrlMax.Value
If ScrlMax.Value < ScrlMin.Value Then ScrlMin.Value = ScrlMax.Value
End Sub

Private Sub ScrlMin_Change()
txtMin = ScrlMin.Value
If ScrlMax.Value < ScrlMin.Value Then ScrlMax.Value = ScrlMin.Value
End Sub

Private Sub Timer1_Timer()
Speed.Caption = "Speed/Sec = " & PCount - PLast & "       Time = " & Format$(Now - tm, "hh:mm:ss") & vbCrLf & "Total = " & PCount & "       Current Password = " & Pass
PLast = PCount
End Sub

