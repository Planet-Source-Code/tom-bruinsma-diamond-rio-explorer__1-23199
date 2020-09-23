VERSION 5.00
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "NSLOCK15VB5.OCX"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Registration Key"
   ClientHeight    =   1710
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4785
   ControlBox      =   0   'False
   HelpContextID   =   130
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1010.326
   ScaleMode       =   0  'User
   ScaleWidth      =   4492.854
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login Form"
   Begin VB.CommandButton Command2 
      Caption         =   "&Continue"
      Enabled         =   0   'False
      Height          =   540
      Left            =   3465
      TabIndex        =   10
      Top             =   4230
      Width           =   1170
   End
   Begin VB.TextBox txtEnc 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Software code"
      Top             =   120
      Width           =   2655
   End
   Begin nslock15vb5.ActiveLock ActiveLock1 
      Left            =   90
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "Sh0wMeTheM$ney"
      SoftwareName    =   "Minds Eye Productions Rio Millennium"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   315
      Top             =   3570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R&un Unregistered"
      Height          =   540
      HelpContextID   =   130
      Left            =   3510
      TabIndex        =   3
      Tag             =   "Unregistered"
      Top             =   1080
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Register"
      Default         =   -1  'True
      Height          =   540
      HelpContextID   =   130
      Left            =   1980
      TabIndex        =   2
      Tag             =   "Register"
      Top             =   1095
      Width           =   1170
   End
   Begin VB.TextBox REGISTRATION 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4800
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Software Code:"
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click F1 for Registration Information or read Register.txt in the installation directory."
      Height          =   405
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Register Please"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   5
      Top             =   3570
      Width           =   4320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   105
      TabIndex        =   4
      Top             =   1980
      Width           =   4740
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Registration Key :"
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1605
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Dim USERNAME As String

Option Explicit
Dim wait As Integer

Private Sub cmdOK_Click()
'This is the register command
'we will encrypt the key generated by ActiveLock and give to the user
'instead.
'So even if the user gets hold of a copy of the Key Generator, And
'your program password as well he still can't register the program.
'because we are not really using the original key but we are using the
'encyrpted key.
'SO, we have Double Lock.
   REGISTRATION.Text = Decrypt(txtEnc.Text)
Dim USER As String
ActiveLock1.LiberationKey = REGISTRATION
    If (ActiveLock1.RegisteredUser) Then
      MsgBox "Thank You for registering MEPRioMillennium."
       USER = USERNAME
       SaveSetting "MEPRioMillennium", "Startup", "Registereduser", USER
         Unload Me
       TheLock!unreg1.Visible = False  'label in the main form
       TheLock!unreg2.Visible = False  'label in the main form
       TheLock.Show
    Else
        MsgBox "Invalid Registration Key, try again!", , "Unlock Failed !"
        txtEnc.SetFocus
        SendKeys "{Home}+{End}"
    End If
'make sure that the user can run your helpfile if he clicks F1
'to get information about registering (link the topic as well)
'no need to let your user find things the hard way if he can do it
'by a click
'if you use txt file then link it instead
End Sub
Private Sub Command1_Click()
'we put a file on the user hard disk
'do not use the app path because we do not want this file removed if
'program was uninstalled
Timer1.Enabled = True
Const FILENAME = "\MEPRIO.ini"
syspath = WindowsDirectory
If Dir(syspath & FILENAME) <> "MEPRIO.ini" Then
Dim SYS As String
SYS = syspath & "\MEPRIO.ini"
'Put something in the file we do not want a 0 byte file.
'do not make it a real ini file because if you do uninstall programs
'will detect it and will remove it.
Open SYS For Random As #1
Put #1, , "BY A WINDOWS PROGRAM SETUP DON'T DELETE THIS FILE IT IS NEEDED TO RUN A PROGRAM  "
Put #1, , "SETTYPE = 0881738891 , ATTRIB = 0 , SYSTEMDATE = LONGDATE , SETUPTYPE = COMPLETE"
Close #1
SaveSetting "MEPRioMillennium", "Startup", "STPD", Date   'save setup date
SaveSetting "MEPRioMillennium", "Startup", "XPD", Date + 30 'save expiry date
Unload Me
TheLock.Show
End If

Dim ready As String
ready = GetSetting("MEPRioMillennium", "Startup", "XPD") 'get expiry date
If ready = "" Then                                    'if not there
frmLogin.Height = 5195                                'lock the program
frmLogin.Top = 2250
frmLogin.Width = 5040
Label1.Caption = "Your Evaluation Period  is over ! Either Register Rio Millennium or remove it from your system. Thank you for trying The Rio Millennium."
Command1.Enabled = False
Else:
SaveSetting "MEPRioMillennium", "Startup", "STPD", Date 'save today date
TheLock.Show
Unload Me
Exit Sub
End If
'the above was only to display the time left in the evaluation version
'and also as a decoy

'now we will do these checks using Active Lock ocx
'please refer to Active Lock help for details.
If ActiveLock1.LastRunDate > Now Then  'check if clock was set backwards
MsgBox "Your system clock has been set backwards,Please reset your system clock, MEPRioMillennium will now exit, Thank you for using MEPRioMillennium"
Unload Me
'End
End If
If ActiveLock1.UsedDays > 30 Then 'check if used more than 30 days
frmLogin.Height = 4695
frmLogin.Top = 2250
frmLogin.Width = 5040
Label1.Caption = "Your Evaluation Period  is over ! Either Register MEPRioMillennium or remove it from your system. Thank you for trying MEPRioMillennium."
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Form_Load()
Dim sBuffer As String
    Dim lSize As Long

'GET THE USERNAME
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
 
 If lSize > 0 Then
        USERNAME = Left$(sBuffer, lSize - 1)
    Else
        USERNAME = "something"
    End If
'if program already registered no need to show me
'load main program form
    Text1.Text = ActiveLock1.SoftwareCode
End Sub
Private Sub Timer1_Timer() 'to flash register
wait = wait + 1
If wait > 15 Then Me.Command2.Enabled = True
If Label2.Visible = True Then
Label2.Visible = False
Else
Label2.Visible = True
End If
End Sub
Public Function Encrypt(ByVal Plain As String)
    Dim i
    Dim Letter As String
    For i = 1 To Len(Plain)
        Letter = Mid$(Plain, i, 1)
        Mid$(Plain, i, 1) = Chr(Asc(Letter) + 1)
    Next i
    Encrypt = Plain
End Function
Public Function Decrypt(ByVal Encrypted As String)
Dim i
Dim Letter As String
    For i = 1 To Len(Encrypted)
        Letter = Mid$(Encrypted, i, 1)
        Mid$(Encrypted, i, 1) = Chr(Asc(Letter) - 1)
    Next i
    Decrypt = Encrypted
End Function
