VERSION 5.00
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "NSLOCK15VB5.OCX"
Begin VB.Form TheLock 
   Caption         =   "Rio Millennium"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Timer Timer3 
      Left            =   240
      Top             =   240
   End
   Begin nslock15vb5.ActiveLock ActiveLock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "Sh0wMeTheM$ney"
      SoftwareName    =   "Minds Eye Productions Rio Millennium"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label unreg1 
      AutoSize        =   -1  'True
      Caption         =   "Unregistered Evaluation Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label unreg2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "TheLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    frmMain.Show
    Unload frmLogin
    Unload TheLock
    'End
End Sub

Private Sub Form_Load()
    'check if user is registered
    Dim USER As String
    USER = GetSetting("MEPRioMillennium", "Startup", "Registereduser", "")
    If USER = "" Then
    LOCKPROG 'call lock function
    Else
    End If
    ' display software code for the user even after register
    'it is handy for the user to write down if he decide to register
    'because no need to restart the program just to get the software code
    'from login form
    Label1.Caption = "Software Code: " & ActiveLock1.SoftwareCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'End
End Sub

Private Sub Timer3_Timer()
    On Error Resume Next
    'if activated by Lockprog it deletes the settings
    'Why? to make the diff = 0 case
    'which will lock the program
    'and simply unloads it.
    DeleteSetting "MEPRioMillennium", "Startup", "stpd"
    DeleteSetting "MEPRioMillennium", "Startup", "xpd"
    MsgBox "Your Evaluation Period  is over ! Either Register Rio Millennium or remove it from your system. Thank you for trying Rio Millennium.", , "Thank you for trying Rio Millennium"
    Unload Me
    End
End Sub

Public Function LOCKPROG()

    Dim DIFF As String, SD As String, ED As String
    SD = GetSetting("MEPRioMillennium", "Startup", "STPD", "31/12/99")
    ED = GetSetting("MEPRioMillennium", "Startup", "XPD", "31/12/99")
    'make sure you 31/2/99 or similar in both of them
    'in case the user delelte the settings form the registry
    'this way you will have a diff = 0 case
    'and the program will lock.
    'get the dates and see the diff between them
    'you can use Select Case but I wanted to make it very clear
    DIFF = DateDiff("D", SD, ED)
    
    If DIFF = 0 Then
        Timer3.Interval = 100
    Else
      unreg1.Visible = True
      unreg2.Visible = True
    End If
    
    If DIFF < 0 Then
    Timer3.Interval = 100
    End If
    
    If DIFF > 30 Then
    Timer3.Interval = 100
    End If
    unreg2.Caption = "Your Evaluation Period Expires After   " & DIFF & "  Day(s)"
End Function

