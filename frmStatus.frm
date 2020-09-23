VERSION 5.00
Object = "{571D9D02-EE3C-11D2-BC70-004005448951}#1.3#0"; "SWBPROG.OCX"
Begin VB.Form frmStatus 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SWBProgressBar.SWBProgress ProgressBarFile 
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   930
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      Percent         =   50
      BarColor        =   65280
      PercentColor    =   0
   End
   Begin SWBProgressBar.SWBProgress ProgressBarTotal 
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      Percent         =   50
      BarColor        =   65280
      PercentColor    =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   3
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


    
Private Sub Form_Load()
Dim HWND_TOPMOST
HWND_TOPMOST = -1

    SetWindowPos Me.hwnd, HWND_TOPMOST, _
        Me.Left / Screen.TwipsPerPixelX, _
        Me.Top / Screen.TwipsPerPixelY, _
        Me.Width / Screen.TwipsPerPixelX, _
        Me.Height / Screen.TwipsPerPixelY, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub
