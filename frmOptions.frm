VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4200
      Left            =   90
      TabIndex        =   9
      Top             =   135
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7408
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "VScroll1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "HistoryDays"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Port"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Color"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&Advanced"
      TabPicture(2)   =   "frmOptions.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(3)=   "Label9"
      Tab(2).Control(4)=   "Text2"
      Tab(2).Control(5)=   "VScroll2"
      Tab(2).Control(6)=   "Text3"
      Tab(2).Control(7)=   "VScroll3"
      Tab(2).Control(8)=   "Text4"
      Tab(2).Control(9)=   "VScroll4"
      Tab(2).Control(10)=   "Text5"
      Tab(2).Control(11)=   "VScroll5"
      Tab(2).ControlCount=   12
      Begin VB.VScrollBar VScroll5 
         Height          =   330
         Left            =   -74400
         TabIndex        =   32
         Top             =   1980
         Width           =   195
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   -74820
         TabIndex        =   31
         Text            =   "0"
         Top             =   1980
         Width           =   420
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   330
         Left            =   -74400
         TabIndex        =   29
         Top             =   1530
         Width           =   195
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   -74820
         TabIndex        =   28
         Text            =   "0"
         Top             =   1530
         Width           =   420
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   330
         Left            =   -74400
         TabIndex        =   26
         Top             =   1080
         Width           =   195
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   -74820
         TabIndex        =   25
         Text            =   "0"
         Top             =   1080
         Width           =   420
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   330
         Left            =   -74400
         TabIndex        =   23
         Top             =   630
         Width           =   195
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   -74820
         TabIndex        =   22
         Text            =   "0"
         Top             =   630
         Width           =   420
      End
      Begin VB.Frame Port 
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   225
         TabIndex        =   13
         Top             =   1620
         Width           =   1365
         Begin VB.OptionButton Option4 
            Caption         =   "Option4"
            Height          =   240
            Left            =   180
            TabIndex        =   21
            Top             =   1395
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   240
            Left            =   180
            TabIndex        =   20
            Top             =   1050
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.OptionButton opt278 
            Caption         =   "Option2"
            Height          =   240
            Left            =   180
            TabIndex        =   19
            Top             =   705
            Width           =   240
         End
         Begin VB.OptionButton opt378 
            Caption         =   "Option1"
            Height          =   240
            Left            =   180
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   240
         End
         Begin VB.Label Label5 
            Caption         =   "LPT4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   495
            TabIndex        =   17
            Top             =   1050
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Label4 
            Caption         =   "LPT3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   495
            TabIndex        =   16
            Top             =   1395
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lbl0x278 
            Caption         =   "LPT2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   495
            TabIndex        =   15
            Top             =   705
            Width           =   510
         End
         Begin VB.Label lbl0x378 
            Caption         =   "LPT1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   495
            TabIndex        =   14
            Top             =   360
            Width           =   510
         End
      End
      Begin VB.TextBox HistoryDays 
         Enabled         =   0   'False
         Height          =   330
         Left            =   270
         TabIndex        =   11
         Text            =   "0"
         Top             =   675
         Width           =   420
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   690
         Min             =   1
         TabIndex        =   10
         Top             =   675
         Value           =   32767
         Width           =   195
      End
      Begin VB.Label Label9 
         Caption         =   "Keep history for how many days?"
         Height          =   240
         Left            =   -74100
         TabIndex        =   33
         Top             =   2025
         Width           =   3435
      End
      Begin VB.Label Label8 
         Caption         =   "Keep history for how many days?"
         Height          =   240
         Left            =   -74100
         TabIndex        =   30
         Top             =   1575
         Width           =   3435
      End
      Begin VB.Label Label7 
         Caption         =   "IO Send Delay"
         Height          =   240
         Left            =   -74100
         TabIndex        =   27
         Top             =   1125
         Width           =   3435
      End
      Begin VB.Label Label6 
         Caption         =   "IO Receive Deley"
         Height          =   240
         Left            =   -74100
         TabIndex        =   24
         Top             =   675
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Keep history for how many days?"
         Height          =   240
         Left            =   990
         TabIndex        =   12
         Top             =   720
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   8
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   7
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   5
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CounterHist As Integer
Dim DefaultPort As String

Private Sub cmdApply_Click()
    'ToDo: Add 'cmdApply_Click' code.
    SaveSetting App.Title, "Settings", "KeepHistory", HistoryDays.Text
    SaveSetting App.Title, "Settings", "DefaultPort", DefaultPort
    frmMain.RioInterface1.DevicePort = DefaultPort
    frmMain.intKeepHistory = KeepHistory
    frmMain.tvTreeView.Nodes.Clear
    frmMain.tvTreeView.ImagesUseMask = True
    frmMain.tvTreeView.UseImageList = False
    If Trim(CInt(DefaultPort)) = 378 Then
        SaveSetting App.Title, "Settings", "DefaultPort", "LPT1"
        fMainForm.RioInterface1.DevicePort = 378
    Else
        SaveSetting App.Title, "Settings", "DefaultPort", "LPT2"
        fMainForm.RioInterface1.DevicePort = 278
    End If
    
 frmMain.LoadDriveTree
 'LoadFavorites
 frmMain.LoadHistory
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'ToDo: Add 'cmdOK_Click' code.
    SaveSetting App.Title, "Settings", "KeepHistory", HistoryDays.Text
    
    frmMain.intKeepHistory = KeepHistory
    frmMain.tvTreeView.Nodes.Clear
    frmMain.tvTreeView.ImagesUseMask = True
    frmMain.tvTreeView.UseImageList = False
    
    If Trim(CInt(DefaultPort)) = 378 Then
        SaveSetting App.Title, "Settings", "DefaultPort", "LPT1"
        fMainForm.RioInterface1.DevicePort = 378
    Else
        SaveSetting App.Title, "Settings", "DefaultPort", "LPT2"
        fMainForm.RioInterface1.DevicePort = 278
    End If
    
 frmMain.LoadDriveTree
 'LoadFavorites
 frmMain.LoadHistory
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If

End Sub


Private Sub tbsOptions_Click()
    

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub

Private Sub Form_Load()
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    Counter = GetSetting(App.Title, "Settings", "KeepHistory", 4)
    DefaultPort = GetSetting(App.Title, "Settings", "DefaultPort", "0x378")
    
       
    If GetSetting(App.Title, "Settings", "DefaultPort", "LPT1") = "LPT1" Then
        lbl0x378_Click
    Else
        lbl0x278_Click
    End If
    VScroll1.Value = 32767 - Counter
    HistoryDays.Text = Counter
End Sub

Private Sub lbl0x278_Click()
    opt278_Click
    opt278.Value = True
End Sub

Private Sub lbl0x378_Click()
    opt378_Click
    opt378.Value = True
End Sub

Private Sub opt278_Click()
    DefaultPort = "278"
End Sub

Private Sub opt378_Click()
    DefaultPort = "378"
End Sub

Private Sub VScroll1_Change()
    Counter = 32767 - VScroll1.Value
    HistoryDays.Text = Counter
End Sub
