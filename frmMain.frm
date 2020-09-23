VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTREE.OCX"
Object = "{571D9D02-EE3C-11D2-BC70-004005448951}#1.3#0"; "SWBPROG.OCX"
Begin VB.Form frmMain 
   Caption         =   "RioMillennium"
   ClientHeight    =   7755
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   9495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7755
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7140
      Top             =   3480
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   5715
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9495
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   9495
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rio Millennium"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Tag             =   " TreeView:"
         Top             =   12
         Width           =   2016
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RIO Contents:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   6
         Tag             =   " ListView:"
         Top             =   12
         Width           =   3216
      End
   End
   Begin SWBProgressBar.SWBProgress ProgressBar1 
      Height          =   255
      Left            =   2130
      TabIndex        =   4
      Top             =   5190
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   450
      BarColor        =   65280
      PercentColor    =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7140
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":203E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   3
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4440
      Left            =   2100
      TabIndex        =   2
      Top             =   720
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   7832
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   65280
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reload"
            Object.ToolTipText     =   "Reload Rio Playlist"
            ImageKey        =   "reload"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sync"
            ImageKey        =   "Sync"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Upload"
            Object.ToolTipText     =   "Upload"
            ImageKey        =   "Upload"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RandomFolder"
            Object.ToolTipText     =   "Random Folder"
            ImageKey        =   "RandomFolder"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RandomPlaylist"
            Object.ToolTipText     =   "Random From Playlist"
            ImageKey        =   "RandomPlaylist"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OneFolder"
            Object.ToolTipText     =   "One Touch Folder"
            ImageKey        =   "OneFolder"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OnePlaylist"
            Object.ToolTipText     =   "One Touch Playlist"
            ImageKey        =   "OnePlaylist"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OneFolderCurrent"
            Object.ToolTipText     =   "One Touch Folder Current"
            ImageKey        =   "OneFolderCurrent"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OnePlaylistCurrent"
            Object.ToolTipText     =   "One Touch Playlist Current"
            ImageKey        =   "OnePlaylistCurrent"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Selected"
            ImageKey        =   "Trash"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteAll"
            Object.ToolTipText     =   "DeleteAll"
            ImageKey        =   "DeleteAll"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Initialize"
            Object.ToolTipText     =   "Initialize"
            ImageKey        =   "Initialize"
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":235A
         Left            =   5130
         List            =   "frmMain.frx":2364
         TabIndex        =   10
         Text            =   "Internal Flash Memory"
         Top             =   45
         Width           =   1680
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3440
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Mem Total:"
            TextSave        =   "Mem Total:"
            Key             =   "MemTotal"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Mem Free:"
            TextSave        =   "Mem Free:"
            Key             =   "MemFree"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Mem Used:"
            TextSave        =   "Mem Used:"
            Key             =   "MemUsed"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "5/15/01"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "1:49 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6060
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8520
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2396
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24A8
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25BA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26CC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27DE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28F0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A02
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B14
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C26
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D38
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E4A
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F5C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":306E
            Key             =   "Sync"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":338A
            Key             =   "Trash"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36AE
            Key             =   "Upload"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CE6
            Key             =   "Initialize"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4002
            Key             =   "RandomFolder"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":431E
            Key             =   "Playlist"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":463A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4956
            Key             =   "reload"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C72
            Key             =   "RandomPlaylist"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52AA
            Key             =   "DeleteAll"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58EA
            Key             =   "fixed"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C0E
            Key             =   "ram"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F32
            Key             =   "remote"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6256
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":657A
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":689E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BC2
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EE6
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":720A
            Key             =   "OnePlaylist"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7526
            Key             =   "OneFolder"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7842
            Key             =   "OneFolderCurrent"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B5E
            Key             =   "OnePlaylistCurrent"
         EndProperty
      EndProperty
   End
   Begin SSActiveTreeView.SSTree tvTreeView 
      Height          =   4815
      Left            =   10
      TabIndex        =   8
      Top             =   720
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   8493
      _Version        =   65536
      BackColor       =   0
      ForeColor       =   65280
      ImagesMaskColor =   14737632
      LabelEdit       =   1
      LineStyle       =   1
      LineType        =   2
      PictureAlignment=   0
      Indentation     =   345
      ImageCount      =   13
      OLEDragMode     =   1
      OLEDropMode     =   1
      Sorted          =   1
      ImagesUseMask   =   -1  'True
      UseImageList    =   0   'False
      HasFont         =   0   'False
      HasMouseIcon    =   0   'False
      HasPictureBackground=   0   'False
      ImageList       =   "<None>"
      Image(0).Index  =   0
      Image(0).Picture=   "frmMain.frx":7E7A
      Image(1).Index  =   1
      Image(1).Picture=   "frmMain.frx":7F76
      Image(1).Key    =   "fixed"
      Image(2).Index  =   2
      Image(2).Picture=   "frmMain.frx":8292
      Image(2).Key    =   "cd"
      Image(3).Index  =   3
      Image(3).Picture=   "frmMain.frx":85AE
      Image(3).Key    =   "ram"
      Image(4).Index  =   4
      Image(4).Picture=   "frmMain.frx":88CA
      Image(4).Key    =   "remove"
      Image(5).Index  =   5
      Image(5).Picture=   "frmMain.frx":8BE6
      Image(5).Key    =   "remote"
      Image(6).Index  =   6
      Image(6).Picture=   "frmMain.frx":8F02
      Image(6).Key    =   "mp3file"
      Image(7).Index  =   7
      Image(7).Picture=   "frmMain.frx":921E
      Image(7).Key    =   "open"
      Image(8).Index  =   8
      Image(8).Picture=   "frmMain.frx":953A
      Image(8).Key    =   "closed"
      Image(9).Index  =   9
      Image(9).Picture=   "frmMain.frx":9856
      Image(9).Key    =   "history"
      Image(10).Index =   10
      Image(10).Picture=   "frmMain.frx":9B72
      Image(10).Key   =   "computer"
      Image(11).Index =   11
      Image(11).Picture=   "frmMain.frx":9E8E
      Image(11).Key   =   "favorites"
      Image(12).Index =   12
      Image(12).Picture=   "frmMain.frx":A1AA
      Image(12).Key   =   "playlist"
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnu_ActionsSync 
         Caption         =   "&Sync"
      End
      Begin VB.Menu mnu_ActionsUpload 
         Caption         =   "&Upload"
      End
      Begin VB.Menu mnu_ActionsInit 
         Caption         =   "&Initialize"
      End
      Begin VB.Menu mnu_ActionsRefresh 
         Caption         =   "&Refresh Dir"
      End
      Begin VB.Menu mnu_ActionsDeleteSel 
         Caption         =   "&Delete Selected"
      End
      Begin VB.Menu mnu_ActionsDeleteAll 
         Caption         =   "D&eleted All"
      End
      Begin VB.Menu mnu_ActionsRanFolder 
         Caption         =   "R&andom from Folder"
      End
      Begin VB.Menu mnu_ActionsRanPlaylist 
         Caption         =   "Ra&ndom from Playlist"
      End
      Begin VB.Menu mnu_ActionsOneTouchFolder 
         Caption         =   "One &Touch Folder"
      End
      Begin VB.Menu mnu_ActionsOneTouchPlaylist 
         Caption         =   "&One Touch Playlist"
      End
      Begin VB.Menu mnu_ActionsOneFolderCurrent 
         Caption         =   "One Touch Folder Current"
      End
      Begin VB.Menu mnu_ActionsOnePlaylistCurrent 
         Caption         =   "One Touch Playlist Current"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnu_HelpRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is a variable to determine how many days to hold history
Public intKeepHistory As Integer

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Dim itmX As ListItem
' Declare the variables to use the browse for folder API calls
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib _
        "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
        "shell32" (ByVal pidList As Long, ByVal lpbuffer _
        As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
        Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
        lpString2 As String) As Long

Private Type BrowseInfo
        hwndOwner As Long
        pIDLRoot As Long
        pszDisplayName As Long
        lpszTitle As Long
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type


Private Type ITEMIDLIST
    mkid As Long
End Type

Private Declare Function SHGetSpecialFolderLocation _
    Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As SHFolders, _
    ppidl As ITEMIDLIST) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal pv As Long)

Public Enum SHFolders
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D '// DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum

Private Sub Combo1_Click()
    Select Case Combo1
        Case "Internal Flash Memory"
            lvListView.ListItems.Clear
            RioInterface1.UseExternal = False
            RioInterface1.GetDirectory
            UpdateCounters
        Case "External Flash Memory"
            lvListView.ListItems.Clear
            RioInterface1.UseExternal = True
            RioInterface1.GetDirectory
            UpdateCounters
    End Select
End Sub

Private Sub Form_Load()
    Dim btnX As Button

   
    Set btnX = tbToolBar.Buttons.Add(1)
    btnX.Style = tbrSeparator
    Set btnX = tbToolBar.Buttons.Add(1)
    btnX.Style = tbrPlaceholder
    btnX.Key = "combo"
    btnX.Width = 2000
With Combo1
        .ZOrder 0
        .Width = tbToolBar.Buttons("combo").Width
        .Top = tbToolBar.Buttons("combo").Top
        .Left = tbToolBar.Buttons("combo").Left + 50
    End With

Dim ssNode As ssNode
'debug.print "Load:" & Time
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    lvListView.ColumnHeaders.Add , , "Title", lvListView.Width / 5
    lvListView.ColumnHeaders.Add , , "Size", lvListView.Width / 5
    lvListView.ColumnHeaders.Add , , "Bit Rate", lvListView.Width / 5
    lvListView.ColumnHeaders.Add , , "Sample", lvListView.Width / 5
    lvListView.ColumnHeaders.Add , , "Time", lvListView.Width / 5
    lvListView.ColumnHeaders(1).Width = GetSetting(App.Title, "Settings", "ListTitle", lvListView.Width / 5)
    lvListView.ColumnHeaders(2).Width = GetSetting(App.Title, "Settings", "ListSize", lvListView.Width / 5)
    lvListView.ColumnHeaders(3).Width = GetSetting(App.Title, "Settings", "ListBitRate", lvListView.Width / 5)
    lvListView.ColumnHeaders(4).Width = GetSetting(App.Title, "Settings", "ListSample", lvListView.Width / 5)
    lvListView.ColumnHeaders(5).Width = GetSetting(App.Title, "Settings", "ListTime", lvListView.Width / 5)
    intKeepHistory = GetSetting(App.Title, "Settings", "KeepHistory", 4)
    If GetSetting(App.Title, "Settings", "DefaultPort", 378) = "LPT1" Then
        RioInterface1.DevicePort = 378
    Else
        RioInterface1.DevicePort = 278
    End If
    
Debug.Print GetSetting(App.Title, "Settings", "DefaultPort", "LPT1")
tvTreeView.ImagesUseMask = True
tvTreeView.UseImageList = False

    
 LoadDriveTree
 'LoadFavorites
 LoadHistory


   ' Set View property to Report.
   lvListView.View = lvwReport

    RioInterface1.GetDirectory
    sbStatusBar.Panels(2) = "Mem Total:" & RioInterface1.MemoryTotal
    sbStatusBar.Panels(3) = "Mem Used:" & RioInterface1.MemoryUsed
    sbStatusBar.Panels(4) = "Mem Free:" & RioInterface1.MemoryFree
    On Error Resume Next
    ProgressBar1.Percent = (RioInterface1.MemoryUsed / RioInterface1.MemoryTotal) * 100

    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "ListTitle", lvListView.ColumnHeaders(1).Width
        SaveSetting App.Title, "Settings", "ListSize", lvListView.ColumnHeaders(2).Width
        SaveSetting App.Title, "Settings", "ListBitRate", lvListView.ColumnHeaders(3).Width
        SaveSetting App.Title, "Settings", "ListSample", lvListView.ColumnHeaders(4).Width
        SaveSetting App.Title, "Settings", "ListTime", lvListView.ColumnHeaders(5).Width
        
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    If Not Me.WindowState = vbMinimized Then SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'set the width
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    tvTreeView.Width = X
    imgSplitter.Left = X
    lvListView.Left = X + 40
    ProgressBar1.Left = lvListView.Left
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40


    'set the top
  

    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If

    lvListView.Top = tvTreeView.Top
    

    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    lvListView.Height = tvTreeView.Height - 260

    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
    ProgressBar1.Top = lvListView.Top + lvListView.Height
    ProgressBar1.Width = Me.Width - (tvTreeView.Width + 170)
        
    
End Sub


Private Sub lvListView_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then mnuFileDelete_Click
End Sub

Private Sub lvListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Counter As Integer
    Dim itnX As ListItem
    Dim returndata() As String
    Dim strReturn As String
    Dim Size As Long
   ' Dim sdir As String
    Dim counter2 As Integer
    Dim ssNode As ssNode, tmpNode As ssNode
    Dim i As Integer
    
    If Data.GetFormat(vbCFText) Then
        If Left(Data.GetFormat(vbCFText), 7) = "History" Then Exit Sub
        'debug.print Data.GetData(vbCFText)
        
        If tvTreeView.SelectedItem.Text = Data.GetData(vbCFText) Then
            If Right(LCase(tvTreeView.SelectedItem.Key), 4) = ".mp3" Then
                Set itmX = lvListView.ListItems.Add(, "new" & lvListView.ListItems.count & "-" & tvTreeView.SelectedItem.Key, Data.GetData(vbCFText), 2, 2)
                strReturn = Getmp3data(tvTreeView.SelectedItem.Key)
                returndata() = Split(strReturn, vbTab)
                itmX.SubItems(1) = returndata(0)
                itmX.SubItems(2) = returndata(1)
                itmX.SubItems(3) = returndata(2)
                itmX.SubItems(4) = returndata(3)
            UpdateCounters
            Else
                tvTreeView.SelectedItem.Expanded = True 'Needed if using OnDemand LoadStyle
            
                If tvTreeView.SelectedItem.Children > 0 Then
                    Set ssNode = tvTreeView.SelectedItem.Child
                    Set tmpNode = tvTreeView.SelectedItem
                    For i = 1 To tmpNode.Children
                        
                        tvTreeView.SelectedNodes.Add ssNode
            
                        
                        If i < tmpNode.Children Then
                            Set ssNode = ssNode.Next
                        End If
                        If Right(LCase(ssNode.Text), 4) = ".mp3" Then
                            Set itmX = lvListView.ListItems.Add(, "new" & lvListView.ListItems.count & "-" & ssNode.Key, ssNode.Text, 2, 2)
                            strReturn = Getmp3data(ssNode.Key)
                            returndata() = Split(strReturn, vbTab)
                            itmX.SubItems(1) = returndata(0)
                            itmX.SubItems(2) = returndata(1)
                            itmX.SubItems(3) = returndata(2)
                            itmX.SubItems(4) = returndata(3)
                        End If
                    Next i
                End If

                
            End If

        End If
    End If
    
    If Data.GetFormat(vbCFFiles) Then
        
        Counter = 1
        
        Do While Not Counter = Data.Files.count + 1
            If Right(LCase(Data.Files(Counter)), 4) = ".mp3" Then
                Set itmX = lvListView.ListItems.Add(, "new" & lvListView.ListItems.count & "-" & Data.Files(Counter), ExtractFileName(Data.Files(Counter)), 2, 2)
                strReturn = Getmp3data(Data.Files(Counter))
                returndata() = Split(strReturn, vbTab)
                itmX.SubItems(1) = returndata(0)
                itmX.SubItems(2) = returndata(1)
                itmX.SubItems(3) = returndata(2)
                itmX.SubItems(4) = returndata(3)
            End If
            
            Counter = Counter + 1
        Loop

UpdateCounters
 End If
    
End Sub


Private Sub mnu_ActionsDeleteAll_Click()
    Dim Counter As Integer
    Dim counter2 As Integer
    If lvListView.SelectedItem Is Nothing Then Exit Sub
    counter2 = 1
    
    Do While lvListView.ListItems.count + 1 > counter2

    
        If Left(lvListView.ListItems(counter2).Key, 3) = "new" Then
            lvListView.ListItems.Remove lvListView.ListItems(counter2).Index
            counter2 = counter2 - 1
        Else

            If Not RioInterface1.FileRemove(lvListView.ListItems(counter2)) Then
                sbStatusBar.Panels(1) = "Delete Failed"
            Else
            lvListView.ListItems.Remove lvListView.ListItems(counter2).Index
            counter2 = counter2 - 1
                sbStatusBar.Panels(1) = "Complete"
            End If
        End If
            
    counter2 = counter2 + 1
    Loop
    Counter = 1
 
 RioInterface1.GetHeader
  UpdateCounters
 
End Sub

Private Sub mnu_ActionsDeleteSel_Click()
    If RioInterface1.FileRemove(lvListView.SelectedItem) Then lvListView.ListItems.Remove (lvListView.SelectedItem.Index)
    UpdateCounters
End Sub

Private Sub mnu_ActionsInit_Click()
    lvListView.ListItems.Clear
    If RioInterface1.RioInitialize Then RioInterface1.GetDirectory
End Sub

Private Sub mnu_ActionsOneFolderCurrent_Click()
    Dim Selection As String
    Selection = BrowseFolder()
    If Not Len(Selection) > 0 Then Exit Sub
    mnu_ActionsDeleteAll_Click
        If RioInterface1.GetDirectory Then UpdateCounters
        RandomFolder Selection
    
    mnu_ActionsSync_Click
    If RioInterface1.GetDirectory Then UpdateCounters
    
End Sub

Private Sub mnu_ActionsOnePlaylistCurrent_Click()
  Dim strFile As String
    Dim strpath As String
    
        dlgCommonDialog.Filter = "Playlist(*.m3u)|*.m3u"
        dlgCommonDialog.ShowOpen
        strpath = dlgCommonDialog.FILENAME
        strFile = dlgCommonDialog.FILENAME
    
        mnu_ActionsDeleteAll_Click
        If RioInterface1.GetDirectory Then UpdateCounters

    mnu_ActionsSync_Click
    UpdateCounters
    
End Sub

Private Sub mnu_ActionsOneTouchFolder_Click()
    Dim Selection As String
    Selection = BrowseFolder()
    If Not Len(Selection) > 0 Then Exit Sub
    
    Combo1.Text = "Internal Flash Memory"
    Combo1_Click
        mnu_ActionsDeleteAll_Click
        If RioInterface1.GetDirectory Then UpdateCounters
        RandomFolder Selection

    mnu_ActionsSync_Click
    
    Combo1.Text = "External Flash Memory"
    Combo1_Click
        mnu_ActionsDeleteAll_Click
        If RioInterface1.GetDirectory Then UpdateCounters
        RandomFolder Selection

    mnu_ActionsSync_Click
    If RioInterface1.GetDirectory Then UpdateCounters
End Sub

Private Sub mnu_ActionsOneTouchPlaylist_Click()
  Dim strFile As String
    Dim strpath As String
    
        dlgCommonDialog.Filter = "Playlist(*.m3u)|*.m3u"
        dlgCommonDialog.ShowOpen
        strpath = dlgCommonDialog.FILENAME
        strFile = dlgCommonDialog.FILENAME
    If strpath = "" Or strFile = "" Then Exit Sub
    
    Combo1.Text = "Internal Flash Memory"
    Combo1_Click
        mnu_ActionsDeleteAll_Click
        If RioInterface1.GetDirectory Then UpdateCounters
        RandomPlaylist strpath, strFile

    mnu_ActionsSync_Click
    UpdateCounters
    
    Combo1.Text = "External Flash Memory"
    Combo1_Click
    mnu_ActionsDeleteAll_Click
        If RioInterface1.GetDirectory Then UpdateCounters
        RandomPlaylist strpath, strFile
    mnu_ActionsDeleteAll_Click
    mnu_ActionsSync_Click
    UpdateCounters
End Sub

Private Sub mnu_ActionsRanFolder_Click()
    Dim Selection As String
    Selection = BrowseFolder()
    If Len(Selection) > 0 Then RandomFolder (Selection)
End Sub

Private Sub mnu_ActionsRanPlaylist_Click()
    RandomPlaylist
    UpdateCounters
    
    
End Sub

Private Sub mnu_ActionsRefresh_Click()
    lvListView.ListItems.Clear
    If RioInterface1.GetDirectory Then UpdateCounters

    
End Sub

Private Sub mnu_ActionsSync_Click()
    Dim Counter As Integer
    Dim TotalFiles As Integer
    Dim CurrentFile As Integer
    
    TotalFiles = 0
    CurrentFile = 0
    
    For Counter = 1 To lvListView.ListItems.count
        If Left(lvListView.ListItems(Counter).Key, 3) = "new" Then TotalFiles = TotalFiles + 1
    Next Counter
    Close #30
    Open App.Path & "\" & "history.dat" For Append As #30
    
    If TotalFiles = 0 Then Exit Sub
    frmStatus.Show
    frmStatus.ProgressBarFile.Percent = 0
    frmStatus.ProgressBarTotal.Percent = 0

    If ProgressBar1.Percent > 100 Then
        MsgBox "Not enough space to handle list, please delete some files.", vbCritical, "Warning"
        Exit Sub
    Else
        For Counter = 1 To lvListView.ListItems.count
            If Left(lvListView.ListItems(Counter).Key, 3) = "new" Then
                sbStatusBar.Panels(1).Text = "Downloading:" & lvListView.ListItems(Counter).Text
                frmStatus.Label1(0) = "Downloading:" & lvListView.ListItems(Counter).Text
                frmStatus.Label1(0).Width = frmStatus.ProgressBarFile.Width
                If Not RioInterface1.FileSend(Mid$(lvListView.ListItems(Counter).Key, InStr(1, lvListView.ListItems(Counter).Key, "-") + 1)) Then
                    sbStatusBar.Panels(1).Text = "Failed to upload file"
                Else
                    Write #30, lvListView.ListItems(Counter); Date
                    Set ssNode = tvTreeView.Nodes.Add("History", ssatChild, "History" & "-" & lvListView.ListItems(Counter).Key, lvListView.ListItems(Counter).Text & vbTab & Date, "mp3file", "mp3file")
                    CurrentFile = CurrentFile + 1
                    frmStatus.ProgressBarTotal.Percent = (CurrentFile / TotalFiles) * 100
                    lvListView.ListItems(Counter).SmallIcon = 3
                    lvListView.ListItems(Counter).Icon = 3
                    lvListView.ListItems(Counter).Key = Str(Counter) & lvListView.ListItems(Counter).Text
                End If
            End If
        Next Counter
                
    End If
    
    Close #30
    frmStatus.Hide
    RioInterface1.GetDirectory
    UpdateCounters
End Sub

Private Sub mnu_ActionsUpload_Click()
    Dim Counter As Integer
    Dim TotalFiles As Integer
    Dim CurrentFile As Integer
    
    TotalFiles = 0
    CurrentFile = 0
    
    
    For Counter = 1 To lvListView.ListItems.count
        If Not Left(lvListView.ListItems(Counter).Key, 3) = "new" And lvListView.ListItems(Counter).Selected Then TotalFiles = TotalFiles + 1
    Next Counter
    
    
    If TotalFiles = 0 Then Exit Sub
    frmStatus.Show
    frmStatus.ProgressBarFile.Percent = 0
    frmStatus.ProgressBarTotal.Percent = 0

    
        For Counter = 1 To lvListView.ListItems.count
            If Not Left(lvListView.ListItems(Counter).Key, 3) = "new" And lvListView.ListItems(Counter).Selected Then
                sbStatusBar.Panels(1).Text = "Uploading:" & lvListView.ListItems(Counter).Text
                frmStatus.Label1(0) = "Uploading:" & lvListView.ListItems(Counter).Text
                
                If Not RioInterface1.FileRetrieve(App.Path & "\" & lvListView.ListItems(Counter)) Then
                    sbStatusBar.Panels(1).Text = "Failed to upload file"
                Else
                    CurrentFile = CurrentFile + 1
                    frmStatus.ProgressBarTotal.Percent = (CurrentFile / TotalFiles) * 100
                    
                End If
            End If
        Next Counter
                
    
    frmStatus.Hide
    UpdateCounters
    
    If Len(lvListView.SelectedItem) > 0 Then
                If Not RioInterface1.FileRetrieve("c:\" & lvListView.SelectedItem) Then
                    sbStatusBar.Panels(1) = "An error occured during transfer"
                Else
                    sbStatusBar.Panels(1) = "Complete"
                    ProgressBar1.Percent = (RioInterface1.MemoryUsed / RioInterface1.MemoryTotal) * 100
                End If
            End If
End Sub

Private Sub mnu_HelpRegister_Click()
    frmLogin.Show
End Sub



Private Sub RioInterface1_DirectoryEntry(ByVal intFileNum As Integer, ByVal lngFileSize As Long, ByVal lngBitRate As Long, ByVal lngSampleFreq As Long, ByVal strTime As String, ByVal strFileName As String)
    
If intFileNum = 1 Then lvListView.ListItems.Clear
    Set itmX = lvListView.ListItems.Add(, "Entry" & Str(intFileNum), strFileName, 3, 3)
    itmX.SubItems(1) = Format(Str(lngFileSize), "###,###,###")
    itmX.SubItems(2) = Str(lngBitRate)
    itmX.SubItems(3) = Str(lngSampleFreq)
    itmX.SubItems(4) = Format(strTime, "m/d/yy h:mm AMPM")
End Sub

Private Sub RioInterface1_ReturnError(ByVal strError As String)
    sbStatusBar.Panels(1).Text = strError
End Sub

Private Sub RioInterface1_RioStatus(ByVal strStatus As String)
    sbStatusBar.Panels(1).Text = strStatus
    DoEvents
End Sub

Private Sub RioInterface1_TransferProgress(ByVal Position As Integer, ByVal Total As Integer)
frmStatus.ProgressBarFile.Percent = (Position / Total) * 100
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Delete"
            mnu_ActionsDeleteSel_Click
        Case "Reload"
            mnu_ActionsRefresh_Click
        Case "Sync"
            mnu_ActionsSync_Click
        Case "Upload"
            mnu_ActionsUpload_Click
        Case "Initialize"
           mnu_ActionsInit_Click
        Case "RandomPlaylist"
            mnu_ActionsRanPlaylist_Click
        Case "RandomFolder"
            mnu_ActionsRanFolder_Click
        Case "DeleteAll"
            mnu_ActionsDeleteAll_Click
        Case "OneFolder"
            mnu_ActionsOneTouchFolder_Click
        Case "OnePlaylist"
            mnu_ActionsOneTouchPlaylist_Click
            
        Case "OneFolderCurrent"
            mnu_ActionsOneFolderCurrent_Click
        Case "OnePlaylistCurrent"
            mnu_ActionsOnePlaylistCurrent_Click
            
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer

    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileDelete_Click()
    Dim Counter As Integer
    Dim counter2 As Integer
    If lvListView.SelectedItem Is Nothing Then Exit Sub
    counter2 = 1
    
    Do While lvListView.ListItems.count + 1 > counter2
    If lvListView.ListItems(counter2).Selected Then
    
        If Left(lvListView.ListItems(counter2).Key, 3) = "new" Then
            'debug.print lvListView.ListItems(counter2).Text
            lvListView.ListItems.Remove lvListView.ListItems(counter2).Index
            counter2 = counter2 - 1
        Else

            If Not RioInterface1.FileRemove(lvListView.ListItems(counter2)) Then
                sbStatusBar.Panels(1) = "Delete Failed"
                'debug.print "delete failed:" & lvListView.ListItems(counter2)
            Else
                'debug.print lvListView.ListItems(counter2).Text
                lvListView.ListItems.Remove lvListView.ListItems(counter2).Index
                counter2 = counter2 - 1
                sbStatusBar.Panels(1) = "Complete"
            End If
        End If
            
        Else
               End If
    counter2 = counter2 + 1
    Loop
    Counter = 1
 
 RioInterface1.GetHeader
  UpdateCounters
 
End Sub




Public Function ExtractFileName(strpath As String) As String
    Dim tmp As String
    ' StrReverse is only working in VB6
    strpath = StrReverse(strpath)
    tmp = StrReverse(Left(strpath, InStr(strpath, "\") - 1))
    ExtractFileName = tmp
    strpath = StrReverse(strpath)
End Function


Public Sub UpdateCounters()
    Dim Size As Long
    Dim Counter As Integer
    Counter = 1
    
    'Calculate Space
            Do While Not Counter = lvListView.ListItems.count + 1
               Size = Size + CLng(lvListView.ListItems.Item(Counter).SubItems(1))
               Counter = Counter + 1
            Loop
            If RioInterface1.MemoryTotal = 0 And RioInterface1.MemoryUsed = 0 Then
                ProgressBar1.Percent = ((Size / 1000) / 32000) * 100
                If ProgressBar1.Percent > 100 Then ProgressBar1.BarColor = vbRed Else ProgressBar1.BarColor = vbGreen
                    sbStatusBar.Panels(1) = "Device Not Connected using defaults"
                    sbStatusBar.Panels(2) = "Mem Total:" & Int(32000)
                    sbStatusBar.Panels(3) = "Mem Used:" & Int(Size / 1000)
                    sbStatusBar.Panels(4) = "Mem Free:" & Int(32000 - (Size / 1000))
            Else
                ProgressBar1.Percent = ((Size / 1000) / RioInterface1.MemoryTotal) * 100
                If ProgressBar1.Percent > 100 Then ProgressBar1.BarColor = vbRed Else ProgressBar1.BarColor = vbGreen
                
                    sbStatusBar.Panels(2) = "Mem Total:" & Int(RioInterface1.MemoryTotal)
                    sbStatusBar.Panels(3) = "Mem Used:" & Int(Size / 1000)
                    sbStatusBar.Panels(4) = "Mem Free:" & Int(RioInterface1.MemoryTotal - (Size / 1000))
            End If
End Sub

Public Sub RandomFolder(Directory As String)
Dim sdir As String
Dim List() As String
Dim count As Long
Dim X As Integer
Dim flag As Boolean
Dim fs As Object
Dim A() ' Sets the maximum number to pick
Dim B()
Dim returndata() As String
Dim Message, Message_Style, Message_Title, Response

On Error GoTo HandleError

Set fs = CreateObject("Scripting.FileSystemObject")

'if the directory does not contain the last '\' then add it
If Not Right(Directory, 1) = "\" Then Directory = Directory & "\"

'get the directory
sdir = Dir(Directory & "*.mp3")
count = 0

'get the count of files in the directory
Do While sdir <> ""
    count = count + 1
    sdir = Dir()
Loop

'if the directory is empty exit the routine
If count = 0 Then Exit Sub


'check to see what the size of the current file list is
If lvListView.ListItems.count > 0 Then
    Counter = 0
    Do While Not Counter = lvListView.ListItems.count + 1
       Size = Size + CLng(lvListView.ListItems.Item(Counter).SubItems(1)) 'this will allow the user to add random to fill the rest of the rio
       Counter = Counter + 1
    Loop
    Size = Size / 1000
Else
    Size = 0
End If
'set the size of the arrays to the ammount of entries in the directory
ReDim List(count - 1)
ReDim A(count - 1)
ReDim B(count - 1) ' Will be the list of new numbers (same as DIM above)

'load the array with the names of the files
sdir = Dir(Directory & "*.mp3")
count = 0
Do While sdir <> ""
    List(count) = sdir
    count = count + 1
    sdir = Dir()
Loop
    
'*******************************
'*  begin random number generation
'*  selection
'*******************************
        'Set the original array
        maxnumber = count - 1 ' Must equal the Dim above


        For seq = 0 To maxnumber
            A(seq) = seq
        Next seq
        'Main Loop (mix em all up)
        StartTime = Timer
        Randomize (Timer)


        For MainLoop = maxnumber To 0 Step -1
            ChosenNumber = Int(MainLoop * Rnd)
            B(maxnumber - MainLoop) = A(ChosenNumber)
            A(ChosenNumber) = A(MainLoop)
        Next MainLoop
        EndTime = Timer
        TotalTime = EndTime - StartTime
        
'!!!!!!! Debug Block
        Message = "The sequence of " + Format(maxnumber, "#,###,###,###") + " numbers has been" + Chr$(10)
        Message = Message + "mixed up in a total of " + Format(TotalTime, "##.######") + " seconds!"
        Message_Style = vbInformationOnly + vbInformation + vbDefaultButton2
        Message_Title = "Sequence Generated"
        
        'generate list

        For X = 0 To maxnumber
            If Size > 31000 Then Exit For
                If Size + (FileLen(Directory & List(B(X))) / 1000) > 32000 Then
                    
                Else
                      If Not ExistsInHistory(List(B(X))) Then
                          Set itmX = lvListView.ListItems.Add(, "new" & lvListView.ListItems.count & "-" & Directory & List(B(X)), List(B(X)), 2, 2)
                            strReturn = Getmp3data(Directory & List(B(X)))
                            returndata() = Split(strReturn, vbTab)
                            itmX.SubItems(1) = returndata(0)
                            itmX.SubItems(2) = returndata(1)
                            itmX.SubItems(3) = returndata(2)
                            itmX.SubItems(4) = returndata(3)
                        Size = Size + (FileLen(Directory & List(B(X))) / 1000)
                    End If
                End If
        Next X
UpdateCounters

Exit Sub
HandleError:
    'debug.print Err.Description
Resume Next
End Sub


Public Function BrowseFolder() As String

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
            .hwndOwner = Me.hwnd ' Owner Form
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    BrowseFolder = sBuffer
End Function


Public Sub RandomPlaylist(Optional dfltpath As String, Optional dfltFile As String)
    Dim strpath     As String
    Dim strFile     As String
    Dim strFiles()  As String
    Dim lngTotal    As Long
    Dim Size        As Long
    Dim count       As Integer
    Dim Counter     As Integer
    Dim fs As Object
    Dim A() ' Sets the maximum number to pick
    Dim B()
    Dim returndata() As String
    Dim Message, Message_Style, Message_Title, Response
    Dim List()       As String
    Dim strLine      As String
    Dim strDesktopDirectory As String
        
    On Error GoTo HandleError
    
    If ProgressBar1.Percent > 100 Then
        MsgBox "Remove songs from list to create room", vbCritical
        Exit Sub
    End If
    If Len(dfltFile) > 0 And Len(dfltpath) > 0 Then
        strpath = dfltpath
        strFile = dfltFile
    Else
        dlgCommonDialog.Filter = "Playlist(*.m3u)|*.m3u"
        dlgCommonDialog.ShowOpen
        strpath = dlgCommonDialog.FILENAME
        strFile = dlgCommonDialog.FILENAME
    End If
    If strFile = "" Then Exit Sub
'check to see what the size of the current file list is
 If lvListView.ListItems.count > 0 Then
    Counter = 1
    Do While Not Counter = lvListView.ListItems.count + 1
       Size = Size + CLng(lvListView.ListItems.Item(Counter).SubItems(1)) 'this will allow the user to add random to fill the rest of the rio
       Counter = Counter + 1
    Loop
    Size = Size \ 1000
Else
    Size = 0
End If
    
    lngTotal = 0
    Open strpath For Input As #1
    Line Input #1, strLine
    
    If Trim$(strLine) = "#EXTM3U" Then
        Do While Not EOF(1)
            If Left(strLine, 8) = "#EXTINF:" Then lngTotal = lngTotal + 1
            Line Input #1, strLine
        Loop
        Close #1
        
        If lngTotal = 0 Then Exit Sub 'if there are no files in playlist file
        
        ReDim strFiles(lngTotal - 1, 2)
        ReDim List(lngTotal - 1)
        ReDim A(lngTotal - 1)
        ReDim B(lngTotal - 1) ' Will be the list of new numbers (same as DIM above)
        Open strpath For Input As #1
        Line Input #1, strLine
        count = 0
        Do While Not EOF(1)
            If Left(strLine, 7) = "#EXTINF" Then
                'remove leader
                strLine = StrReverse(strLine)
                strLine = Left(strLine, Len(strLine) - 8)
                strLine = StrReverse(strLine)
                'get file length and title
                strlength = Left(strLine, InStr(1, strLine, ",") - 1)
                strtitle = Mid$(strLine, InStr(1, strLine, ",") + 1)
                'If Len(strTitle) < 40 Then strTitle = strTitle & Space(40 - Len(strTitle))
                lngTime = CLng(Trim$(strlength))
                'get title
                strlength = Trim$(Str(Int(lngTime \ 60 Mod 60))) & ":" & Format(Trim$(Str(Int(lngTime Mod 60))), "0#")
                Line Input #1, strLine
                
                'Set ssNode = tvtreeview.Nodes.Add("Playlist", ssatChild, strLine, strLength & vbTab & strTitle, 3, 3)
                
                'add to array
                strFiles(count, 1) = strtitle
                strFiles(count, 2) = strLine
                
                count = count + 1
            End If
            If Not EOF(1) Then Line Input #1, strLine
        Loop
    
    Else
        Close #1
        
        strDesktopDirectory = FolderLocation(16)
        
        
        Open strpath For Input As #1
            Do While Not EOF(1)
                If Not Left(strLine, 1) = "#" Then lngTotal = lngTotal + 1
                Line Input #1, strLine
            Loop
        Close #1
        
        If lngTotal = 0 Then Exit Sub 'if there are no files in playlist file
        
        ReDim strFiles(lngTotal - 1, 2)
        ReDim List(lngTotal - 1)
        ReDim A(lngTotal - 1)
        ReDim B(lngTotal - 1) ' Will be the list of new numbers (same as DIM above)

        Open strpath For Input As #1
        Do While Not EOF(1)
            lngTotal = lngTotal + 1
            Line Input #1, strLine
            If Mid(Trim$(strLine), 2, 1) = ":" Then
                If strLine <> "" Then
                    List1.AddItem strLine
                    strFiles(count, 1) = ExtractFileName(strLine)
                    strFiles(count, 2) = strLine
                    count = count + 1
                End If
            Else
                If strLine <> "" Then
                    If Not Left(strLine, 1) = "\" Then
                        List1.AddItem strDesktopDirectory & "\" & Trim$(strLine)
                        strFiles(count, 1) = ExtractFileName(strLine)
                        strFiles(count, 2) = strDesktopDirectory & "\" & strLine
                        count = count + 1
                    End If
                End If
            End If
        Loop
    End If
        Close #1

'************
'*  begin random number generation
'*  selection
'*******************************
        'Set the original array
        maxnumber = count - 1 ' Must equal the Dim above


        For seq = 0 To maxnumber
            A(seq) = seq
        Next seq
        'Main Loop (mix em all up)
        StartTime = Timer
        Randomize (Timer)


        For MainLoop = maxnumber To 0 Step -1
            ChosenNumber = Int(MainLoop * Rnd)
            B(maxnumber - MainLoop) = A(ChosenNumber)
            A(ChosenNumber) = A(MainLoop)
        Next MainLoop
        EndTime = Timer
        TotalTime = EndTime - StartTime
        
'!!!!!!! Debug Block
        Message = "The sequence of " + Format(maxnumber, "#,###,###,###") + " numbers has been" + Chr$(10)
        Message = Message + "mixed up in a total of " + Format(TotalTime, "##.######") + " seconds!"
        Message_Style = vbInformationOnly + vbInformation + vbDefaultButton2
        Message_Title = "Sequence Generated"
        'debug.print Message
'!!!!!!!  End Debug Block

        'generate list
        'Size = 0
        For X = 0 To maxnumber
            If Size > 31000 Then Exit For
                If Size + (FileLen(strFiles(B(X), 2)) / 1000) > 32000 Then
                    
                Else
                    'debug.print strFiles(B(X), 1)
                        If Not ExistsInHistory(ExtractFileName(strFiles(B(X), 2))) Then
                          Set itmX = lvListView.ListItems.Add(, "new" & lvListView.ListItems.count & "-" & Directory & strFiles(B(X), 2), strFiles(B(X), 1), 2, 2)
                            strReturn = Getmp3data(strFiles(B(X), 2))
                            returndata() = Split(strReturn, vbTab)
                            itmX.SubItems(1) = returndata(0)
                            itmX.SubItems(2) = returndata(1)
                            itmX.SubItems(3) = returndata(2)
                            itmX.SubItems(4) = returndata(3)
                            Size = Size + (FileLen(strFiles(B(X), 2)) / 1000)
                        End If
                End If
                        
        Next X
'debug.print Size
Close #1

UpdateCounters

Exit Sub
HandleError:
    'debug.print Err.Description
Resume Next


Close #1
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Sub tvTreeView_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    Dim ssPreviousNode As ssNode
    Dim ssExpandedNode As ssNode
    
    ' It is possible that at some point we will hit a node that is equal to nothing.
    On Error GoTo NoMoreNodes
    
    If indrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        Set tvTreeView.DropHighlight = tvTreeView.HitTest(X, Y)
        
        ' Scroll up at the bottom if over the last node
        If tvTreeView.GetVisibleNode(tvTreeView.GetVisibleCount) = tvTreeView.HitTest(X, Y) Then
            Set tvTreeView.TopNode = tvTreeView.GetVisibleNode(2)
        End If
                
        If tvTreeView.TopNode = tvTreeView.HitTest(X, Y) Then
        
            If tvTreeView.TopNode.Parent Is Nothing And tvTreeView.TopNode.Previous Is Nothing Then
                Exit Sub
            End If
            
            ' This is the node the will ultimately be set.
            Set ssPreviousNode = tvTreeView.TopNode.Previous
            
            ' Loop until we have found the node
            Do While True
                ' The first sibling node will return nothing for a previous node
                If ssPreviousNode Is Nothing Then
                    ' Set the node to the parent and exit
                    Set ssPreviousNode = tvTreeView.TopNode.Parent
                    Exit Do
                ' If the previous node is expanded,
                ' get the last sibling and see if that is expanded as well.
                ElseIf ssPreviousNode.Expanded Then
                    Set ssExpandedNode = ssPreviousNode.Child
                    Set ssPreviousNode = ssExpandedNode.LastSibling
                Else
                    Exit Do
                End If
            Loop
        
            Set tvTreeView.TopNode = ssPreviousNode
        End If
                
    End If
    
    Exit Sub
    
NoMoreNodes:

    Exit Sub
End Sub

Private Sub tvTreeView_Expand(Node As SSActiveTreeView.ssNode)
    Dim sdir        As String
    ' declare a dynamic array
    Dim strCache()  As String
    Dim strDirName  As String
    Dim lngCount    As Long
    Dim tmp         As String
    List1.Clear
    Dim ssNode      As ssNode

    If Not Node.Children = 0 Then Exit Sub

    ' Some files and directories, like PAGEFILE.SYS, have attributes that
    ' would cause a VB error. If so, just try to continue.
    On Error GoTo HandleError

    Select Case Node.Key
    Case "Favorites"
    Case "Playlists"
    Case "History"
    Case Else
         strDirName = Dir(Node.Key & "\", vbDirectory + vbSystem)
    
        Do While strDirName <> ""
            If strDirName <> "." And strDirName <> ".." Then
                If (GetAttr(Node.Key & "\" & strDirName) And vbDirectory) = vbDirectory Then
                   If DirName = UCase(DirName) Then
                        List1.AddItem (StrConv(strDirName, vbProperCase))
                    Else
                        List1.AddItem (strDirName)
                    End If
                lngCount = lngCount + 1
                
                End If
                
            End If
            ' get the next file or directory name
            ''debug.print strDirName
            strDirName = Dir
            If strDirName = "" Then
                'debug.print "hold"
            End If
            DoEvents
        Loop
            
        'AddDirectory entries
        For j = 1 To List1.ListCount
            Set ssNode = tvTreeView.Nodes.Add(Node.Key, ssatChild, Node.Key & "\" & List1.List(j - 1), List1.List(j - 1), "closed", "open")
            ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
            DoEvents
        Next j
        
        List1.Clear
        
        tmp = Dir(Node.Key & "\*.mp3")
        Do Until tmp = ""
            If tmp <> "." And tmp <> ".." Then
                If LCase(Right(tmp, 4)) = ".mp3" Then
                    List1.AddItem StrConv(tmp, vbProperCase)
                End If
            End If
            tmp = Dir
        Loop
        
            For j = 1 To List1.ListCount
            Set ssNode = tvTreeView.Nodes.Add(Node.Key, ssatChild, Node.Key & "\" & List1.List(j - 1), List1.List(j - 1), "mp3file", "mp3file")
            ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
            DoEvents
        Next j
        
        'debug.print "expand:" & Node.Key
    End Select
    
    Exit Sub
HandleError:
    'debug.print Err.Description
Resume Next
End Sub

Private Sub tvTreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As ssNode

Set n = tvTreeView.HitTest(X, Y)

If Not n Is Nothing Then
    '   We will set the selected item only if the SSTree is in single selection mode,
    '   otherwise we could interfere with some multi-select functionality
    If tvTreeView.NodeSelectionStyle = ssatNodeSelectSingle Then
        Set tvTreeView.SelectedItem = n
    End If
End If
Set ssNodeX = tvTreeView.SelectedItem ' Set the item being dragged.

End Sub

Private Sub tvTreeView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    ' Signal a Drag operation.
'    If Button = vbLeftButton Then
'        ' Set the flag to true.
'        indrag = True
'        ' Set the drag icon with the CreateDragImage method.
'        tvTreeView.DragIcon = tvTreeView.SelectedItem.CreateDragImage
'         ' Drag operation.
'        tvTreeView.Drag vbBeginDrag
'    End If
'
End Sub

Private Sub tvTreeView_OLEDragOver(Data As SSActiveTreeView.SSDataObject, Effect As SSActiveTreeView.SSReturnLong, Button As Integer, Shift As Integer, X As Single, Y As Single, State As SSActiveTreeView.SSReturnShort)
   
    Set tvTreeView.DropHighlight = tvTreeView.HitTest(X, Y)
        
    If Not tvTreeView.DropHighlight Is Nothing Then
            
        If tvTreeView.DropHighlight.Level > 1 Then
            Exit Sub
        End If
        
    End If
    
End Sub

Private Sub tvTreeView_OnDemandFetch(ByVal FetchBuffer As SSActiveTreeView.SSFetchBuffer)
'debug.print "on demand fetch"
End Sub

Private Sub tvTreeView_OnDemandPrepare(ParentNode As SSActiveTreeView.ssNode, Result As SSActiveTreeView.SSReturnBoolean)
    'debug.print "on demand prepare"
End Sub

Public Sub LoadDriveTree()
Set ssNode = tvTreeView.Nodes.Add(, , "ListComputer", "Computer", "computer", "computer")
     ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
     ssNode.Font.Bold = True
  
    Dim DriveNum As String
    Dim DriveType As Long
    DriveNum = 64
    On Error Resume Next

    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        Select Case DriveType
            Case 0: Set ssNode = tvTreeView.Nodes.Add("ListComputer", Chr$(DriveNum) & ":", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "unknown")
                    ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
            Case 2: Set ssNode = tvTreeView.Nodes.Add("ListComputer", ssatChild, Chr$(DriveNum) & ":", "(" & Chr$(DriveNum) & ":)", "remove")
                    ssNode.Expanded = False
            Case 3: Set ssNode = tvTreeView.Nodes.Add("ListComputer", ssatChild, Chr$(DriveNum) & ":", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "fixed")
                    ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
                    ssNode.Expanded = False
            Case 4: Set ssNode = tvTreeView.Nodes.Add("ListComputer", ssatChild, Chr$(DriveNum) & ":", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "remote")
                    ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
                    ssNode.Expanded = False
            Case 5: Set ssNode = tvTreeView.Nodes.Add("ListComputer", ssatChild, Chr$(DriveNum) & ":", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "cd")
                    ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
                    ssNode.Expanded = False
            Case 6: Set ssNode = tvTreeView.Nodes.Add("ListComputer", ssatChild, Chr$(DriveNum) & ":", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "ram")
                    ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
                    ssNode.Expanded = False
        End Select
    Loop
End Sub

Public Sub LoadHistory()
    Dim ssNode As ssNode
    Dim fs As Object
    Dim strString As String
    Dim itmX As ListItem
    Dim Counter As Integer
    Dim strTime As Date
    Dim DropOld As Boolean
    On Error Resume Next
    Set ssNode = tvTreeView.Nodes.Add(, , "History", "History", "history", "history")
    ssNode.LoadStyleChildren = ssatLoadStyleChildrenAddItem
    ssNode.Font.Bold = True
    
    List1.Clear
    Counter = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(App.Path & "\" & "history.dat") Then
        Open App.Path & "\" & "history.dat" For Input As #1
        Do While Not EOF(1)
                Input #1, strString, strTime
                If Len(strString) > 0 Then
                    If Len(strTime) > 0 Then
                        If strTime < (Date - intKeepHistory) Then
                            DropOld = True
                        Else
                            List1.AddItem (strString & vbTab & strTime)
                        End If
                    End If
                End If
        Loop
        
        If DropOld Then
            Close #1
            Open App.Path & "\" & "history.dat" For Output As #1
            For Counter = 0 To List1.ListCount
                If Len(List1.List(Counter)) > 0 Then Set ssNode = tvTreeView.Nodes.Add("History", ssatChild, "History" & "-" & List1.List(Counter), List1.List(Counter), "mp3file", "mp3file")
                If Len(List1.List(Counter)) > 0 Then Write #1, Left(List1.List(Counter), InStr(1, List1.List(Counter), vbTab) - 1); CDate(Mid$(List1.List(Counter), InStr(1, List1.List(Counter), vbTab) + 1))
            Next Counter
        Else
            For Counter = 0 To List1.ListCount
                If Len(List1.List(Counter)) > 0 Then Set ssNode = tvTreeView.Nodes.Add("History", ssatChild, "History" & "-" & List1.List(Counter), List1.List(Counter), "mp3file", "mp3file")
            Next Counter
        End If
        Close #1
    End If
    
End Sub

Public Function ExistsInHistory(FILENAME As String) As Boolean
    Dim ssNode As ssNode
    Dim tmpNode As ssNode
    Dim temp1 As String
    Dim temp2 As String
    
    temp1 = LCase("history-" & FILENAME)
    
    
    
    tvTreeView.Nodes.Item("History").Selected = True
    If tvTreeView.SelectedItem.Children > 0 Then
                    Set ssNode = tvTreeView.SelectedItem.Child
                    Set tmpNode = tvTreeView.SelectedItem
                    For i = 1 To tmpNode.Children
                        'ssNode.Expanded = True 'Needed if using OnDemand LoadStyle
            
                        tvTreeView.SelectedNodes.Add ssNode
            
                        'debug.print ssNode.Text
            
                        If i < tmpNode.Children Then
                            Set ssNode = ssNode.Next
                        End If
                        temp2 = LCase(Left(ssNode.Key, InStr(1, ssNode.Key, Chr(9)) - 1))
                        Debug.Print temp1 & "=" & temp2
                        If temp1 = temp2 Then
                           Debug.Print temp1 & "=" & temp2
                           ExistsInHistory = True
                           Exit Function
                        End If
                    Next i
    End If
    ExistsInHistory = False
    
End Function

Public Sub LoadFavorites()
    Dim ssNode As ssNode
    Dim fs As Object
    Dim strString As String
    Dim itmX As ListItem
    Dim Counter As Integer
    Dim strTime As Date
    Dim DropOld As Boolean
    
    
    List1.Clear
    Counter = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(App.Path & "\" & "Favorites.dat") Then
        Open App.Path & "\" & "Favorites.dat" For Input As #1
        Do While Not EOF(1)
                Input #1, strString, strTime
                If Len(strString) > 0 Then
                    If Len(strTime) > 0 Then
                        If strTime < (Date - intKeepHistory) Then
                            DropOld = True
                        Else
                            List1.AddItem (strString & vbTab & strTime)
                        End If
                    End If
                End If
        Loop
       
            For Counter = 0 To List1.ListCount
                If Len(List1.List(Counter)) > 0 Then Set ssNode = tvTreeView.Nodes.Add("History", ssatChild, "Favorites" & "-" & List1.List(Counter), List1.List(Counter), "mp3file", "mp3file")
            Next Counter

        Close #1
    End If
End Sub



Private Function FolderLocation(lFolder As SHFolders) As String

Dim lp As ITEMIDLIST
Dim tmpStr As String
'Get the PIDL for this folder
SHGetSpecialFolderLocation hwnd, lFolder, lp
'Convert it to a string path
tmpStr = Space$(255)
SHGetPathFromIDList lp.mkid, tmpStr
If InStr(tmpStr, Chr$(0)) > 0 Then
    'Strip nulls from the string
    tmpStr = Left$(tmpStr, InStr(tmpStr, Chr$(0)) - 1)
End If
'Free the PIDL
CoTaskMemFree lp.mkid
'Return
FolderLocation = tmpStr

End Function

Public Function Delay()
    Timer1.Enabled = True
    Do While Timer1.Enabled
        DoEvents
    Loop
End Function
