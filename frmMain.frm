VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00808080&
   Caption         =   "Advanced Word Extreme! - By VB Dude"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Pics 
      Left            =   2880
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Image"
      FileName        =   "C:\"
      Filter          =   "*.jpeg;*.jpg*.gif;*.wmf;*.ico;*.bmp"
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2880
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0986
            Key             =   "Bullet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CAA
            Key             =   "Spell Check"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1732
            Key             =   "Redo"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   3960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5220
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5477
            Text            =   "Welcome to Advance Word 1.0"
            TextSave        =   "Welcome to Advance Word 1.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2/10/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "2:26 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21DC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22EE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2400
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2512
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2624
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2736
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2848
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":295A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A6C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B7E
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C90
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DA2
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox PictureI 
         Height          =   255
         Left            =   6120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   135
         Left            =   5760
         Top             =   120
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar Formats 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   6960
         MaskColor       =   &H00004080&
         Picture         =   "frmMain.frx":34EC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Remove Bullet"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   300
         Left            =   6600
         MaskColor       =   &H00004080&
         Picture         =   "frmMain.frx":37F6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Insert Bullet"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmMain.frx":3B00
         Left            =   3720
         List            =   "frmMain.frx":3B19
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   0
         Width           =   2775
      End
      Begin VB.ComboBox Sizes 
         Height          =   315
         ItemData        =   "frmMain.frx":3B4E
         Left            =   2760
         List            =   "frmMain.frx":3B79
         TabIndex        =   4
         Text            =   "Sizes"
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox Fonts 
         Height          =   315
         ItemData        =   "frmMain.frx":3BAF
         Left            =   120
         List            =   "frmMain.frx":3BBF
         TabIndex        =   3
         Text            =   "Fonts"
         Top             =   0
         Width           =   2535
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Begin VB.PictureBox Scale 
         BackColor       =   &H8000000E&
         Height          =   300
         Left            =   120
         MouseIcon       =   "frmMain.frx":3BF7
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3E09
         ScaleHeight     =   240
         ScaleWidth      =   11715
         TabIndex        =   7
         Top             =   0
         Width           =   11775
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   10680
            TabIndex        =   27
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   9960
            TabIndex        =   26
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   9240
            TabIndex        =   25
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   8520
            TabIndex        =   24
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   7800
            TabIndex        =   23
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   7080
            TabIndex        =   22
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Fourp5 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   6360
            TabIndex        =   21
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Four 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   5640
            TabIndex        =   20
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Threep5 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   4920
            TabIndex        =   19
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Three 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   4200
            TabIndex        =   18
            Top             =   0
            Width           =   135
         End
         Begin VB.Label TwopFive 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   3480
            TabIndex        =   17
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Two 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2760
            TabIndex        =   16
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Onep5 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2040
            TabIndex        =   15
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Zp2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   600
            TabIndex        =   14
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000007&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   135
         End
         Begin VB.Label One 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1320
            TabIndex        =   10
            Top             =   0
            Width           =   135
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndos 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuRedos 
         Caption         =   "Redo"
      End
      Begin VB.Menu afasfasf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu gfjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu afadfadf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbars 
         Caption         =   "Toolbars"
         Begin VB.Menu mnuViewGeneral 
            Caption         =   "General"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewFormat 
            Caption         =   "Format"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewMesaure 
            Caption         =   "Measurement"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu gfdjhdjdgjgj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertPicture 
         Caption         =   "&Picture from File"
      End
      Begin VB.Menu mnuInsertSep 
         Caption         =   "Sepera&tor Line"
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertSign 
         Caption         =   "Signature"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalcu 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu pagesteup 
         Caption         =   "&Page Setup"
      End
      Begin VB.Menu afafasf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebBrowser 
         Caption         =   "&Web Browser"
      End
      Begin VB.Menu mnuHTMlEdit 
         Caption         =   "&HTML Editor"
      End
      Begin VB.Menu kgkj 
         Caption         =   "-"
      End
      Begin VB.Menu refresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu pop 
      Caption         =   "PopupMenus"
      Visible         =   0   'False
      Begin VB.Menu pUndo 
         Caption         =   "U&ndo"
      End
      Begin VB.Menu agadgdag 
         Caption         =   "-"
      End
      Begin VB.Menu pCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu pCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu pPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu afaafsasfsaf 
         Caption         =   "-"
      End
      Begin VB.Menu pSeperasd 
         Caption         =   "Insert Seperator Line"
      End
      Begin VB.Menu pgasf 
         Caption         =   "Insert Gestures"
         Begin VB.Menu pHello 
            Caption         =   "Hello!"
         End
         Begin VB.Menu pBye 
            Caption         =   "Bye Bye!"
         End
         Begin VB.Menu pSmile 
            Caption         =   ": )"
         End
         Begin VB.Menu pmnuNothappy 
            Caption         =   ":("
         End
         Begin VB.Menu psucks 
            Caption         =   ":-p"
         End
         Begin VB.Menu pblink 
            Caption         =   ";)"
         End
      End
      Begin VB.Menu jhguguyjhgjh 
         Caption         =   "-"
      End
      Begin VB.Menu pChcase 
         Caption         =   "Change Case"
         Begin VB.Menu pUppercase 
            Caption         =   "Upper Cars"
         End
         Begin VB.Menu plowercase 
            Caption         =   "Lower Case"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#Advance Word 2000
'#By Reynard Chan
'#
'#Please use this code freely as you wish, though
'#easy to understand, as long as you give credit
'#to me, Reynard Chan
'#Please vote for me and give me comments at
'#Planet Source Code!
'#(c) This source code is not to be sold or used
'#for commercial use. Thankyou!
Option Explicit
Private Const WM_PASTE = &H302
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Const EM_UNDO = &HC7
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

Private Sub black_Click()
ActiveForm.rtfText.SelColor = vbBlack

End Sub

Private Sub blue_Click()
ActiveForm.rtfText.SelColor = vbBlue

End Sub

Private Sub green_Click()
ActiveForm.rtfText.SelColor = &HFF00&

End Sub


Private Sub cmdButtonBar_Click(Index As Integer)
ActiveForm.rtfText.SelText = "-------------------------------------------------------------------------------"
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()
ActiveForm.rtfText.SelFontSize = Sizes.Text
End Sub

Private Sub Combo2_Click()
ActiveForm.SetFocus
'Get color dude
Dim Wo As String
If Combo2.Text = "Black" Then
Wo = vbBlack
ElseIf Combo2.Text = "Orange" Then
Wo = &H80FF&
ElseIf Combo2.Text = "Red" Then
Wo = 255
ElseIf Combo2.Text = "Green" Then
Wo = &HFF00&
ElseIf Combo2.Text = "Purple" Then
Wo = &HFF00FF
ElseIf Combo2.Text = "Blue" Then
Wo = vbBlue
ElseIf Combo2.Text = "Yellow" Then
Wo = &HFFFF&
End If
ActiveForm.rtfText.SelColor = Wo
'Red?
'Black?
'Green?
'Orange?
'Purple
'Blue
'Yellow

End Sub

Private Sub Command1_Click()
On Error Resume Next
ActiveForm.rtfText.SetFocus
ActiveForm.rtfText.SelBullet = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
ActiveForm.rtfText.SelBullet = False
ActiveForm.rtfText.SetFocus
End Sub

Private Sub Fonts_Click()
ActiveForm.rtfText.SelFontName = Fonts.Text

End Sub

Private Sub Four_Click()
Label1.Left = Four.Left
SizeScale
End Sub

Private Sub Fourp5_Click()
Label1.Left = Fourp5.Left
SizeScale
End Sub

Private Sub Label2_Click()
Label1.Left = Label2.Left
SizeScale
End Sub

Private Sub Label3_Click()
Label1.Left = Label3.Left
SizeScale
End Sub

Private Sub Label4_Click()
Label1.Left = Label4.Left
SizeScale
End Sub

Private Sub Label5_Click()
Label1.Left = Label5.Left
SizeScale
End Sub

Private Sub Label6_Click()
Label1.Left = Label6.Left
SizeScale
End Sub

Private Sub Label7_Click()
Label1.Left = Label7.Left
SizeScale
End Sub

Private Sub MDIForm_Load()
frmOptions.Sig.Text = GetSetting("Text", "Nick", "Stuff", frmOptions.Sig.Text)
  Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
End Sub


Private Sub LoadNewDoc()
Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Show
frmD.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
Fonts.Text = ActiveForm.rtfText.SelFontName
Sizes.Text = ActiveForm.rtfText.SelFontSize
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuTextSize_Click()

End Sub

Private Sub orange_Click()
ActiveForm.rtfText.SelColor = &H80FF&

End Sub

Private Sub purple_Click()
ActiveForm.rtfText.SelColor = &HFF00FF
End Sub

Private Sub red_Click()
ActiveForm.rtfText.SelColor = 255

End Sub

Private Sub mnuCalcu_Click()
frmCalculator.Show vbModal, Me
End Sub

Private Sub mnuEditFind_Click()
frmFind.Show vbModal, Me
End Sub

Private Sub mnuEditSelectAll_Click()
ActiveForm.rtfText.SetFocus
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText)
End Sub

Private Sub mnuHTMlEdit_Click()
Form2.Show vbModal, Me
End Sub

Private Sub mnuInsertPicture_Click()
On Error Resume Next
frmImageOpen.Show vbModal, Me
End Sub

Private Sub mnuInsertSep_Click()
Call pSeperasd_Click
End Sub

Private Sub mnuInsertSign_Click()
ActiveForm.rtfText.SelText = frmOptions.Sig.Text
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show vbModal, Me
End Sub

Private Sub mnuRedos_Click()
 On Error Resume Next
 Call Undo(False)
End Sub

Private Sub mnuUndos_Click()
On Error Resume Next
    Call Undo(True)
  
End Sub
Public Sub Undo(ByVal bUndo As Boolean)
    
  Dim OK As Long
    
  OK = SendMessage(Screen.ActiveForm.ActiveControl.hwnd, EM_UNDO, 0&, 0&)
  
 Exit Sub
End Sub
Private Sub mnuViewFormat_Click()
If mnuViewFormat.Checked = True Then
mnuViewFormat.Checked = False
Formats.Visible = False
Else
mnuViewFormat.Checked = True
Formats.Visible = True
End If
End Sub

Private Sub mnuViewGeneral_Click()
If mnuViewGeneral.Checked = True Then
mnuViewGeneral.Checked = False
tbToolBar.Visible = False
Else
mnuViewGeneral.Checked = True
tbToolBar.Visible = True
End If
End Sub

Private Sub mnuViewMesaure_Click()
If mnuViewMesaure.Checked = True Then
mnuViewMesaure.Checked = False
Toolbar1.Visible = False
Else
mnuViewMesaure.Checked = True
Toolbar1.Visible = True
End If
End Sub

Private Sub mnuViewStatus_Click()
If mnuViewStatus.Checked = True Then
mnuViewStatus.Checked = False
sbStatusBar.Visible = False
Else
mnuViewStatus.Checked = True
sbStatusBar.Visible = True
End If
End Sub


Private Sub mnuWebBrowser_Click()
frmBrowser.Show
End Sub

Sub SizeScale()
ActiveForm.rtfText.Left = Label1.Left
End Sub
Private Sub One_Click()
Label1.Left = 1320
SizeScale
End Sub

Private Sub Onep5_Click()
Label1.Left = Onep5.Left
SizeScale
End Sub

Private Sub pagesteup_Click()
frmPage.Show vbModal, Me
End Sub

Private Sub pblink_Click()
ActiveForm.rtfText.SelText = pblink.Caption
End Sub

Private Sub pBye_Click()
ActiveForm.rtfText.SelText = pBye.Caption

End Sub

Private Sub pCopy_Click()
Call mnuEditCopy_Click
End Sub

Private Sub pCut_Click()
Call mnuEditCut_Click
End Sub

Private Sub pHello_Click()
ActiveForm.rtfText.SelText = pHello.Caption
End Sub

Private Sub Picture1_Click()
ActiveForm.SetFocus
End Sub

Private Sub plowercase_Click()
Clipboard.SetText ActiveForm.rtfText.SelText
ActiveForm.rtfText.SelText = LCase(Clipboard.GetText)

End Sub

Private Sub pmnuNothappy_Click()
ActiveForm.rtfText.SelText = pmnuNothappy.Caption

End Sub

Private Sub pPaste_Click()
Call mnuEditPaste_Click
End Sub

Private Sub pSeperasd_Click()
ActiveForm.rtfText.SelText = "-----------------------------------------------------------------"
End Sub

Private Sub pSmile_Click()
ActiveForm.rtfText.SelText = pSmile.Caption

End Sub

Private Sub psucks_Click()
ActiveForm.rtfText.SelText = psucks.Caption

End Sub

Private Sub pUndo_Click()
Call mnuUndos_Click
End Sub

Private Sub pUppercase_Click()
Clipboard.SetText ActiveForm.rtfText.SelText
ActiveForm.rtfText.SelText = UCase(Clipboard.GetText)

End Sub

Private Sub Sizes_Click()
ActiveForm.rtfText.SelFontSize = Sizes.Text
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
Form1.Show vbModal, Me
End Sub
Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub



Private Sub mnuFileExit_Click()
    'unload the form
    End
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub




Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Text Files (*.txt)|*.txt;"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = Me.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub


Private Sub mnuFileOpen_Click()
    Dim sFile As String


    If ActiveForm Is Nothing Then LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Text Files (*.txt)|*.txt"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.rtfText.LoadFile sFile
    Me.Caption = sFile

End Sub

Private Sub mnuFileNew_Click()
If MsgBox("Wouldn't you want to save your work?", vbInformation + vbYesNo, "Save your Work!") = vbYes Then
    Call mnuFileSave_Click
Else
LoadNewDoc
End If
End Sub

Private Sub white_Click()
ActiveForm.rtfText.SelColor = vbWhite

End Sub

Private Sub yeallow_Click()
ActiveForm.rtfText.SelColor = &HFFFF&
End Sub




Private Sub Three_Click()
Label1.Left = Three.Left
SizeScale
End Sub

Private Sub Threep5_Click()
Label1.Left = Threep5.Left
SizeScale
End Sub

Private Sub Two_Click()
Label1.Left = Two.Left
SizeScale
End Sub

Private Sub TwopFive_Click()
Label1.Left = TwopFive.Left
SizeScale
End Sub


Private Sub Zp2_Click()
Label1.Left = Zp2.Left
SizeScale
End Sub
