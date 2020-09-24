VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3975
            Begin VB.TextBox Sig 
               Height          =   285
               Left            =   1320
               TabIndex        =   6
               Text            =   "Nickname"
               Top             =   240
               Width           =   2535
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Signature Name:"
               Height          =   195
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   1185
            End
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
SaveSetting "General", "Check", "Stuff", Check1.Value
ElseIf Check1.Value = 0 Then
SaveSetting "General", "Check", "Stuff", Check1.Value
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
SaveSetting "Text", "Nick", "Stuff", Sig.Text
Unload Me
End Sub

Private Sub Form_Load()
Sig.Text = GetSetting("Text", "Nick", "Stuff", Sig.Text)
End Sub
