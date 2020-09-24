VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImageOpen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open Image"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Image"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   720
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog Pics 
      Left            =   2280
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Image"
   End
   Begin VB.Label Label2 
      Caption         =   "(c) Reynard Chan"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmImageOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

Private Sub Command1_Click()
Pics.Filter = "Image Files|*.jpeg;*.jpg;*.gif;*.bmp;*.ico;*.wmf;"
Pics.DialogTitle = "Select Image"
Pics.ShowOpen
If Pics.FileName <> "" Then
Picture1.Picture = LoadPicture(Pics.FileName)
Text1.Text = Pics.FileName
End If
End Sub

Private Sub Command2_Click()
   'Want to copy this code? Please ask first before
   'taking. I had to work on this for ages y' know.
   'Ask at this address: reychan@hotmail.com
   'Thanks!
   'PS: I'll most probally say yes, but due to U.S
   'copyright laws... You have to have my permission...
   
   ' Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture
    
    ' Paste the picture into the RichTextBox.
    SendMessage frmMain.ActiveForm.rtfText.hwnd, WM_PASTE, 0, 0
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

