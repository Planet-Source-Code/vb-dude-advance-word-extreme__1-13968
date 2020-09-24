VERSION 5.00
Begin VB.Form frmPage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Setup"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2730
   Icon            =   "frmPage.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Height of Page:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Width of Page:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmMain.ActiveForm.Width = Text1.Text
frmMain.ActiveForm.Height = Text2.Text
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = frmMain.ActiveForm.Width
Text2.Text = frmMain.ActiveForm.Height
End Sub

Private Sub Text1_Change()
If Not IsNumeric(Text1.Text) Then
Text1.Text = ""
End If
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2.Text) Then
Text2.Text = ""
End If

End Sub
