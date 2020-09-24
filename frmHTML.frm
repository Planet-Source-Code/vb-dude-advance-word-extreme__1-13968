VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Editor - By VB Dude"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmHTML.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Preview"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmHTML.frx":0442
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "HTML:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd.Filter = "Webpage Files|*.html;*.htm;"
cd.DialogTitle = "Save File"
cd.ShowSave
If cd.FileName <> "" Then
Open cd.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
End If
End Sub

Private Sub Command2_Click()
Text1.SetFocus
cd.Filter = "Webpage Files|*.html;*.htm;"
cd.ShowOpen
If cd.FileName <> "" Then
    Open cd.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, lineoftext$
    alltext$ = alltext$ & lineoftext$
    Text1.Text = alltext$
    Loop
    Close #1
End If

End Sub

Private Sub Command3_Click()
Open App.Path & "\preview.htm" For Output As #1
Print #1, Text1.Text
Close #1
frmBrowser.brwWebBrowser.Navigate App.Path & "\preview.html"
frmBrowser.cboAddress.Text = frmBrowser.brwWebBrowser.LocationURL
frmBrowser.Show vbModal, Me


End Sub

Private Sub Command4_Click()
Form1.Show vbModal, Me
End Sub
