VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Words"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Find:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim SearchResult As Long
Dim op As String
op = ""
SearchResult = frmMain.ActiveForm.rtfText.Find(Text1.Text, 0, Len(frmMain.ActiveForm.rtfText.Text), Val(op))
        If SearchResult >= 0 Then
            Unload Me
            Exit Sub
        Else
            MsgBox "Search text not found.", vbOKOnly + vbExclamation, "Advance Word 2000"
            Exit Sub
        End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
