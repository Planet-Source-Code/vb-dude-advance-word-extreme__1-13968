VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2730
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Hidden 
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
      Begin VB.CommandButton Command19 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   18
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command17 
         Caption         =   "¸"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command16 
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command15 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "__"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   14
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   12
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   9
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Label Sign 
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "By VB Dude - Reynard Chan"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text1.SelText = "0"
End Sub

Private Sub Command10_Click()
If Text1.Text = "0" Then
Text1.Text = "3"
Else
Text1.SelText = 3
End If
End Sub

Private Sub Command13_Click()
Sign.Caption = "+"
Hidden.Text = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command14_Click()
Sign.Caption = "-"
Hidden.Text = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command15_Click()
Sign.Caption = "*"
Hidden.Text = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command16_Click()
Text1.Text = "0"
End Sub

Private Sub Command17_Click()
Sign.Caption = "/"
Hidden.Text = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command19_Click()
Dim x As String
Dim y As String
Dim z As String
If Sign.Caption = "+" Then
x = Text1.Text
y = Hidden.Text
z = Str(Val(x) + Val(y))
Text1.Text = z
ElseIf Sign.Caption = "-" Then
x = Text1.Text
y = Hidden.Text
z = Str(Val(y) - Val(x))
Text1.Text = z
ElseIf Sign.Caption = "/" Then
x = Text1.Text
y = Hidden.Text
z = Str(Val(y) / Val(x))
Text1.Text = z
ElseIf Sign.Caption = "*" Then
x = Text1.Text
y = Hidden.Text
z = Str(Val(y) * Val(x))
Text1.Text = z
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "0" Then
Text1.Text = "7"
Else
Text1.SelText = 7
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "0" Then
Text1.Text = "8"
Else
Text1.SelText = "8"
End If

End Sub

Private Sub Command4_Click()
If Text1.Text = "0" Then
Text1.Text = "9"
Else
Text1.SelText = 9
End If

End Sub

Private Sub Command5_Click()
If Text1.Text = "0" Then
Text1.Text = "4"
Else
Text1.SelText = 4
End If
End Sub

Private Sub Command6_Click()
If Text1.Text = "0" Then
Text1.Text = "1"
Else
Text1.SelText = 1
End If
End Sub

Private Sub Command7_Click()
If Text1.Text = "0" Then
Text1.Text = "6"
Else
Text1.SelText = 6
End If

End Sub

Private Sub Command8_Click()
If Text1.Text = "0" Then
Text1.Text = "5"
Else
Text1.SelText = 5
End If

End Sub

Private Sub Command9_Click()
If Text1.Text = "0" Then
Text1.Text = "2"
Else
Text1.SelText = 2
End If

End Sub
