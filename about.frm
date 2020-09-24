VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3435
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1680
      ScaleHeight     =   1665
      ScaleWidth      =   3765
      TabIndex        =   3
      Top             =   1200
      Width           =   3795
      Begin VB.Label L2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   $"about.frx":0000
         Height          =   2895
         Left            =   0
         TabIndex        =   5
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advance Word Extreme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   4
         Top             =   1680
         Width           =   3735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "By VB Dude - Reynard Chan"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"about.frx":01A3
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = True
End Sub

Private Sub L2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If L1.Top <= -800 Then L1.Top = 1560
If L2.Top <= -240 Then L2.Top = 2640

L1.Top = L1.Top - 15
L2.Top = L2.Top - 15

End Sub
