Attribute VB_Name = "Module2"
'#This module is courtesy of Xtreme-Pad
'#Use this module only knowing where you got it from
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

Global Secret As Integer, SecStr As String
Global DefaultTitle As String, JustChanged As Boolean

Const WM_USER = &H400
Const EM_UNDO = WM_USER + 23
Global Docs As Integer
Global ChildForms(1 To 30) As Form1
Global UnAvail(1 To 30) As Boolean
Global Pos As Integer
Global SearchStr As String
Global MatchCase As Boolean
Global DefaultFontName As String
Global DefaultFontSize As Integer
Global DefaultFontColor As Long
Global DefaultFontBold As Boolean
Global DefaultFontItalic As Boolean
Global DefaultFontUnderline As Boolean
Global DefaultFontStrikethru As Boolean
Global UndoText(1 To 30) As String, Opened As Boolean
Global DocTemp As Integer, NeedSaved(30) As Boolean
Global File(1 To 30) As String, PFile(1 To 30) As String

Function GetBinary(Number As Integer) As String
Dim binstr As String
binstr = ""
Number = Number + 1
For x = 7 To 0 Step -1
  If Number > 2 ^ x Then
    Number = Number - 2 ^ x
    binstr = binstr & "1"
  Else
    binstr = binstr & "0"
  End If
Next
GetBinary = binstr
End Function

Function BintoDec(binstr As String) As Integer
Dim Number As Integer
For x = 0 To 7
  If Mid$(binstr, x + 1, 1) = "1" Then
    Number = Number + (2 ^ (7 - x))
  End If
Next
BintoDec = Number
End Function

Function frm() As Integer
On Error GoTo CreateNew
frm = Val(MDIForm1.ActiveForm.Tag)
Exit Function
CreateNew:
Dim ret As Integer
DocTemp = FirstAvail
If DocTemp <> -1 Then
  Set ChildForms(DocTemp) = New Form1
  ChildForms(DocTemp).Caption = "Document " & DocTemp
  ChildForms(DocTemp).Tag = DocTemp
Else
  MsgBox "You are only allowed 30 Documents opened at one time."
End If
frm = Val(MDIForm1.ActiveForm.Tag)
End Function

Function FirstAvail() As Integer
For x = 1 To 30
  If UnAvail(x) = False Then
    UnAvail(x) = True
    FirstAvail = x
    Exit Function
  End If
Next
FirstAvail = -1
End Function

Sub Point(mdiFrm As Form)
    With ChildForms(frm).Text1
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            .SelBullet = True
        ElseIf .SelBullet = True Then
            .SelBullet = False
            .SelHangingIndent = False
        End If
    End With
End Sub
v

