Attribute VB_Name = "Module1"
Public fMainForm As frmMain
Public Const WM_USER = &H400
Public Const EM_UNDO = &HC7


Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub
