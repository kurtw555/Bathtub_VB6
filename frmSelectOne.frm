VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectOne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select One Item"
   ClientHeight    =   5115
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   3210
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "frmSelectOne.dsx":0000
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSelectOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Ichoice = 0
Unload Me
End Sub

Private Sub HelpButton_Click()
ShowHelp (Ihelp)
End Sub

Private Sub OKButton_Click()
'MsgBox "Selected Text: " & UserForm1.ListBox1.Text
'MsgBox "You have selected Item # " & UserForm1.ListBox1.ListIndex

'MsgBox "OK Clicked"
If ListBox1.MultiSelect > 0 Then
'multi pick
nN = ListBox1.ListCount
For j = 1 To nN
If ListBox1.Selected(j - 1) Then
    Iwork(j) = 1
    Else
    Iwork(j) = 0
    End If
    Next j
Else
'single pick
Ichoice = ListBox1.ListIndex + 1
End If
'MsgBox "OK bye"
Unload Me
End Sub
Private Sub UserForm_Load()
'MsgBox "form loaded"
'ListBox1.ListIndex = 0
End Sub

