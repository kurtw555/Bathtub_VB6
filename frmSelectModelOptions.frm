VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectModelOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Model Options"
   ClientHeight    =   11250
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6855
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "frmSelectModelOptions.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectModelOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDefaults_Click()
For i = 1 To NOptions
Iwork(i) = IopDefault(i)
Next i
ComboBox1.ListIndex = Iwork(1)
ComboBox2.ListIndex = Iwork(2)
ComboBox3.ListIndex = Iwork(3)
ComboBox4.ListIndex = Iwork(4)
ComboBox5.ListIndex = Iwork(5)
ComboBox6.ListIndex = Iwork(6)
ComboBox7.ListIndex = Iwork(7)
ComboBox8.ListIndex = Iwork(8)
ComboBox9.ListIndex = Iwork(9)
ComboBox10.ListIndex = Iwork(10)
ComboBox11.ListIndex = Iwork(11)

End Sub

Private Sub btnHelp_Click()
    ShowHelp (Ihelp)
End Sub

Private Sub btnOK_Click()
For i = 1 To NOptions
    Iop(i) = Iwork(i)
    Next i
'WriteCase   'save in workbook
FormUpdate
Unload Me
End Sub

Private Sub ComboBox1_Change()
Iwork(1) = ComboBox1.ListIndex
End Sub

Private Sub ComboBox10_Change()
Iwork(10) = ComboBox10.ListIndex
End Sub

Private Sub ComboBox11_Change()
Iwork(11) = ComboBox11.ListIndex
End Sub

Private Sub ComboBox2_Change()
Iwork(2) = ComboBox2.ListIndex
End Sub

Private Sub ComboBox3_Change()
Iwork(3) = ComboBox3.ListIndex
End Sub

Private Sub ComboBox4_Change()
Iwork(4) = ComboBox4.ListIndex
End Sub

Private Sub ComboBox5_Change()
Iwork(5) = ComboBox5.ListIndex
End Sub

Private Sub ComboBox6_Change()
Iwork(6) = ComboBox6.ListIndex
End Sub

Private Sub ComboBox7_Change()
Iwork(7) = ComboBox7.ListIndex
End Sub

Private Sub ComboBox8_Change()
Iwork(8) = ComboBox8.ListIndex
End Sub

Private Sub ComboBox9_Change()
Iwork(9) = ComboBox9.ListIndex
End Sub

Private Sub UserForm_initialize()
For i = 1 To NOptions
Iwork(i) = Iop(i)

For j = 1 To Mop(i)
    fn = Format(j - 1, "00") & " " & OptionName(i, j)
    Select Case i
    Case 1
    ComboBox1.AddItem fn
    Case 2
    ComboBox2.AddItem fn
    Case 3
    ComboBox3.AddItem fn
    Case 4
    ComboBox4.AddItem fn
    Case 5
    ComboBox5.AddItem fn
    Case 6
    ComboBox6.AddItem fn
    Case 7
    ComboBox7.AddItem fn
    Case 8
    ComboBox8.AddItem fn
    Case 9
    ComboBox9.AddItem fn
    Case 10
    ComboBox10.AddItem fn
    Case 11
    ComboBox11.AddItem fn
    End Select
    Next j
Next i
ComboBox1.ListIndex = Iwork(1)
ComboBox2.ListIndex = Iwork(2)
ComboBox3.ListIndex = Iwork(3)
ComboBox4.ListIndex = Iwork(4)
ComboBox5.ListIndex = Iwork(5)
ComboBox6.ListIndex = Iwork(6)
ComboBox7.ListIndex = Iwork(7)
ComboBox8.ListIndex = Iwork(8)
ComboBox9.ListIndex = Iwork(9)
ComboBox10.ListIndex = Iwork(10)
ComboBox11.ListIndex = Iwork(11)
End Sub
