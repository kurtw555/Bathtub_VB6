VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtBox 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmBox.frx":09AA
      Top             =   720
      Width           =   10335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1111
      ButtonWidth     =   1614
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "btnSave"
            Object.ToolTipText     =   "Save input values "
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy Table"
            Object.ToolTipText     =   "Copy table to windows clipboard"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "Get Help"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "btnClose"
            Object.ToolTipText     =   "Stop editing & return to program menu"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim outfilE As String

Select Case Button

Case "Save"

     With CommonDialog1
        .FileName = "*.txt"
        .Filter = "Text File(*.txt)|*.txt|"
        .FilterIndex = 1
        .ShowSave
        outfilE = .FileName
        End With
    
    If outfilE = "" Or InStr(outfilE, "*") > 0 Then Exit Sub
    
    'If Not ValidFile(outfilE) Then Exit Sub
    If FileExists(outfilE) Then
        If MsgBox("File: " & outfilE & " already exists, overwrite?", vbYesNo) <> vbYes Then Exit Sub
        End If
    
    Open outfilE For Output As #1
    Write #1, txtBox.Text
    Close #1
    MsgBox "File: " & outfilE & " saved"
    
'Case "Print"

'    PrintForm

Case "Help"
    ShowHelp (ContextId)

Case "Copy Table"
    Clipboard.Clear
    Clipboard.SetText txtBox.Text
    
Case "Close"
    Unload Me

End Select

End Sub

Private Sub txtBox_Change()

'frmBox.Visible = False
'frmBox.WindowState = 2

    Dim txt As String
    txt = txtBox.Text
    Call TextDim(txt, w, h)
    F = txtBox.Font.Size
    Fw = F * 14
    Fh = F * 32
    dw = 500
    dh = 1200
    
txtBox.Width = MAx(dw * 7, Fw * w)
txtBox.Height = MAx(dh * 2, Fh * MIn(h, 50))
frmBox.Width = txtBox.Width + dw
frmBox.Height = txtBox.Height + dh
HelpContextID = ContextId
'frmBox.Visible = True
'frmBox.WindowState = 0
End Sub


Sub TextDim(txt As String, w, h)
'count lines & maximum width of text string

    n1 = Len(txt)
    n1 = MIn(n1, 2000)
    h = 0
    i = 1
    w = 0
    w1 = 0
Do Until i >= n1
    j = InStr(i, txt, vbCrLf)
    w1 = j - i + 1
    w = MAx(w, w1)
    h = h + 1
    i = j + 1
    Loop
'MsgBox "width = " & w & "  height = " & h
End Sub
