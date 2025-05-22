VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalibration 
   Caption         =   "Calibration"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   46
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Calibration Factors Apply To"
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Calibration Method"
      Top             =   3600
      Width           =   3615
      Begin VB.OptionButton OptOption 
         Caption         =   "Predicted Concentrations"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton OptOption 
         Caption         =   "Sedimentation Rates (default)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Calculations"
      Height          =   3015
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   7695
      Begin VB.TextBox txtCalib 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Text            =   "Calibration.frx":0000
         ToolTipText     =   "Shows progress of calibration calculations"
         Top             =   360
         Width           =   7215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Segments"
      Height          =   3495
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   2940
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   12
         ToolTipText     =   "Select Segments to be Used in Calibration (not available if calibration type = global)"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calibration Type"
      Height          =   1575
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
      Begin VB.OptionButton optMethod 
         Caption         =   "By Segment"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Global applies to all segments"
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "By Segment Group"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Global"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   1111
      ButtonWidth     =   1296
      ButtonHeight    =   953
      Appearance      =   1
      HelpContextID   =   46
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run"
            Key             =   "btnRun"
            Object.ToolTipText     =   "Run calibration"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Key             =   "btnReset"
            Object.ToolTipText     =   "Reset calibration factors for selected variables and segments"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset All"
            Key             =   "btnResetAll"
            Object.ToolTipText     =   "Reset calibration factors for all segments and variables"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "btnList"
            Object.ToolTipText     =   "List calibration results"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Get help"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            Key             =   "btnQuit"
            Object.ToolTipText     =   "Return to program menu"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calibration Variables"
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Select variables to be calibrated (in sequence)"
      Top             =   1560
      Width           =   2415
      Begin VB.CheckBox chkVariable 
         Caption         =   "Chlorophyll-a"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkVariable 
         Caption         =   "Total N"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkVariable 
         Caption         =   "Total P"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkVariable 
         Caption         =   "Conservative Subst"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label lblDefinitions 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frmCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

i = MAx(0, MIn(2, Iop(7) - 1))
OptOption(i) = True

lblDefinitions.Caption = "Ready"
txtCalib.Text = ""
For j = 1 To 4
    chkVariable(j - 1).Value = 0
    chkVariable(j - 1).Enabled = False
    k = 0
    For i = 1 To Nseg
            If Cobs(i, j) > 0 Then k = k + 1
            Next i
    If k > 0 And Iop(j) > 0 Then chkVariable(j - 1).Enabled = True
    Next j
If chkVariable(1).Enabled = True Then chkVariable(1).Value = 1

With List1
    .Clear
    For i = 1 To Nseg
        .AddItem Format(i, "00") & " " & SegName(i)
        Next i
    .ListIndex = 1
    End With
    
optMethod(0) = True

'segment list
If optMethod(0).Value = True Then
    List1.Enabled = False
    Else
    List1.Enabled = True
    End If

End Sub

Function optM()
    For i = 0 To 2
    If optMethod(i).Value = True Then optM = i
    Next i
End Function

Private Sub optMethod_Click(index As Integer)

If index = 0 Then
    List1.Enabled = False
    Else
    List1.Enabled = True
    End If

End Sub

Private Sub OptOption_Click(index As Integer)
    Iop(7) = index + 1
    Iop(8) = index + 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "Run"

Select Case optM
    Case 0   'global
        Call GlobalCalib(2)
    Case 1   'regional
        Call LocalCalib(1)
    Case 2   'local
        Call LocalCalib(0)
    End Select
    FitUpdate

Case "Reset All"
    txtCalib.Text = ""
     For j = 1 To 4
        Xk(j) = 1
        For i = 1 To Nseg
            Cal(i, j) = 1
            Next i
        Next j
     Model
     FitUpdate
     
Case "Reset"
    txtCalib.Text = ""
     For j = 1 To 4
        If chkVariable(j - 1).Value Then
            If optM = 0 Then
                Xk(j) = 1
            Else
                For i = 1 To Nseg
                    If List1.Selected(i - 1) = True Then Cal(i, j) = 1
                    Next i
            End If
        End If
        Next j
     Model
     FitUpdate
     
Case "List"
    Model
    ContextId = frmCalibration.HelpContextID
    FitUpdate
            
Case "Help"
    ShowHelp (frmCalibration.HelpContextID)

Case "Quit"
    Icalc = 0
    FormUpdate
    Unload Me
End Select

End Sub
Sub FitUpdate()
    If Icalc = 0 Then Exit Sub
    lblDefinitions.Caption = "Listing Results..."
    List_Fits
    ViewSheet ("calibrations")
    lblDefinitions.Caption = "Ready"
End Sub
