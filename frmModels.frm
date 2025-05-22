VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModels 
   Caption         =   "Select Models"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   14
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   11
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3735
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   10
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3135
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   9
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2535
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   8
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1935
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   7
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1335
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   6
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   735
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3735
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3135
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2535
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1935
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1335
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   735
      Width           =   3255
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      HelpContextID   =   14
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Defaults"
            Key             =   "bthDefaults"
            Object.ToolTipText     =   "Assign default values to all input values"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Undo"
            Key             =   "btnUndo"
            Object.ToolTipText     =   "Restore initial values"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Get Help"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Key             =   "btnCancel"
            Object.ToolTipText     =   "Ignore edits & return to menu"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OK"
            Key             =   "btnOK"
            Object.ToolTipText     =   "Save edits & return to program menu"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Select Box and Hit F1 to Get Help,    *=Default"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Output Destination"
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   24
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Mass Balance Tables"
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   22
      Top             =   3240
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Availability Factors"
      Height          =   255
      Index           =   9
      Left            =   6120
      TabIndex        =   20
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Error Analysis"
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   18
      Top             =   2040
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Nitrogen Calibration "
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   16
      Top             =   1440
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Phosphorus Calibration "
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   14
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Longitudinal Dispersion"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Transparency"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Chlorophyll-a"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Total Nitrogen"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Total Phosphorus"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Conservative Substance"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2400
   End
End
Attribute VB_Name = "frmModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Changed As Boolean

Private Sub Combo1_click(index As Integer)
    j = Iwork(index + 1)
    Iwork(index + 1) = Combo1(index).ListIndex
    If j <> Combo1(index).ListIndex Then Changed = True
End Sub

Private Sub Form_Load()

    For i = 1 To NOptions
        Iwork(i) = Iop(i)
        For j = 1 To Mop(i)
            fn = Format(j - 1, "00") & " " & OptionName(i, j)
            If j - 1 = IopDefault(i) Then fn = fn & " *"
            Combo1(i - 1).AddItem fn
            
        Next j
    Combo1(i - 1).ListIndex = Iwork(i)
    Combo1(i - 1).HelpContextID = frmModels.HelpContextID
    Next i
'    Changed = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "Defaults"
    If MsgBox("Assign Default Values for All Options?", vbYesNo) = vbYes Then
    For i = 1 To NOptions
        Iwork(i) = IopDefault(i)
        Combo1(i - 1).ListIndex = Iwork(i)
        Next i
'    Changed = True
    End If

Case "Undo"
 '   If MsgBox("Assign Previously Saved Values?", vbYesNo) = vbYes Then
     For i = 1 To NOptions
         Iwork(i) = Iop(i)
         Combo1(i - 1).ListIndex = Iwork(i)
         Next i
'      Changed = False
 '    End If

Case "Help"
    ShowHelp (frmModels.HelpContextID)

Case "Cancel"
    Unload Me
 
Case "OK"

'If Changed = True Then
's = MsgBox("Save Edits?", vbYesNoCancel)
'    If s = vbCancel Then Exit Sub
'    If s = vbYes Then
    For i = 1 To NOptions
        Iop(i) = Iwork(i)
        Next i
'        End If
'     End If
Icalc = 0
FormUpdate
Unload Me

End Select

End Sub
