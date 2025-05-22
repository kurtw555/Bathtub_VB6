VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Information"
   ClientHeight    =   2880
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5628
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1987.827
   ScaleMode       =   0  'User
   ScaleWidth      =   5282.165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Disclaimer"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   362
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Read disclaimer statement"
      Top             =   2280
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Documentation"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   362
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "View documentation & help screens"
      Top             =   2280
      Width           =   1500
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   432
      Left            =   120
      Picture         =   "frmAbout.frx":09AA
      ScaleHeight     =   263.118
      ScaleMode       =   0  'User
      ScaleWidth      =   263.118
      TabIndex        =   0
      Top             =   120
      Width           =   432
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   362
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Start program"
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "U.S. Army Corps of Engineers Waterways Experiment Station"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   2805
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Simplified Techniques for Eutrophication Assessment and Prediction"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Bathtub for Windows Version 6.14"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5210.835
      Y1              =   1490.87
      Y2              =   1490.87
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Check1_Click(Index As Integer)
' frmMenu.mnuUser_Click (Index)
'End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
ShowHelp (100)
End Sub

Private Sub Command2_Click()
ShowHelp (31)
End Sub

Private Sub form_initialize()
'define help
On Error Resume Next
'In some installations the App.path returns a null string - read/write privileges?
If (App.Path = "") Then MsgBox ("FrmAbout: WARNING: App.Path is Null - could be read/write privilege issue")
Directory = App.Path & Application.PathSeparator
App.HelpFile = Directory & BathtubHelpFile
hHelp.CHMFile = Directory & BathtubHelpFile
If DebugMode Then MsgBox ("frmAbout: ROOT DIRECTORY is established as " & Directory)
On Error GoTo 0
End Sub

Private Sub Form_Load()
'If Check1(0).Value = Checked Then
'    Index = 0
'    Else
'    Index = 1
'    End If
'frmMenu.mnuUser_Click (Index)

End Sub

