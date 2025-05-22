VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChannels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Transport Channels"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   HelpContextID   =   8
   Icon            =   "frmChannels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select Upstream Segment "
      Top             =   2760
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select Downstream Segment"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Channel to be Edited"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   3600
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2760
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      HelpContextID   =   8
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "btnList"
            Object.ToolTipText     =   "List segment, tributary, & channel network"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "btnAdd"
            Object.ToolTipText     =   "Add a new channel"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "btnDelete"
            Object.ToolTipText     =   "Delete selected channel"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Defaults"
            Key             =   "bthDefaults"
            Object.ToolTipText     =   "Assign default values to all input values for selected channel"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Undo"
            Key             =   "btnUndo"
            Object.ToolTipText     =   "Restore initial values for all channels"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Get Help"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Key             =   "btnCancel"
            Object.ToolTipText     =   "Ignore all edits & return to program menu"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OK"
            Key             =   "btnOK"
            Object.ToolTipText     =   "Save edits & return to program menu"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "To Segment:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   15
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From Segment:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Advective Flow (hm3/yr):"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Exchange Flow (hm3/yr):"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Channel Name:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   960
      TabIndex        =   11
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mean"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CV"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4680
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblDefinitions 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
   End
End
Attribute VB_Name = "frmChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nboX As Integer = 4
Dim jpipE As Integer

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "List"
    List_Tree

Case "Add"
    If Npipe = NPMAX Then
            MsgBox "Too many channels!", vbExclamation
     Else
        If Npipe > 0 Then UpdateChannelValues (2) 'save current channel
        If CheckValues Then
            Npipe = Npipe + 1
            jpipE = Npipe
            Call PipeZero(jpipE)
            PipeName(jpipE) = "New Channel"
            UpdateCombos
            UpdateChannelValues (1)
            End If
            End If

Case "Delete"

'delete a pipe
If Npipe = 0 Or jpipE = Npipe + 1 Then Exit Sub
If MsgBox("Delete This Channel?", vbYesNo) = vbYes Then
    k = 0
    For i = 1 To Npipe
        If i <> jpipE Then
            k = k + 1
            Call PipeCopy(i, k)
            End If
        Next i
       Call PipeZero(Npipe)
       Npipe = Npipe - 1
       jpipE = 1
       UpdateCombos
       UpdateChannelValues (1)
       End If

Case "Clear"
If MsgBox("Clear Values for Current Channel?", vbYesNo) = vbYes Then
            UpdateChannelValues (3)
            UpdateChannelValues (2)
            UpdateCombos
            End If

Case "Undo"
    Backup (1)
    jpipE = 1
    UpdateCombos
    UpdateChannelValues (1)

Case "Help"
    ShowHelp (HelpContextID)

Case "Cancel"
    Backup (1)
    Unload Me
    
Case "OK"
     UpdateChannelValues (2)
        If CheckValues Then
            Icalc = 0
            Unload Me
            End If

End Select

End Sub

Private Sub Form_Load()
    jpipE = 1
    UpdateCombos
    SetToolTips
    Backup (0)    'save backup file
    UpdateChannelValues (1)
End Sub
Private Sub Combo1_click()
'load new channel
    j = Combo1.ListIndex + 1
    'UpdateChannelValues (2)
    If j <> jpipE And jpipE <= Npipe Then
            UpdateChannelValues (2)
            If Not CheckValues Then
                Combo1.ListIndex = jpipE - 1
                Exit Sub
                End If
            End If
       
    jpipE = j
    Combo1.ListIndex = jpipE - 1
    UpdateChannelValues (1)
End Sub

Private Sub UpdateCombos()
    
    With Combo1
    .Clear
    For i = 1 To Npipe
            .AddItem Format(i, "00") & " " & PipeName(i)
            'End If
        Next i
    .AddItem " "
    .ListIndex = jpipE - 1
    End With

    Label2.Caption = "Total Channels = " & Npipe

    With Combo2
    .Clear
    .AddItem "None"
    For i = 1 To Nseg
        .AddItem Format(i, "00") & " " & SegName(i)
        Next i
    .ListIndex = Ito(jpipE)
    End With
    
    With Combo3
    .Clear
    .AddItem "None"
    For i = 1 To Nseg
        .AddItem Format(i, "00") & " " & SegName(i)
        Next i
    .ListIndex = Ifr(jpipE)
    End With
    
    If Npipe > 0 And PipeName(jpipE) <> "New Channel" And Ifr(jpipE) = Ito(jpipE) Then MsgBox "The inflow segment cannot equal the outflow segment", vbExclamation
    
        
    End Sub

Private Sub Combo3_Click()
'change to segment
    Text1(5) = Combo3.ListIndex
End Sub
Private Sub Combo2_Click()
'change from segment
    Text1(6) = Combo2.ListIndex
End Sub

Private Sub UpdateChannelValues(io)
'io=1 copy source values to temporary array
'io=2 copy from temporary array to original
'io=3 copy defaults

Select Case io

Case 1   'fill temporary array
   
    Text1(0) = PipeName(jpipE)   'channel label
    Text1(1) = Qpipe(jpipE)
    Text1(2) = CvQpipe(jpipE)
    Text1(3) = Epipe(jpipE)
    Text1(4) = CvEpipe(jpipE)
    Text1(5) = Ifr(jpipE)
    Text1(6) = Ito(jpipE)
    On Error Resume Next
    Combo2.ListIndex = Ito(jpipE)
    Combo3.ListIndex = Ifr(jpipE)
    On Error GoTo 0

Case 2    'copy from temporary array back to original values
    If jpipE <= Npipe Then
        PipeName(jpipE) = Text1(0) 'channel label
        Qpipe(jpipE) = Text1(1)
        CvQpipe(jpipE) = Text1(2)
        Epipe(jpipE) = Text1(3)
        CvEpipe(jpipE) = Text1(4)
        Ifr(jpipE) = Text1(5)
        Ito(jpipE) = Text1(6)
        On Error Resume Next
        Combo1.List(jpipE - 1) = Format(jpipE, "00") & " " & Text1(0)
        On Error GoTo 0
    End If
    
Case 3    'set defaults

    For j = 1 To 4
        Text1(j) = 0
        Next j

Case default

End Select

End Sub

Function CheckValues() As Boolean

If Npipe > 0 And (Ifr(jpipE) = 0 Or Ito(jpipE) = 0 Or Ito(jpipE) = Ifr(jpipE)) Then
    MsgBox "Invalid source or downstream segment"
    CheckValues = False
    Else
    CheckValues = True
    End If

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = ""
End Sub

Private Sub SetToolTips()
             Text1(0).ToolTipText = "Channel Name"
             Text1(1).ToolTipText = "Advective Flow (hm3/yr)"
             Text1(2).ToolTipText = "CV of Advective Flow"
             Text1(3).ToolTipText = "Exchange Flow (hm3/yr)"
             Text1(4).ToolTipText = "CV of Exchange Flow"
             Text1(5).ToolTipText = "Source Segment"
             Text1(6).ToolTipText = "Downstream Segment"
            End Sub

Private Sub text1_Change(index As Integer)
    If index <> 0 Then
        If Not VerifyPositive(Text1(index).Text) Then Text1(index).Text = ""
        End If
            
End Sub
Private Sub text1_lostfocus(index As Integer)
    If index <> 0 Then
        If Text1(index).Text = "" Then Text1(index).Text = 0
        End If
            
End Sub
Private Sub text1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = Text1(index).ToolTipText
End Sub


