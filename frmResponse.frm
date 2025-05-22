VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResponse 
   Caption         =   "Load Response"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1009
   Icon            =   "frmResponse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbOption 
      Height          =   360
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   14
      ToolTipText     =   "Select method for varying loads"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   13
      Text            =   "2.0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   12
      Text            =   "0.2"
      ToolTipText     =   "Scale factors applied to existing loads"
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox cmbVariable 
      Height          =   360
      ItemData        =   "frmResponse.frx":09AA
      Left            =   1080
      List            =   "frmResponse.frx":09AC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Test this response variable"
      Top             =   2400
      Width           =   3015
   End
   Begin VB.ComboBox cmbSegment 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Test response of this segment"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.ComboBox cmbTrib 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Vary P load in this tributary"
      Top             =   1440
      Width           =   3015
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   1111
      ButtonWidth     =   1640
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run"
            Object.ToolTipText     =   "Run Model"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Object.ToolTipText     =   "List Results"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy Chart"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "Get Help"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            Object.ToolTipText     =   "Return to program menu"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Metamodel"
         EndProperty
      EndProperty
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Variable:"
         Height          =   900
         Index           =   4
         Left            =   960
         TabIndex        =   9
         Top             =   3480
         Width           =   3000
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Method:"
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   15
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5775
      Left            =   120
      ToolTipText     =   "Load/Response Plot"
      Top             =   3000
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "High:"
      Height          =   375
      Index           =   6
      Left            =   7080
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TP Load Scale   Low:"
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   10
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Variable:"
      Height          =   255
      Index           =   2
      Left            =   -120
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Segment:"
      Height          =   255
      Index           =   1
      Left            =   -120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tributary:"
      Height          =   255
      Index           =   0
      Left            =   -120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShowStash As Boolean
'Dim CurrentWKChart As Object 'Excel.Chart
Dim reD As Boolean
Dim Fname As String

Sub ClearIt()
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
    Image1.Picture = LoadPicture("")
    reD = False
End Sub
Private Sub cmbSegment_Click()
    If reD Then ClearIt
End Sub
Private Sub cmbtrib_Click()
    If reD Then ClearIt
End Sub
Private Sub cmbvariable_Click()
    If reD Then ClearIt
End Sub
Private Sub cmboption_Click()
    If reD Then ClearIt
End Sub
Private Sub RunModels()
Dim icount As Integer
     If MsgBox("  This Will Take A While - OK to Proceed?  ", vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
     For RespVarCode = 1 To 12 'Full run is 12
       Lindex = RespVarCode
       If Lindex > 1 Then Lindex = Lindex + 1
       If Lindex > 4 Then Lindex = Lindex + 14
       cmbVariable.ListIndex = Lindex
       For Model_type = 0 To 1
       cmbOption.ListIndex = Model_type
       ShowStash = ShowWarnings
       ShowWarnings = False
       ClearIt
       Run_Response
       If Ier = 0 Then
       'Restored 7/3/2009 DMS
       If DebugMode Then MsgBox ("frmresponse N25: chart Object Count=" & Str(Wkb.Sheets("plot response").ChartObjects.Count))
       Set CurrentWKChart = Wkb.Sheets("plot response").ChartObjects(1).Chart
        Fname = Directory & "temp.gif"
        CurrentWKChart.Export FileName:=Fname, FilterName:="GIF"
        Set CurrentWKChart = Wkb.Sheets("plot response").ChartObjects(1).Chart
         Image1.Picture = LoadPicture(Fname)
         If DebugMode Then MsgBox ("frmresponse N26: About to Delete " & Directory & "temp.gif")
         Kill Fname
        '=====================================
       
       
       
          reD = True
          End If
       ShowWarnings = ShowStash

       If reD Then
       ViewSheet ("load response")
       'Copy to MetaModels
       With Wko
         If .Sheets.Count > 2 Then
         Set gSheetout = .Worksheets("MetaModels")
         ResponseCount = ResponseCount + 10
         Else
           ResponseCount = 2
           Wkb.Sheets("MetaModels").Copy After:=.Worksheets("load response")
           .ActiveSheet.Name = "MetaModels"
           Set gSheetout = .ActiveSheet
           Wka.ActiveWindow.DisplayGridlines = False
           End If
         End With
       gLSht.Rows("11:20").Copy gSheetout.Range("A" & ResponseCount)
       'Copy ancillary info on parameter and model type
       For j = ResponseCount To ResponseCount + 9
          gSheetout.Cells(j, "I") = cmbVariable.ListIndex
          gSheetout.Cells(j, "J") = cmbOption.ListIndex
          icount = cmbSegment.ListIndex + 1
          If icount = cmbSegment.ListCount Then icount = 0
          gSheetout.Cells(j, "K") = icount
          gSheetout.Cells(j, "L") = cmbVariable.Text
          gSheetout.Cells(j, "M") = cmbOption.Text
          gSheetout.Cells(j, "N") = cmbSegment.Text
          Next j
       End If
     Next Model_type
   Next RespVarCode
   Wko.Sheets("load response").Delete
   Wko.Sheets("MetaModels").Activate

End Sub 'Run_Models
Private Sub Form_Load()

reD = True
'build boxes
With cmbTrib
    .Clear
    .AddItem "All"
    For i = 1 To NTrib
        .AddItem Format(i, "00") & " " & TribName(i)
        Next i
    .ListIndex = 0
End With

With cmbSegment
    .Clear
    For i = 1 To Nseg + 1
        If i <= Nseg Then
            .AddItem Format(i, "00") & " " & SegName(i)
            Else
            .AddItem SegName(i)
            End If
        Next i
    .ListIndex = Nseg
End With

With cmbVariable
    .Clear
    For i = 1 To NDiagnostics
        .AddItem DiagName(i)
        Next i
    .ListIndex = 1
End With

With cmbOption
    .Clear
    .AddItem "Vary Inflow Concs"
    .AddItem "Vary Flows"
    .ListIndex = 0
    End With

    ClearIt
If gRunMetaModels Then RunModels

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "Run"
     ShowStash = ShowWarnings
     ShowWarnings = False
     ClearIt
     Run_Response
     If Ier = 0 Then
        i = Wkb.Sheets("plot response").ChartObjects.Count
        If DebugMode Then MsgBox "frmresponse: N15: Creating CurrentWkChart of " & Str(i)
        DebugCount = DebugCount + 1
        'Need to Destroy the old sheet before you do this
        'CurrentWKChart is an Excel Chart Object
        'Point the object at Chart1
        Set CurrentWKChart = Wkb.Sheets("plot response").ChartObjects(1).Chart
        Fname = Directory & "temp.gif"
        CurrentWKChart.Export FileName:=Fname, FilterName:="GIF"
        'Clear any existing picture - give a little delay for export before reload
        Image1.Picture = LoadPicture()
        Image1.Picture = LoadPicture(Fname)
        Kill Fname
        reD = True
        End If
     ShowWarnings = ShowStash
     

Case "List"
     If reD Then ViewSheet ("load response")
       
Case "Metamodel" 'SEGMENT ADDED BY DMS 8/6/2008
     RunModels
            
Case "Help"
    ShowHelp (1009)

Case "Quit"
    Unload Me
    
Case "Copy Chart"
    If reD Then
        Clipboard.Clear
        Clipboard.SetData Image1.Picture
        End If
    
    End Select
'Set CurrentWKChart = DestroyObject ' WAS DEACTIVATED BY DMS

End Sub
Private Sub txtScale_Change(index As Integer)
    ClearIt
    If Not VerifyPositive(txtScale(index).Text) Then txtScale(index).Text = ""
End Sub

Private Sub txtScale_Validate(index As Integer, Cancel As Boolean)
    If txtScale(index).Text = "" Then Cancel = True
    ClearIt
End Sub
