VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bathtub Version 6.13"
   ClientHeight    =   6135
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "bath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClearOutput 
      Caption         =   "Clear Output"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      ToolTipText     =   "Clear all sheets in excel output workbook"
      Top             =   5400
      Width           =   1345
   End
   Begin VB.CommandButton btnSaveOutput 
      Caption         =   "Save Output"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Save current output workbook to a new file"
      Top             =   5400
      Width           =   1345
   End
   Begin VB.CommandButton ContinueBtn 
      BackColor       =   &H0000FF00&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaskColor       =   &H0000FF00&
      TabIndex        =   9
      ToolTipText     =   "List error messages"
      Top             =   773
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbUserMode 
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
      HelpContextID   =   196
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Select user mode (standard or advanced)"
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox cmbOutputDest 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select destination for output listings"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Run the model"
      Top             =   773
      Width           =   855
   End
   Begin VB.CommandButton btnErrorMessages 
      Caption         =   "Errors"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "List error messages"
      Top             =   773
      Width           =   1095
   End
   Begin VB.TextBox txtReport 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "bath.frx":09AA
      ToolTipText     =   "Description of current case"
      Top             =   1440
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label ExcelVersionLabel 
      Caption         =   "Excel Version:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Output Workbook:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label VersionofExcel 
      Caption         =   "<version>"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Output Destination:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblOutputWorkbook 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Name of the Excel workbook used to store output"
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label StatusLabel 
      Caption         =   "Status:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Program Activity"
      Top             =   240
      Width           =   2895
   End
   Begin VB.Menu mnuCase 
      Caption         =   "&Case"
      HelpContextID   =   17
      Begin VB.Menu mnuReadCase 
         Caption         =   "&Read"
         HelpContextID   =   17
         Begin VB.Menu mnuRead_CaseFile 
            Caption         =   "Case File (.btb)"
         End
         Begin VB.Menu mnuRead_Worksheet 
            Caption         =   "Spreadsheet (.xls)"
         End
      End
      Begin VB.Menu mnuTranslateCase 
         Caption         =   "&Translate"
         HelpContextID   =   17
      End
      Begin VB.Menu mnuSaveCase 
         Caption         =   "&Save"
         HelpContextID   =   17
      End
      Begin VB.Menu mnuSaveCaseAs 
         Caption         =   "Save &As"
         HelpContextID   =   17
      End
      Begin VB.Menu mnuNewCase 
         Caption         =   "&New"
         HelpContextID   =   17
      End
      Begin VB.Menu mnuReadDefault 
         Caption         =   "Read &Default"
         HelpContextID   =   17
      End
      Begin VB.Menu mnuSaveWorksheet 
         Caption         =   "&Save Worksheet"
         Enabled         =   0   'False
         HelpContextID   =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReadWorksheet 
         Caption         =   "&Read Worksheet"
         Enabled         =   0   'False
         HelpContextID   =   17
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   1006
      Begin VB.Menu mnuModels 
         Caption         =   "&Model Selections"
         HelpContextID   =   14
      End
      Begin VB.Menu mnuGlobals 
         Caption         =   "&Global Variables"
         HelpContextID   =   2
      End
      Begin VB.Menu mnuSegments 
         Caption         =   "&Segments"
         HelpContextID   =   3
      End
      Begin VB.Menu mnuTributaries 
         Caption         =   "&Tributaries"
         HelpContextID   =   4
      End
      Begin VB.Menu mnuLandUses 
         Caption         =   "&Export Coefficients"
         HelpContextID   =   6
      End
      Begin VB.Menu mnuChannels 
         Caption         =   "&Channels"
         HelpContextID   =   8
      End
      Begin VB.Menu mnuCoefficients 
         Caption         =   "Model &Coefficients"
         HelpContextID   =   5
      End
      Begin VB.Menu mnuWorksheet 
         Caption         =   "&Worksheet"
         HelpContextID   =   308
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      HelpContextID   =   1008
      Begin VB.Menu mnuRunModel 
         Caption         =   "&Model"
         HelpContextID   =   20
      End
      Begin VB.Menu mnuRunSensitivity 
         Caption         =   "&Sensitivity Analysis"
         HelpContextID   =   21
      End
      Begin VB.Menu mnuLoadResponse 
         Caption         =   "&Load Response"
         HelpContextID   =   1009
      End
      Begin VB.Menu mnuCalibration 
         Caption         =   "&Calibration"
         HelpContextID   =   46
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "&List"
      HelpContextID   =   53
      Begin VB.Menu mnuListInputs 
         Caption         =   "&Inputs"
         HelpContextID   =   18
      End
      Begin VB.Menu mnuListNetwork 
         Caption         =   "Segment &Network"
         HelpContextID   =   315
      End
      Begin VB.Menu mnuListHydraulics 
         Caption         =   "&Hydraulics + Morphometry"
         HelpContextID   =   23
      End
      Begin VB.Menu mnuMassBalances 
         Caption         =   "&Mass Balances"
         HelpContextID   =   24
         Begin VB.Menu mnuListGrossBalances 
            Caption         =   "&Overall"
            HelpContextID   =   24
         End
         Begin VB.Menu mnuListSegBal 
            Caption         =   "&By Segment"
            HelpContextID   =   24
         End
         Begin VB.Menu mnuListSummary 
            Caption         =   "S&ummary"
            HelpContextID   =   24
         End
         Begin VB.Menu mnuListErrors 
            Caption         =   "&Errors"
            Enabled         =   0   'False
            HelpContextID   =   310
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuListDiagnostics 
         Caption         =   "&Predicted Concentrations"
         HelpContextID   =   26
         Index           =   0
      End
      Begin VB.Menu mnuListDiagnostics 
         Caption         =   "Predicted + &Observed"
         HelpContextID   =   26
         Index           =   1
      End
      Begin VB.Menu mnuListProfiles 
         Caption         =   "Pro&files"
         HelpContextID   =   27
      End
      Begin VB.Menu mnuListTTests 
         Caption         =   "&T-Tests"
         HelpContextID   =   25
      End
      Begin VB.Menu mnuListStatistics 
         Caption         =   "&Calibration Statistics"
         HelpContextID   =   43
      End
      Begin VB.Menu mnuListErrorMessages 
         Caption         =   "&Error Messages"
         HelpContextID   =   57
      End
      Begin VB.Menu mnnListAll 
         Caption         =   "All to Workbook"
         HelpContextID   =   1007
      End
   End
   Begin VB.Menu mnuChart 
      Caption         =   "&Plot"
      HelpContextID   =   22
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "O&ptions"
      HelpContextID   =   51
      Begin VB.Menu mnuwarn 
         Caption         =   "&Warning Messages"
         Index           =   0
         Begin VB.Menu mnuwarnings 
            Caption         =   "&Show"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuwarnings 
            Caption         =   "&Don't Show"
            Index           =   1
         End
      End
      Begin VB.Menu mnuU 
         Caption         =   "&User Mode"
         HelpContextID   =   196
         Begin VB.Menu mnuUser 
            Caption         =   "&Standard"
            Checked         =   -1  'True
            HelpContextID   =   196
            Index           =   0
         End
         Begin VB.Menu mnuUser 
            Caption         =   "&Advanced"
            HelpContextID   =   196
            Index           =   1
         End
         Begin VB.Menu mnu_debugMode 
            Caption         =   "&Debug"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   51
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help Contents"
         HelpContextID   =   195
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   195
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
      HelpContextID   =   51
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public WithEvents Wko As Excel.Workbook
Private Sub btnClearOutput_Click()
'start a new output workbook
If DebugMode Then MsgBox "frmMenu N1: clearing output from btnclear"
   ClearOutput
End Sub

Private Sub btnSaveOutput_Click()
Dim outfilE As String

    On Error GoTo Abhort

        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
    outfilE = Wko.Name
    With CommonDialog1
        .FileName = outfilE
        .Filter = "Excel Files (*.xls)|*.xls|"
        .FilterIndex = 1
        .HelpContext = btnSaveOutput.HelpContextID
        .CancelError = True
        .ShowSave
        outfilE = .FileName
        End With
        
    If outfilE = "" Then GoTo Abhort
    
    Wka.WindowState = xlNormal
    Wko.Sheets(1).Activate
    Wko.SaveAs FileName:=outfilE
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
Abhort:
    FormUpdate

End Sub

Private Sub btnErrorMessages_Click()
    List_Errors
End Sub
Private Sub btnModelOptions_Click()
    SetUserMode (ListIndex)
'   mnuModels_Click
End Sub
Private Sub btnRun_Click()
    mnuRunModel_Click
End Sub
Private Sub cmbUserMode_click()
    SetUserMode (cmbUserMode.ListIndex)
    FormUpdate
'    mnuUser_Click (ListIndex)
    On Error Resume Next
    frmMenu.txtReport.SetFocus
    On Error GoTo 0
End Sub
Private Sub launchModels()
    StatList
        If Icalc > 0 Then
            ContextId = mnuLoadResponse.HelpContextID
            frmResponse.Show vbModal
            FormUpdate
            End If
     Status ("Ready")
End Sub

Private Sub ContinueBtn_Click()
If gReturnFromXLS Then
  gKeepEdits = False
  If MsgBox("Keep Edits ?", vbYesNo) = vbYes Then
   gKeepEdits = True
  End If
  gReturnFromXLS = False
End If
End Sub

Private Sub Form_Click()
If gReturnFromXLS Then
  gKeepEdits = False
  If MsgBox("Keep Edits ?", vbYesNo) = vbYes Then
   gKeepEdits = True
  End If
  gReturnFromXLS = False
End If

End Sub

Private Sub Form_Load()
    Dim Lstring As String
    Dim Lpos As Integer
      
    frmMenu.txtReport.Text = " "
    'Get the command string to see what's going on here
    Lstring = Command()
    gCase_Name = ""
    gReturnFromXLS = False
    gRunMetaModels = False
    gTASTRMode = False
    DebugMode = False
    DebugMode2 = False
    DebugMode3 = False
    DebugCount = 0
    Lstring = UCase(Lstring)
    If (Lstring = "DEBUG3") Then
    DebugMode3 = True
    End If
    If (Lstring = "DEBUG") Then
    DebugMode = True
    ElseIf Lstring <> "" Then
       gTASTRMode = True
       gCase_Name = Lstring
       Lpos = InStr(Lstring, ",")
       If Lpos > 0 Then
         gCase_Name = Left(Lstring, Lpos - 1)
         Lstring = Right(Lstring, 1)
         If Lstring = "1" Then gRunMetaModels = True
         End If
       End If
       
      
    frmMenu.Caption = "Bathtub Version " + gVersionNumber
    If gTASTRMode Then frmMenu.Caption = frmMenu.Caption & " (TASTR Mode)"
       
    'CALL THE STARTUP ROUTINE in Module 1
    StartUp
    If Ier > 0 Then
        If DebugMode Then MsgBox ("DEBUG 01 @ Unload Required" & Str(DebugCount))
        DebugCount = DebugCount + 1
        Unload Me
        Exit Sub
        End If
    If DebugMode Then MsgBox ("DEBUG 02 End of Load" & Str(DebugCount))
    DebugCount = DebugCount + 1
    Show
    If gRunMetaModels Then launchModels
        
End Sub

Sub Check_OutputDest()
   
   With frmMenu
   .cmbOutputDest.ListIndex = Iop(12)
   If Iop(12) = 2 Then
            .Label1.Visible = True
            .btnSaveOutput.Visible = True
            .btnClearOutput.Visible = True
            .lblOutputWorkbook.Visible = True
            If Wka <> nil Then Wka.Visible = True
            Else
            .Label1.Visible = False
            .btnSaveOutput.Visible = False
            .btnClearOutput.Visible = False
            .lblOutputWorkbook.Visible = False
            realversion = CDbl(Wka.Version)
            If (realversion < 15) Then Wka.WindowState = xlMinimized
   '         Wka.Visible = False
            End If
   On Error Resume Next
     .txtReport.SetFocus
   On Error GoTo 0
   End With
   
End Sub 'Check_OutputDest

Private Sub cmbOutputDest_click()
   j = Iop(12)
   Iop(12) = cmbOutputDest.ListIndex
   realversion = CDbl(Wka.Version)
   If j <> Iop(12) Then
        If j = 2 Then
            If (realversion < 15) Then Wka.WindowState = xlMinimized
            If MsgBox("Save Previous Output Workbook " & Wko.Name & " ?", vbYesNo) = vbYes Then btnSaveOutput_Click
            ClearOutput
            End If
        If Iop(12) = 2 Then ClearOutput
        If DebugMode Then MsgBox ("clear output from OutputDestClick")
        End If
      
   Check_OutputDest
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If gReturnFromXLS Then
    MsgBox "You Must Exit From EXCEL Editing FIRST" & Chr$(13) & "Please Press Bathtub Continue Button"
    Cancel = 1
    End If
End Sub

Private Sub form_terminate()
CleanUp
End Sub

Function GetOutputFileName(default As String)
    On Error GoTo Abhort
    GetOutputFileName = default
    With CommonDialog1
        .FileName = GetOutputFileName
        .Filter = "Bathtub Version 6 Files (*.btb)|*.btb|"
        .FilterIndex = 1
        .HelpContext = ContextId
        .ShowSave
        .CancelError = True
        GetOutputFileName = .FileName
        End With
        On Error GoTo 0
        Exit Function
Abhort:
    GetOutputFileName = ""
    On Error GoTo 0
End Function

Private Sub mnuListInputs_Click()
    If DebugMode Then MsgBox ("Loading Excel from MnuListInputs")
    'LoadExcel
    ClearOutput
    If Ier > 0 Then Exit Sub
    ContextId = mnuListInputs.HelpContextID
        List_Inputs
        ViewSheet ("Case Data")
End Sub

Private Sub mnnListAll_Click()
    StatList
    If Icalc > 0 Then List_All
End Sub

Private Sub mnuCalibration_Click()
    StatList
    If Icalc > 0 Then
        frmCalibration.Show vbModal
    '    Icalc = 0
        FormUpdate
        End If
End Sub
Private Sub mnuChannels_Click()
    frmChannels.Show vbModal
    Icalc = 0
    FormUpdate
End Sub

Private Sub mnuCoefficients_Click()
    frmCoefficients.Show vbModal
'    Icalc = 0
    FormUpdate
End Sub

Private Sub mnuLoadResponse_Click()
    StatList
        If Icalc > 0 Then
            ContextId = mnuLoadResponse.HelpContextID
            frmResponse.Show vbModal
            FormUpdate
            End If
     Status ("Ready")
End Sub

Private Sub mnuNewCase_Click()
    If MsgBox("Create New Case ?", vbYesNo) = vbYes Then
        n1 = InputBox("Number of Segments?")
        If n1 <= 0 Then Exit Sub
        n2 = InputBox("Number of Tributaries?")
        If n2 <= 0 Then Exit Sub
        s = InputBox("Title?")
        AllZero
        Nseg = n1
        NTrib = n2
        For i = 1 To Nseg
            SegName(i) = "Segname " & i
            If i = Nseg Then
                Iout(i) = 0
             Else
                Iout(i) = i + 1
             End If
             Next i
        For i = 1 To NTrib
            TribName(i) = "Trib " & i
            Iseg(i) = 1
            Next i
        mnuSaveCaseAs_Click
        Icalc = 0
        FormUpdate
        End If
End Sub

Private Sub mnuRead_CaseFile_Click()

Dim infilE As String
Dim outfilE As String

    On Error GoTo OUt
    infilE = "*.btb"
    ContextId = mnuReadCase.HelpContextID
    With CommonDialog1
        .FileName = infilE
        '.Filter = "Bathtub Files (*.btb)|*.btb|Version 5 Files (*.bin)|*.bin|Excel Files (*.xls)|*.xls|"
        .Filter = "Bathtub Files (*.btb)|*.btb|"
        .FilterIndex = 1
        .CancelError = True
        .HelpContext = ContextId
        .ShowOpen
        infilE = .FileName
        End With

If Not ValidFile(infilE) Or Not FileExists(infilE) Then GoTo Abhort

'MsgBox infilE
    Status ("Reading Case File...")
   
    Read_btb (infilE)
    If Ier = 0 Then Exit Sub
    
Abhort:
    MsgBox "File Not Read: " & infilE
OUt:
    Status ("Ready")
    Icalc = 0
    FormUpdate
End Sub

Private Sub mnuRead_Worksheet_Click()

Dim pFilename As String
Dim Lstring   As String
Dim outfilE   As String
    On Error GoTo Abhort2
    infilE = "*.xls"
    ContextId = mnuReadCase.HelpContextID
    With CommonDialog1
        .FileName = infilE
        .Filter = "Bathtub XLS Files (*.xls)|*.xls"
        '.Filter = "Bathtub Files (*.btb)|*.btb|"
        .FilterIndex = 1
        .CancelError = True
        .HelpContext = ContextId
        .ShowOpen
        pFilename = .FileName
        End With

If Not FileExists(pFilename) Then
   MsgBox ("Input File Not Found " & pFilename)
   GoTo Abhort2
   End If
   
Lstring = GetFileName(pFilename) & ".xls"

'MsgBox infilE
    Status ("Opening Worksheet")
'Load EXCEL and Setup XLSInputbook;

Set XLSInputApp = Nothing
Set Wka = Nothing
'Set Wka = New Excel.Application

'DEFINE EXCEL OBJECTS

Set XLSInputApp = CreateObject("Excel.Application") 'excel object for input
'NO! Wka = CreateObject("Excel.Application")
XLSInputApp.Workbooks.Open FileName:=pFilename
Set XLSWorkBk = XLSInputApp.ActiveWorkbook 'Workbooks(Lstring)

XLSInputApp.EnableEvents = True
XLSInputApp.WindowState = xlMinimized
XLSInputApp.Visible = True
MsgBox ("N304: calling Read_XLS")
Read_xls ("Inputs") 'Reads from XLSWorkBk
MsgBox ("N305: done with input, Now Closing XlsWorkBk")
MsgBox ("just a placeholder to force processing")
'XLSWorkBk.Close (savechanges = False) TEMPORARY
'XLSInputApp.Quit
'Set XLSInputApp = Nothing
'If Ier = 0 Then Exit Sub
MsgBox ("N305: Pointing WKa to XLSInputApp")
Wka = XLSInputApp
MsgBox ("N306: Jumping to Done and End Sub")
GoTo Done
    
Abhort2:
    MsgBox ("starting Abhort2 in mnuRead_Worksheet_Click")
    Lstring = "E300: Unable to Read Worksheet: ' " & pFilename & Chr$(13)
    MsgBox Lstring & "Be Sure Named Fields in Inputs Worksheet Conform to Template Bath.xla.xls"
    XLSWorkBk.Close (savechanges = False)
    Wka = XLSInputApp
    XLSInputApp.Quit
    Set XLSInputApp = Nothing
Done:
    MsgBox ("Now Starting DONE")
    Status ("Ready")
    Icalc = 0
    If Wka Is Nothing Then MsgBox ("N403: ERROR! Cannot leave here with a null WKA")
    FormUpdate
End Sub

Private Sub mnuSaveCase_Click()
    Dim LoutFile As String
    Dim Lname  As String
    Dim Msg   As String
        Lname = UCase(Directory & "DEFAULT.BTB")
        If CaseFile = "" Then mnuSaveCaseAs_Click
        Status ("Saving File...")
        LoutFile = CaseFile
        Msg = "THIS OVERWRITES THE DEFAULT CASE, PROCEED ANYWAY"
        If InStr(Lname, UCase(LoutFile)) <> 0 Then '4 is YES NO
        If MsgBox(Msg, 4, "OverWrite Default?") = vbNo Then Exit Sub
        End If
        Save_btb (LoutFile)
        FormUpdate
        Status ("Ready")
End Sub
Private Sub mnuSaveCaseAs_Click()
'save

Dim outfilE As String
On Error GoTo Abhort
    
ContextId = mnuSaveCaseAs.HelpContextID
outfilE = GetOutputFileName("*.btb")
If Not ValidFile(outfilE) Then GoTo Abhort

If FileExists(outfilE) Then
        If MsgBox("File: " & outfilE & " already exists, overwrite?", vbYesNo) <> vbYes Then GoTo Abhort
        End If
  
Status ("Saving Case...")

    Save_btb (outfilE)
    MsgBox outfilE & " saved"
    WorkingDirectory = ExtractPath(outfilE)
    CaseFile = outfilE
    FormUpdate
    Status ("Ready")
    Exit Sub

Abhort:
MsgBox "File Not Saved: " & outfilE
Status ("Ready")

End Sub
Private Sub mnuTranslateCase_Click()
'translate bin file

Dim TmpFile As String
Dim sN As String
Dim infilE As String
Dim outfilE As String
Dim WhenerroR As String
Dim InfileDir As String
Dim InfileName As String

Ier = 0

    On Error GoTo OUt
    ContextId = mnuTranslateCase.HelpContextID
    infilE = "*.bin"
    With CommonDialog1
        .FileName = infilE
        .Filter = "Bathtub Version 5 Files (*.bin)|*.bin"
        .FilterIndex = 1
        .CancelError = True
        .ShowOpen
        .HelpContext = ContextId
        infilE = .FileName
        End With
    If infilE = "" Or InStr(infilE, "*") > 0 Or Not FileExists(infilE) Then GoTo OUt
    InfileDir = ExtractPath(infilE)
    InfileName = ExtractFile(infilE)
        
    sN = infilE
    Mid(sN, Len(sN) - 2) = "btb"
    MsgBox "Now specify a new file to store translated case (*.btb)"
    outfilE = frmMenu.GetOutputFileName(sN)
    If Not ValidFile(outfilE) Then GoTo OUt
    If FileExists(outfilE) Then
        If MsgBox("File: " & outfilE & " already exists, overwrite?", vbYesNo) <> vbYes Then GoTo OUt
        End If
        
'first create temporary translation file
    'TmpFile = Directory & "bin_btb.btb"
    ChDir InfileDir
    TmpFile = "bin_btb.btb"
    If FileExists(TmpFile) Then Kill TmpFile
    sN = Directory & "convert.exe" & " " & InfileName & " " & TmpFile
'execute dos conversion program
    WhenerroR = "in Convert.exe"
    On Error GoTo Abhort
    result = Shell(sN, 0)
    On Error Resume Next
    MsgBox "Translating File: " & infilE & " to File: " & outfilE
    
'now read translated file & copy
If Not Err And FileExists(TmpFile) Then
    Read_bin_btb (TmpFile)
    WhenerroR = "in Step 2"
    If Ier > 0 Then GoTo Abhort
    If FileExists(TmpFile) Then Kill TmpFile
    Save_btb (outfilE)
    If Ier > 0 Then
        WhenerroR = "in Step 3"
        GoTo Abhort
        End If
    
    Read_btb (outfilE)
    Icalc = 0
    If Ier = 0 Then Exit Sub
    End If
        
Abhort:
    MsgBox "File Not Translated: " & WhenerroR & " " & infilE
    Icalc = 0
OUt:
    FormUpdate
    Status ("Ready")
    Err.Clear
    ChDir WorkingDirectory
    On Error GoTo 0
End Sub
Private Sub mnuAbout_Click()
   frmAbout.lblTitle(0).Caption = "Bathtub for Windows Version " + gVersionNumber
   frmAbout.Show vbModal
End Sub
Private Sub mnuChart_Click()
    StatList
    If Icalc > 0 Then frmPlot.Show
    Status ("Ready")
End Sub

Private Sub mnuGlobals_Click()
    frmGlobals.Show vbModal
'    Icalc = 0
    FormUpdate
End Sub
Private Sub mnuLandUses_Click()
    frmLandUse.Show vbModal
'    Icalc = 0
    FormUpdate
End Sub

Private Sub mnuListDiagnostics_Click(index As Integer)
'index=0 predicted index=1 both
        StatList
        If Icalc > 0 Then
        ContextId = mnuListDiagnostics(index).HelpContextID
        List_Diagnostics (index)
        ViewSheet ("Diagnostics")
        End If
End Sub

Private Sub mnuListErrorMessages_Click()
    List_Errors
   ' StatList
   ' Set gLSht = Wkb.Worksheets("errors")
   ' ViewSheet ("errors")
End Sub

Private Sub mnuListErrors_Click()
    StatList
    If Icalc > 0 Then
        ContextId = mnuListErrors.HelpContextID
        List_Verify
        ViewSheet ("Verify")
        End If
End Sub
Private Sub mnuListNetwork_Click()
            ContextId = mnuListNetwork.HelpContextID
            List_Tree
            ViewSheet ("Segment Network")
     End Sub
Private Sub mnuListGrossBalances_Click()
    StatList
        If Icalc > 0 Then
            ContextId = mnuListGrossBalances.HelpContextID
            List_GrossBalances
            ViewSheet ("Overall Balances")
            End If
End Sub

Sub StatList()
    If Icalc = 0 Then
        MsgBox ("You Must Run the Model First")
        Else
        'ContextId = HelpContextID
        ScreenOff
        If DebugMode Then MsgBox ("Loading Excel from StatList")
        'LoadExcel
        ClearOutput
'        Status ("Creating Output")
        End If
End Sub

Private Sub mnuListHydraulics_Click()
    StatList
    If Icalc > 0 Then
        ContextId = mnuListHydraulics.HelpContextID
        List_Hydraulics
        ViewSheet ("Hydraulics")
        End If
End Sub
Private Sub mnuListStatistics_Click()
    StatList
    If Icalc > 0 Then
        ContextId = mnuListStatistics.HelpContextID
        List_Fits
        ViewSheet ("Calibrations")
        End If
End Sub

Private Sub mnuModels_Click()
    Icalc = 0
    frmModels.Show vbModal
    FormUpdate
End Sub
        
Private Sub mnuListProfiles_Click()
    StatList
    If Icalc > 0 Then
    ContextId = mnuListProfiles.HelpContextID
        List_Profiles
        ViewSheet ("Profiles")
        End If
End Sub

Private Sub mnuQuit_Click()
'end of program
'    Wkb.Saved = True
    If MsgBox("Quit?", vbYesNo + vbQuestion, "End Program") = vbYes Then
        CleanUp
        Unload Me
        End If
End Sub
Private Sub mnuReadDefault_Click()
    Dim infilE As String
    infilE = Directory & "default.btb"
    Read_btb (infilE)           'read default case
End Sub

Private Sub mnuRunModel_Click()
    Run
    FormUpdate
End Sub

Private Sub mnuRunSensitivity_Click()
    StatList
    If Icalc > 0 Then
        ContextId = mnuRunSensitivity.HelpContextID
        Run_Sensitivity
        If Ier > 0 Then Exit Sub
        If Icalc > 0 Then ViewSheet ("Sensitivity")
        End If
End Sub
Private Sub mnuListSegBal_Click()
    StatList
    If Icalc > 0 Then
        ContextId = mnuListSegBal.HelpContextID
        List_SegBalances
        ViewSheet ("Segment Balances")
        End If
End Sub
Private Sub mnuListSummary_Click()
    StatList
    If Icalc > 0 Then
        ContextId = mnuListSummary.HelpContextID
        List_Terms
        ViewSheet ("Summary Balances")
        End If
End Sub

Private Sub mnuListTTests_Click()
    StatList
    If Icalc > 0 Then
    ContextId = mnuListTTests.HelpContextID
        List_TTests
        ContextId = 25
        ViewSheet ("T tests")
        End If
End Sub
Private Sub mnuSegments_Click()
    frmSegments.Show vbModal
'    Icalc = 0
    FormUpdate
End Sub

Private Sub mnuTributaries_Click()
    frmTribs.Show vbModal
'    Icalc = 0
    FormUpdate
End Sub

Private Sub mnuUser_Click(index As Integer)
'set user mode
    SetUserMode (index)
    FormUpdate
End Sub
Sub SetUserMode(index As Integer)
If index = 0 Then
    mnuUser(0).Checked = True
    mnuUser(1).Checked = False
    NoviceUser = True
    mnuListTTests.Enabled = False
    mnuWorksheet.Enabled = False
    mnuRunSensitivity.Enabled = False
    mnuCalibration.Enabled = False
    mnuListSummary.Enabled = False
    mnuListStatistics.Enabled = False
    mnuCoefficients.Enabled = False
    mnuChannels.Enabled = False
    mnuLoadResponse.Enabled = False
    mnuListProfiles.Enabled = False
    Else
    mnuUser(1).Checked = True
    mnuUser(0).Checked = False
    NoviceUser = False
    mnuListProfiles.Enabled = True
    mnuListTTests.Enabled = True
    mnuWorksheet.Enabled = True
    mnuRunSensitivity.Enabled = True
    mnuCalibration.Enabled = True
    mnuListSummary.Enabled = True
    mnuListStatistics.Enabled = True
    mnuCoefficients.Enabled = True
    mnuChannels.Enabled = True
    mnuLoadResponse.Enabled = True
    End If
End Sub

Private Sub mnuWorksheet_Click()
    ContextId = mnuWorksheet.HelpContextID
    Edit_xls
'    Icalc = 0
    FormUpdate
    Status ("Ready")
End Sub

Private Sub mnuwarnings_Click(index As Integer)
    If index = 0 Then
        ShowWarnings = True
        Else
        ShowWarnings = False
        End If
    CheckWarnings
End Sub
Sub CheckWarnings()
    If ShowWarnings Then
        mnuwarnings(0).Checked = True
        mnuwarnings(1).Checked = False
        Else
        mnuwarnings(1).Checked = True
        mnuwarnings(0).Checked = False
        End If
End Sub

Private Sub mnuHelpContents_Click()
    ShowHelp (0)
End Sub
Private Sub mnureadWorksheet_Click()
'Wka.Visible = True
' THIS IS NEVER CALLED
    Status ("Reading Worksheet")
    Read_xls ("Inputs")
    If Ier = 0 Then MsgBox ("Input Data Read from Worksheet")
    Status ("Ready")
End Sub

Private Sub mnuSaveWorksheet_Click()
'Wka.Visible = True
    Status ("Saving Worksheet")
    Save_xls
    If Ier = 0 Then MsgBox ("Input Data Saved to Worksheet")
    Status ("Ready")
End Sub

'Private Sub Wko_BeforeClose(Cancel As Boolean)
'MsgBox "you should not close this workbook"
'Cancel = True
'End Sub
Private Sub Read_CaseFile_Click()

End Sub
