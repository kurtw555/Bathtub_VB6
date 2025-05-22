Attribute VB_Name = "Module1"

' BathTub Version 6.14e
' This uses Microsoft Excel 12.0 Object Library (see References under the Project Tab)
' However 12.0 is NOT fully compatible with newer versions of Excel like 16 (see below)
' Nov 2016. cmbVariable on frmResponse form is NOT responding to user selections with
' Excel ver 16
' April 2015 Image1 in frmResponse is not getting refreshed with NEW gif file
' when user runs a new response.
' April 2013 Excel object ignores Add(after), so SheetCopy needed logic reworked,
' as did List_all

' 09/24/2011 changes to the averaging period in globals now WARNS
' 09/23/2011 substantial mods in LOADEXCEL to accomodate EXCEL's repeated plotting
' 09/09/2011 minor mods to accomodate debug
' BATHTUB TASTR Version 6.14 April 2, 2012
' LATE BINDING OF EXCEL
' The "TASTR" mode of Bathtub is determined in startup
' Specifically in the frmMenu, Form_Load() event.  Based on the Command String
' A NULL command string means NON-TASTR mode

' TO RUN IN DEBUG MODE "Bath14.exe debug"

' Excel Version 15 (see wka.version and Debug 10 and Debug 10b) has switched to SDI vs. MDI and the minimize command does not work.
' Also version 15: Wka.EnableEvents = True causes problems??



'CRITICAL NOTES:
' Late binding required some coding change because
 'references to a generic OBJECT is a little different than direct refs to specific
 

'NOTE THERE ARE ALSO PROBLEMS WITH THE WAY EXCEL DEALS WITH COLLECTIONS
'IN LATER VERSIONS - COMMENTED OUT THE .DELETE Action in "FOR EACH" Loop


' CONFIDENCE INTERVALS:  Walker assumes errors are log-normally distributed
' SEE PAGE 1-9 of the Bathtub Manual
'==============================================================
'THIS IS THE TASTR VERSION.  IT HAS BEEN MODIFIED TO ACCOMODATE
'AUTO START (input case specified as .exe parameter
'AND AUTOMATIC GENERATION OF METAMODEL OUTPUT
'SEE FORM_LOAD, gCase_Name, and gRunMetaModels
'IT ALSO INCLUDES ABILITY TO READ FROM AN XLS CASE FILE

'CONSIDER USE OF DoEvents statement as equivalent of Application.ProcessMessages
'While waiting for EXCEL to exit or edit to complete.
'====================================================================

 
 Public Const gVersionNumber As String = "6.14f (04/28/2015)"  'program version
 'Public Const Directory As String = "c:\0jobs\bathtub\"       'program directory'
 Public Directory As String                                    'bathtub.exe folder
 Public WorkingDirectory As String                             'user's working directory'
 Public Const BathtubHelpFile As String = "bathtub.chm"        'help file
 Public Const BathBook As String = "bath.xla"                  'bathtub workbook
 Public Const BathOutXLS As String = "bathtub_output.xls"      'bathtub output workbook
 Public Const BackupFile As String = "edit_backup.btx"         'backup file used to undo edits
 Public ContextId As Integer
 Public XLSWorkBk As Excel.Workbook
 'LATE BINDING APPROACH FOLLOWS
 Public Hdr  As Object 'Excel.Worksheet        'header sheet
 Public Wko  As Object 'Excel.Workbook         'output workbook used for bath_output.xls
 Public gSheetout As Object 'Excel.Worksheet        'current output sheet
 Public XLSApp As Object
 Public Wka As Object     'Excel.Application
 Public Wkb As Object     'Excel Workbook pointing at Template Bath.xla
 Public CurrentWKChart As Object  'Excel.Chart
 Public gLSht As Object 'Excel Worksheet used for Holding results
 '======================================================
 Public CaseFile As String            'name of case file
 Public TestObj As Object
 
 Public hHelp As New HTMLHelp      'html help object
 Public ErrTxt As String           'error message string
 Public NoviceUser As Boolean      'user mode
 
'constant dimensions
    Public gxla_Loaded      As Boolean
    Public gTASTRMode       As Boolean
    Public DebugMode        As Boolean    'show debug messages set in frmmenu Sub Form_Load
    Public DebugMode2       As Boolean    'Type 2 debug messages only
    Public DebugMode3       As Boolean    'type 3 debug messages only
    Public DebugCVMode      As Boolean    'debugging confidence intervals
    Public ShowWarnings     As Boolean    'show warning messages
    Public gRunMetaModels   As Boolean    'run metamodels from start
    Public gReturnFromXLS   As Boolean    'flag return from edit_xls
    Public gKeepEdits       As Boolean    'xls edits
    Public gShowEditXLSNote As Boolean    'One time flag
    
    Public DebugMsgCount As Integer       'Keeps count of debug messages sent
    Public NGlobals As Integer              'global inputs
    Public Const NCAT As Integer = 8        'land use types dimension
    Public Const NSMAX As Integer = 50      'segment dimension
    
    Public Const NPMAX As Integer = 10      'pipe dimension
    Public Const NTMAX As Integer = 100     'tributary dimension
    Public NXk As Integer
    Public NDiagnostics As Integer          'diagnostic / output  variables
    Public NtermS As Integer                'number of mass balance terms
    Public NVariables As Integer            'concentration variables in input file
    Public NOptions As Integer              'model options
    Public Const MAXIT As Integer = 100     'maximum iterations
    
    Public Ihelp As Integer
    Public Ichoice As Integer
    Public Title As String
    Public GlobalName(10) As String
    Public Cfile As String
    Public Key As String
    Public Note(10) As String
    Public Nmsg As Integer
    Public XkName(20) As String
    Public Cshort(28) As String
    Public OptionName(12, 12) As String
    Public gCase_Name As String      'added by DMS 8/04/08
    
    Public VariableName(10) As String
    Public TribName(NTMAX) As String
    Public LandUseName(NCAT) As String
    Public SegName(NSMAX + 1) As String
    Public DiagName(50) As String
    Public PipeName(NPMAX) As String
    Public CalibName(10) As String
    Public NCalib As Integer
    Public ResponseCount As Integer
    
    Public Darea(NTMAX) As Single
    Public Stat(30, 5) As Single
    Public x(500) As Single
    Public y(500) As Single
    Public Fav(8) As Single
    Public Conc(NSMAX + 1) As Single
    Public Agreg(NSMAX + 1, 6) As Single
    Public Wt(NSMAX + 1) As Single
    
    Public N_Type_Codes As Integer
    Public Type_Code(8) As String
    
'    Public Pmax(2) As Single
'    Public Qmin(2) As Single
'    Public Fmax(2) As Single
'    Public Ecod(2) As Single
'    Public Ipri(NSMAX + 1) As Integer
    Public Targ(NSMAX + 1) As Single
    Public Icrit(NSMAX + 1) As Integer
    Public Jsg(NSMAX + 1) As Integer
    Public Ecoreg(NTMAX) As Single
    Public Dareal(NTMAX) As Single
    Public Concil(NTMAX, 7) As Single
    Public CvCil(NTMAX, 7) As Single
    Public FlowL(NTMAX) As Single
    Public CvFlowL(NTMAX) As Single
    Public Wi(NSMAX + 1) As Single
    Public Q(NSMAX + 1, NSMAX + 1) As Single
    Public E(NSMAX + 1, NSMAX + 1) As Single
    Public Qt(20) As Single
    Public Bt(20) As Single
 
    Public A(NSMAX + 1, NSMAX + 1) As Single
    Public Dx(NSMAX + 1, NSMAX + 1) As Single
    Public Qnet(NSMAX + 1) As Single
    Public Xp(500) As Single
    Public Yp(500) As Single
    
    Public Warea(NTMAX, NCAT) As Single
    Public Ur(NCAT) As Single
    Public Uc(NCAT, 7) As Single
    Public CvUr(NCAT) As Single
    Public CvUc(NCAT, 7) As Single
    
    Public Qadv(NSMAX + 1) As Single
    Public Exch(NSMAX + 1) As Single
    Public Area(NSMAX + 1) As Single
    Public Zmn(NSMAX + 1) As Single
    Public Slen(NSMAX + 1) As Single
    Public Iag(NSMAX + 1) As Integer
    Public Iout(NSMAX + 1) As Integer
    Public Iseg(NTMAX) As Integer
    Public NTrib As Integer
    Public Nseg As Integer
    Public Npipe As Integer
    Public Icode(NTMAX) As Integer
    Public Iop(12) As Integer
    Public IopDefault(12) As Integer
       
    Public Iord(28) As Integer
    Public Nord As Integer

    Public Ier As Integer
    Public Izap(NSMAX + 1) As Integer
    Public Mop(12) As Integer
    Public Imap(3) As Integer

    Public Iwork(100) As Integer
    Public Ilogd(28) As Integer
    Public Icalc As Integer
    
    Public Ito(NPMAX)  As Integer
    Public Ifr(NPMAX) As Integer
    Public Icoef(4) As Integer
    Public Turbi(NSMAX + 1) As Single
    Public CvTurbi(NSMAX + 1) As Single
    Public Imsg As Integer
    Public line_no As Integer
    Public Tol As Single
    Public Sigma As Single
    
    Public Xk(20) As Single
    Public Cal(NSMAX + 1, 9) As Single
    Public Atm(7) As Single
    Public P(9) As Single                'P are from Global Vars Screen
                                         'P(1) = Avr. Period (years)
                                         'P(2) = Precip (meters)
                                         'P(3) = Evap (meters)
                                         'P(4) = Increase in Storage(m)
    Public Conci(NTMAX, 9) As Single
    Public Flow(NTMAX) As Single
    Public Cobs(NSMAX + 1, 30) As Single 'Observed conc?
    Public Zmx(NSMAX + 1) As Single    'used in calculations
    Public Zmxi(NSMAX + 1) As Single   'input
    Public CvZmxi(NSMAX + 1) As Single
    Public Zhyp(NSMAX + 1) As Single
    
    Public Turb(NSMAX + 1) As Single
    Public Qpipe(NPMAX) As Single
    Public Epipe(NPMAX) As Single
    Public CvXk(20) As Single
    Public CvCal(NSMAX + 1, 9) As Single
    Public CvAtm(7) As Single
    Public Cp(9) As Single
    Public CvCi(NTMAX, 7) As Single
    Public CvFlow(NTMAX) As Single
    Public CvCobs(NSMAX + 1, 30) As Single
    Public CvZmx(NSMAX + 1) As Single
    Public CvZhyp(NSMAX + 1) As Single
    Public InternalLoad(NSMAX + 1, 7) As Single
    Public CvInternalLoad(NSMAX + 1, 7) As Single
    
    Public CvTurb(NSMAX + 1) As Single
    Public CvQpipe(NPMAX) As Single
    Public CvEpipe(NPMAX) As Single
    Public Term(4, 20) As Single
    Public Cest(NSMAX + 1, 30) As Single
    Public CvTerm(4, 20) As Single
    Public CvCest(NSMAX + 1, 30) As Single
    
    Public XkDefault(20) As Single
    Public CvXkDefault(20) As Single
        
    Public TermName(20) As String
    Public Nkord As Integer
    Public Kord(20) As Integer
    Public Njord As Integer
    Public Jord(20) As Integer
    Public MassBalName(2) As String
    Public Mord As Integer
    Public Lord(10) As Integer
    
    Public Xe(10000) As Single     'error analysis swap vectors
    Public Cxe(10000) As Single
    Public Nxe As Integer
    Public Nxe_1 As Integer
    Public Ye(10000) As Single
    Public Cye(10000) As Single
    Public Nye As Integer
    Public Ysave(3000) As Single
   Sub Main()
    StartUp
   End Sub
   Sub StartUp()
   Dim LoadMsg As String
    'start up program
    LoadMsg = "Missing Files"
    frmAbout.lblTitle(0).Caption = "Bathtub for Windows Version " + gVersionNumber
    If Not gTASTRMode Then
      'frmSplash.Show vbModal
      'the FrmAbout.show event initializes a number of important variables
      frmAbout.Show vbModal 'Needed to show this form to get paths assigned in NON-TASTR Mode
    Else
    'TASTR MODE ONLY: the following was moved here From FrmAbout
      Directory = App.Path & Application.PathSeparator
      App.HelpFile = Directory & BathtubHelpFile
      hHelp.CHMFile = Directory & BathtubHelpFile
      End If
    '==================================================================
    '======     C H E C K    F O R   M I S S I N G    F I L E S   =====
    '==================================================================
    
    If Not FileExists(Directory & "\" & "Bath.xla") Then
       MsgBox ("BathTub Must Abort, Missing Critical File: " & Directory & "\" & "Bath.xla")
       GoTo Abhort
       End If
       
      If Not FileExists(Directory & "\" & "Default.btb") Then
       MsgBox ("BathTub Must Abort, Missing Critical File: " & Directory & "\" & "Default.btb")
       GoTo Abhort
       End If
    
    DebugCVMode = False

    Icalc = 0
    Ier = 0
    Status ("Starting Up")
    
    Set XLSInputApp = CreateObject("Excel.Application") 'excel object for input
    Set Wka = CreateObject("Excel.Application")
    Wka.DisplayAlerts = False
    gxla_Loaded = False
    'Set Wkb = CreateObject("Excel.Workbooks") we don't create these anymore here but
    'Set Wko = CreateObject("Excel.Workbooks")

    If DebugMode Then MsgBox ("Loading Excel at Startup to Test Availability")
    LoadExcel

    If DebugMode Then MsgBox ("DEBUG 03 Excel Load Attempt Completed " & Str(DebugCount))
    DebugCount = DebugCount + 1
    LoadErr = "Error Loading Excel"
    If Ier > 0 Then GoTo Abhort
    LoadErr = "Initialization Step - Output_Init "
    On Error GoTo Abhort
    Output_Init                     'initialize variables
    
    If DebugMode Then MsgBox ("DEBUG 04 @ Module 1 Read Defaults " & Str(DebugCount))
    DebugCount = DebugCount + 1
    LoadErr = "Unable to Read Defaults"
    ReadDefaults
    If Ier > 0 Then GoTo Abhort
    
    Status ("Ready")
    
    If DebugMode Then MsgBox ("DEBUG 05 @ Module1 RUN " & Str(DebugCount))
    DebugCount = DebugCount + 1
    LoadErr = "Unable to Run Model (CALCON)"
    Run
    LoadErr = "Unable to Update Main Form (FormUpdate)"
    FormUpdate
    gShowEditXLSNote = True
    Exit Sub

Abhort:
    Ier = 1
    MsgBox "Program Could Not Start: " & LoadErr & " Check Installation "
    CleanUp
    End Sub

Sub ReadDefaults()
'read default input file & assign values
    Dim i As Long
    
    ReadKey                         'read parameter key
    If Ier > 0 Then Exit Sub
    
    With frmMenu.cmbUserMode
        .Clear
        .AddItem "Standard Mode"
        .AddItem "Advanced Mode"
        .ListIndex = 1
    End With
    
    With frmMenu.cmbOutputDest
        .Clear
        For i = 1 To 3
            .AddItem OptionName(12, i)
            Next i
        .ListIndex = 1
        End With
        
    frmMenu.SetUserMode (1)
    FormUpdate
    CaseFile = gCase_Name
    If gCase_Name = "" Then CaseFile = Directory & "default.btb"
    Read_btb (CaseFile)             'read default case
    If Ier > 0 Then Exit Sub
    
'assign defaults
    For i = 1 To NOptions
        IopDefault(i) = Iop(i)
        Next i
       
    For i = 1 To NXk
        XkDefault(i) = Xk(i)
        CvXkDefault(i) = CvXk(i)
        Next i
End Sub
Sub CleanUp()
'clear all instances of excel that have bath.xls or bath_output.xls
    
    On Error Resume Next
        'MsgBox ("clearing excel")
        Wko.Saved = True
        Wko.Close (savechanges = False)
        Set Wko = Nothing
        XLSWorkBk.Close (savechanges = False)
        XLSInputApp.Close
        Wka.EnableEvents = False
        Wka.Quit
        Set Wka = Nothing
        Set Wkb = Nothing
        Set XLSWorkBk = Nothing
        Set XLSInputApp = Nothing
        hHelp.HHClose
        Set hHelp = Nothing
Quit:
    On Error GoTo 0
End Sub
Sub RefreshMainTextBox()
'fill text box on main menu screen
    Dim txt As String
    txt = "File: " & ExtractFile(CaseFile) & vbCrLf
    txt = txt & "Title: " & Title & vbCrLf
    txt = txt & "Segments: " & Nseg & "  Tributaries: " & NTrib & vbCrLf
    If Nmsg > 0 Then txt = txt & "Error Messages: " & Nmsg & vbCrLf
    If Icalc > 0 Then
        txt = txt & "Predicted Area-Weighted Means:" & vbCrLf
        If Iop(2) > 0 Then txt = txt & "---Total P(ppb) = " & Format(Cest(Nseg + 1, 2), "###0") & vbCrLf
        If Iop(3) > 0 Then txt = txt & "---Total N(ppb) = " & Format(Cest(Nseg + 1, 3), "###0") & vbCrLf
        If Iop(4) > 0 Then txt = txt & "---Chl-a(ppb)   = " & Format(Cest(Nseg + 1, 4), "###0") & vbCrLf
        If Iop(5) > 0 Then txt = txt & "---Secchi(m)    = " & Format(Cest(Nseg + 1, 5), "#0.0")
        End If
frmMenu.txtReport.Text = txt
'frmMenu.txtReport.SetFocus

'If Wka is nothing Then
'what does this REALLY want from Excel?? this LOADEXCEL is REQUIRED downstream,
'but what part?
  'If DebugMode2 Then MsgBox ("N456: RefreshMainTextBox: LoadExcel")
  'LOADEXCEL REQUIRES/ASSUMES WKA and WKB are OK
  LoadExcel
  'ClearOutput
  'End If
frmMenu.lblOutputWorkbook = Wko.Name
SegName(Nseg + 1) = "Area-Wtd Mean"
If NoviceUser Then
    frmMenu.cmbUserMode.ListIndex = 0
    Else
    frmMenu.cmbUserMode.ListIndex = 1
    End If

End Sub 'RefreshMainTextBox
Sub FormUpdate()
'update main form & menu entries
  '  On Error GoTo quitshow
    'If DebugMode2 Then MsgBox "FormUpdate"
    RefreshMainTextBox
    If Icalc = 0 Then ClearOutputWorkbook
    With frmMenu
        .CheckWarnings
        If Nmsg > 0 Then
            .btnErrorMessages.Visible = True
            Else
            .btnErrorMessages.Visible = False
            End If
    On Error Resume Next
    frmMenu.cmbOutputDest.ListIndex = Iop(12)
    On Error GoTo 0
    End With
     
    Exit Sub
Quit:
   MsgBox "Form Update Error"
   On Error GoTo 0
End Sub
   
Sub Run()
Status ("Running...")
Nmsg = 0 'error message count

If DebugMode Then MsgBox ("DEBUG 06 @ Module1 CalCon " & Str(DebugCount))
DebugCount = DebugCount + 1
Calcon
If Ier > 0 Then
    Status ("Errors")
    Icalc = 0
    Else
    Status ("Ready")
    Icalc = 1
    End If

If Nmsg > 0 And ShowWarnings Then List_Errors
   
End Sub
Sub List_Errors()
    ContextId = 196
    With frmBox
        .Caption = "Error Messages"
        .txtBox.Text = ErrTxt
        .Show vbModal
        End With
End Sub
  
    Sub Read_bin_btb(infilE As String)
'read input file created by translation utility 'convert.exe'

    Ier = 0
    On Error GoTo Abhort

    Dim junk As String

    Open infilE For Input As #1
    
    Call AllZero     'zero all input variables

'header
      Input #1, vers1
      Input #1, Title
      
'dimensions
      Input #1, junk
      Input #1, Nseg, NTrib, Npipe
      
' Parameters & Options
      Input #1, junk
       For i = 1 To NGlobals + 4
          Input #1, P(i), Cp(i)
          Next i
    
 'Model Options
    Input #1, junk
    For i = 1 To 11
      Input #1, Iop(i)
      Next i
        
'Print options  'not used
    Input #1, junk
    For i = 1 To 10
        Input #1, j
        Next i
     
'Globals
    Input #1, junk
    Xk(1) = P(6)
    CvXk(1) = Cp(6)
    For i = 1 To 12
        Input #1, Xk(i + 1), CvXk(i + 1)
        Next i

'Atmospherics
    Input #1, junk
    For i = 1 To NVariables
        Input #1, junk, Atm(i), CvAtm(i), Fav(i)
        Next i
    Xk(14) = Fav(2)
    Xk(15) = Fav(4)
    Xk(16) = Fav(3)
    Xk(17) = Fav(5)
 
'Segments
    Input #1, junk
    For i = 1 To Nseg
    Input #1, SegName(i), Iout(i), Iag(i)
    Input #1, Area(i), Zmn(i), Slen(i), Zmxi(i), CvZmxi(i), Zhyp(i), CvZhyp(i), Turbi(i), CvTurbi(i)
    For j = 1 To 9
          Input #1, Cobs(i, j), CvCobs(i, j)
          Next j
    For j = 1 To 6
        If j = 6 Then
            k = 1
            Else
            k = j + 1
            End If
        Input #1, Cal(i, k), CvCal(i, k)
        Next j
    
    Input #1, k, Icrit(i), Targ(i)
    'Input #1, x(1), x(2)
    Next i
    
' c Tribs
    Input #1, junk
    For i = 1 To NTrib
        Input #1, TribName(i), Iseg(i), Icode(i)
        Input #1, Darea(i), Flow(i), CvFlow(i)
    For j = 1 To NVariables
        Input #1, Conci(i, j), CvCi(i, j)
        Next j
    For j = 1 To NCAT
       Input #1, Warea(i, j)
       Next j
    'Input #1, Ecoreg(i)
    Input #1, x(1)
    Next i

' Channels
    Input #1, junk
    For i = 1 To Npipe
     Input #1, PipeName(i), Ifr(i), Ito(i)
     Input #1, Qpipe(i), CvQpipe(i), Epipe(i), CvEpipe(i)
     Next i

'export categories
    Input #1, junk
    For i = 1 To NCAT
      Input #1, LandUseName(i), Ur(i), CvUr(i)
      For k = 1 To NVariables
        Input #1, Uc(i, k), CvUc(i, k)
        Next k
      Next i
       
' notes
    Input #1, junk
    For i = 1 To 10
     Input #1, Note(i)
    Next i

'Allocation
'    Input #1, junk
'    For i = 1 To 2
'        Input #1, Fmax(i), Ecod(i), Qmin(i), Pmax(i)
'        Next i
'    Input #1, Ipri(1), Ipri(2), Ipri(3)
    
    Close #1
    On Error GoTo 0
    

'translate type=2 tribs
    For i = 1 To NTrib
    If Icode(i) = 2 Then
       x(1) = 0
        For k = 1 To NCAT
            x(1) = x(1) + Warea(i, k)
            Next k
        If x(1) = 0 And Flow(i) > 0 Then
           Icode(i) = 1
           MsgBox ("Type Code for Trib " & i & " " & TribName(i) & " Changed from 2 to 1")
           End If
           End If
           Next i

'translate internal loads
    For i = NTrib To 1 Step -1
        If Icode(i) = 5 Then
                For j = 1 To NVariables
                    InternalLoad(Iseg(i), j) = Conci(i, j)
                    CvInternalLoad(Iseg(i), j) = CvCi(i, j)
                    Next j
                Call TribEdit(i, 0)
                MsgBox "Tributary " & i & " " & TribName(i) & " Handled as Internal Load"
                End If
               Next i
        For i = 1 To NTrib
            If Icode(i) = 6 Then Icode(i) = 5
            Next i

'translate mixed layer depths
    For i = 1 To Nseg
        If Zmxi(i) <= 0 Then
            Zmxi(i) = ZmixEst(Zmn(i))
            CvZmx(i) = 0.12
            End If
        Next i
          
'translate non-algal turbidity
    For i = 1 To Nseg
        If Turbi(i) <= 0 Then
            Call TurbEst(Cobs(i, 4), CvCobs(i, 4), Cobs(i, 5), CvCobs(i, 5), x(1), x(2))
            Turbi(i) = x(1)
            CvTurbi(i) = x(2)
            End If
        Next i
                    
    Exit Sub
    
Abhort:
    Close #1
    Ier = 1
    MsgBox "Error in Translating File "
    On Error GoTo 0
    End Sub
    Sub Read_btb(infilE As String)
'read standard bathtub input file
    
    Ier = 0
    On Error GoTo Abhort

    Dim junk As String

'    MsgBox "Reading: " & infilE
    Open infilE For Input As #1
    
    Call AllZero     'zero all input variables

'header
    Input #1, vers1
    Line Input #1, Title
      
' Parameters & Options
    Input #1, j, junk
      For i = 1 To NGlobals
          Input #1, j, junk, P(i), Cp(i)
          Next i
 
 'Model Options
    Input #1, j, junk
    For i = 1 To NOptions
      Input #1, j, junk, Iop(i)
      Next i
        
'Globals
    Input #1, j, junk
    For i = 1 To NXk
        Input #1, j, junk, Xk(i), CvXk(i)
        Next i

'Atmospherics
    Input #1, j, junk
    For i = 1 To NVariables
        Input #1, j, junk, Atm(i), CvAtm(i)
        Next i
 
'Segments
    Input #1, Nseg, junk
    For i = 1 To Nseg
    Input #1, j, SegName(i), Iout(i), Iag(i), Area(i), Zmn(i), Slen(i), Zmxi(i), CvZmxi(i), Zhyp(i), CvZhyp(i), Turbi(i), CvTurbi(i), Icrit(i), Targ(i)
    For j = 1 To 3
        Input #1, k, junk, InternalLoad(i, j), CvInternalLoad(i, j)
        Next j
    For j = 1 To 9
          Input #1, k, junk, Cobs(i, j), CvCobs(i, j), Cal(i, j), CvCal(i, j)
          Next j
    Next i

' c Tribs
    Input #1, NTrib, junk
    For i = 1 To NTrib
        Input #1, j, TribName(i), Iseg(i), Icode(i), Darea(i), Flow(i), CvFlow(i), Ecoreg(i)
    For j = 1 To NVariables
        Input #1, k, junk, Conci(i, j), CvCi(i, j)
        Next j
    
    Input #1, k, junk, Warea(i, 1), Warea(i, 2), Warea(i, 3), Warea(i, 4), Warea(i, 5), Warea(i, 6), Warea(i, 7), Warea(i, 8)
    Next i

' Channels
    Input #1, Npipe, junk
    For i = 1 To Npipe
     Input #1, j, PipeName(i), Ifr(i), Ito(i), Qpipe(i), CvQpipe(i), Epipe(i), CvEpipe(i)
     Next i

'export categories
    Input #1, j, junk
    For i = 1 To NCAT
      Input #1, j, LandUseName(i)
      Input #1, j, junk, Ur(i), CvUr(i)
      For k = 1 To NVariables
        Input #1, j, junk, Uc(i, k), CvUc(i, k)
        Next k
      Next i
       
' notes
    Input #1, junk
    For i = 1 To 10
     Line Input #1, Note(i)
    Next i

'Allocation
'    Input #1, junk
'    For i = 1 To 2
'        Input #1, Fmax(i), Ecod(i), Qmin(i), Pmax(i)
'        Next i
'    Input #1, Ipri(1), Ipri(2), Ipri(3)
    
    Close #1
    
    On Error GoTo 0
    'MsgBox "output destination " & OptionName(12, Iop(12) + 1)
    CaseFile = infilE
    WorkingDirectory = ExtractPath(infilE)
    Icalc = 0
    If DebugMode Then MsgBox ("N3: clear output from readbtb in Module1")
    ClearOutput
'    On Error Resume Next 'next statement will fail at startup
    'frmMenu.cmbOutputDest.ListIndex = Iop(12)
    frmMenu.Check_OutputDest
    Status ("Ready")
    
    On Error GoTo 0
    Exit Sub
    
Abhort:
    Close #1
    Ier = 1
    MsgBox "Input File Error", vbCritical
    On Error GoTo 0
    End Sub
    
    
    Sub Save_btb(Lout As String)

'writes output file to disk in ascii format
    
    Dim junk As String
    
'    MsgBox "saving : " & outfilE
    Open Lout For Output As #1

'header
    Print #1, "Vers " & gVersionNumber
    Print #1, Title
      
' Parameters & Options
      Write #1, NGlobals, "Global Parmameters"
       For i = 1 To NGlobals
          Write #1, i, GlobalName(i), P(i), Cp(i)
          Next i
 
 'Model Options
    Write #1, NOptions, "Model Options"
    For i = 1 To NOptions
      Write #1, i, OptionName(i, 0), Iop(i)
      Next i
  
 'Print options
  '  Write #1, "Print Options"
  '  For i = 1 To 10
  '      Write #1, j
  '      Next i

'Globals
    Write #1, NXk, "Model Coefficients"
    For i = 1 To NXk
        Write #1, i, XkName(i), Xk(i), CvXk(i)
        Next i

'Atmospherics
    Write #1, NVariables, "Atmospheric Loads"
    For i = 1 To NVariables
        Write #1, i, VariableName(i), Atm(i), CvAtm(i)
        Next i
 
'Segments
    Write #1, Nseg, "Segments"
    For i = 1 To Nseg
    Write #1, i, SegName(i), Iout(i), Iag(i), Area(i), Zmn(i), Slen(i), Zmxi(i), CvZmxi(i), Zhyp(i), CvZhyp(i), Turbi(i), CvTurbi(i), Icrit(i), Targ(i)
    For j = 1 To 3
        Write #1, i, VariableName(j), InternalLoad(i, j), CvInternalLoad(i, j)
        Next j
    For j = 1 To 9
      Write #1, i, DiagName(j), Cobs(i, j), CvCobs(i, j), Cal(i, j), CvCal(i, j)
      Next j
    Next i

' c Tribs
    Write #1, NTrib, "Tributaries"
    For i = 1 To NTrib
        Write #1, i, TribName(i), Iseg(i), Icode(i), Darea(i), Flow(i), CvFlow(i), Ecoreg(i)
    For j = 1 To NVariables
        Write #1, i, VariableName(j), Conci(i, j), CvCi(i, j)
        Next j
    Write #1, i, "LandUses", Warea(i, 1), Warea(i, 2), Warea(i, 3), Warea(i, 4), Warea(i, 5), Warea(i, 6), Warea(i, 7), Warea(i, 8)
    Next i

' Channels
    Write #1, Npipe, "Channels"
    For i = 1 To Npipe
     Write #1, i, PipeName(i), Ifr(i), Ito(i), Qpipe(i), CvQpipe(i), Epipe(i), CvEpipe(i)
     Next i

'export categories
    Write #1, NCAT, "Land Use Export Categories"
    For i = 1 To NCAT
      Write #1, i, LandUseName(i)
      Write #1, i, "Runoff", Ur(i), CvUr(i)
      For k = 1 To NVariables
        Write #1, i, VariableName(k), Uc(i, k), CvUc(i, k)
        Next k
      Next i
       
' notes
    Write #1, "Notes"
    For i = 1 To 10
        Print #1, Note(i)
    Next i

'Allocation
'    Write #1, "Allocation"
'    For i = 1 To 2
'        Write #1, Fmax(i), Ecod(i), Qmin(i), Pmax(i)
'        Next i
'    Write #1, Ipri(1), Ipri(2), Ipri(3)
    Close #1
  
    End Sub
Sub Edit_xls()
'edit inputs on worksheet
    Dim Lstring As String
    
    Status ("Loading Worksheet")
    Ier = 0
    'If DebugMode Then
    MsgBox ("Loading Excel from sub Edit_xls")
    LoadExcel
    If Ier > 0 Then Exit Sub
    Set gLSht = Wkb.Worksheets("Inputs")
    Save_xls              'writes current inputs to "inputs" sheet of WKB
    If Ier > 0 Then Exit Sub
    
    'Copy the filled-in template to the WKO workbook
    Wkb.Worksheets("Inputs").Copy before:=Wko.Sheets(1)
    Set gSheetout = Wko.Worksheets("Inputs")
    
    If Ier > 0 Then Exit Sub
    Status ("Edit Worksheet")
    frmMenu.ContinueBtn.Visible = True
    Lstring = "Click On the Bathtub <CONTINUE> Button When Done Editing in EXCEL" & Chr$(13)
    Lstring = Lstring & "          When Done Editing, No Worksheet Cells Can Be Active."
    If gShowEditXLSNote Then
    MsgBox Lstring & Chr$(13) & "                     DO NOT CLOSE EXCEL MANUALLY"
    End If
    gShowEditXLSNote = False
    'frmMenu.WindowState = vbMinimized
    Wka.WindowState = xlNormal
    gSheetout.Range("C3").Select
    gReturnFromXLS = True
    Do While gReturnFromXLS
    DoEvents
    Loop
    frmMenu.ContinueBtn.Visible = False
   ' After user edits are done
    
   If gKeepEdits Then
       On Error GoTo Quit   'in case user closes window
       gSheetout.UsedRange.Copy gLSht.Range("a1")
       Set XLSWorkBk = Wko
       Read_xls ("Inputs")
       End If
Wka.DisplayAlerts = False
gSheetout.Delete
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
Wka.DisplayAlerts = True
Quit:
    Set XLSWorkBk = Nothing
    Status ("Ready")
    End Sub


Sub Save_xls()

'writes input data to excel sheet without saving to disk
'first update main screen

On Error GoTo Abhort

With gLSht 'gLSht is defined as Wkb.Worksheets("Inputs") in Edit_xls

.Range("c3:HA1000").ClearContents

'header
'[Version] = gVersionNumber
.Range("title").Value = Title

'dimensions
.Range("Nseg").Value = Nseg
.Range("Ntrib").Value = NTrib
.Range("Npipe").Value = Npipe

' Parameters & Options
With .Range("global_factors")
   For i = 1 To 4
        .Offset(i, 2) = P(i)
        .Offset(i, 3) = Cp(i)
        Next i
    End With

'Model Options
With .Range("model_options")
For i = 1 To NOptions
    .Offset(i, 2) = Iop(i)
    Next i
   End With
   
'Globals
    With .Range("calibration_factors")
    For i = 1 To NXk
        .Offset(i, 2).Value = Xk(i)
        .Offset(i, 3).Value = CvXk(i)
        Next i
        End With

'Atmospherics
    With .Range("atmos_loads")
    For i = 1 To NVariables
 '       .Offset(i, 1).Value = VariableName(i)
        .Offset(i, 2).Value = Atm(i)
        .Offset(i, 3).Value = CvAtm(i)
         Next i
        End With
    
 'Segments
    With .Range("segment_data").Offset(0, 1)
    For i = 1 To Nseg
    .Offset(1, i) = i
    .Offset(2, i) = SegName(i)
    .Offset(3, i) = Iout(i)
    .Offset(4, i) = Iag(i)
    .Offset(6, i) = Area(i)
    .Offset(7, i) = Zmn(i)
    .Offset(8, i) = Slen(i)
    .Offset(9, i) = Zmxi(i)
    .Offset(10, i) = CvZmxi(i)
    .Offset(11, i) = Zhyp(i)
    .Offset(12, i) = CvZhyp(i)
    .Offset(14, i) = Turbi(i)
    .Offset(15, i) = CvTurbi(i)
    
   For j = 1 To 9
        .Offset(14 + j * 2, i) = Cobs(i, j)
        .Offset(14 + j * 2 + 1, i) = CvCobs(i, j)
        Next j
        
    For j = 1 To 9
        .Offset(33 + j * 2, i) = Cal(i, j)
        .Offset(33 + j * 2 + 1, i) = CvCal(i, j)
         Next j
    
    For j = 1 To 3
        .Offset(53 + j * 2, i) = InternalLoad(i, j)
        .Offset(53 + j * 2 + 1, i) = CvInternalLoad(i, j)
        Next j
     
    'Write #1, k, Icrit(i), Targ(i)
    Next i
    End With
    
' c Tribs
    With .Range("Tributary_data").Offset(0, 1)
    For i = 1 To NTrib
    .Offset(1, i) = i
    .Offset(2, i) = TribName(i)
    .Offset(3, i) = Iseg(i)
    .Offset(4, i) = Icode(i)
    .Offset(5, i) = Darea(i)
    .Offset(6, i) = Flow(i)
    .Offset(7, i) = CvFlow(i)
    
    For j = 1 To NVariables
        .Offset(6 + 2 * j, i) = Conci(i, j)
        .Offset(6 + 2 * j + 1, i) = CvCi(i, j)
        Next j
       
    For j = 1 To NCAT
        .Offset(18 + j, i) = Warea(i, j)
       Next j
  
    Next i
    End With

' Channels
    With .Range("transport_channels").Offset(0, 1)
    For i = 1 To Npipe
        .Offset(1, i) = i
        .Offset(2, i) = PipeName(i)
        .Offset(3, i) = Ifr(i)
        .Offset(4, i) = Ito(i)
        .Offset(5, i) = Qpipe(i)
        .Offset(6, i) = CvQpipe(i)
        .Offset(7, i) = Epipe(i)
        .Offset(8, i) = CvEpipe(i)
        Next i
     End With

'export categories
    With .Range("categories").Offset(0, 1)
    For i = 1 To NCAT
       .Offset(1, i) = i
       .Offset(2, i) = LandUseName(i)
       .Offset(3, i) = Ur(i)
       .Offset(4, i) = CvUr(i)
       
      'Write #1, LandUseName(i), Ur(i), CvUr(i)
      For k = 1 To NVariables
        .Offset(3 + k * 2, i) = Uc(i, k)
        .Offset(3 + k * 2 + 1, i) = CvUc(i, k)
       ' Write #1, Uc(i, k), CvUc(i, k)
        Next k
      Next i
      End With

' notes
    With .Range("notes")
    For i = 1 To 10
        .Offset(i - 1, 0) = Note(i)
        Next i
    End With

End With

'Allocation
'    Write #1, "Allocation"
'    For i = 1 To 2
'        Write #1, Fmax(i), Ecod(i), Qmin(i), Pmax(i)
'        Next i
'    Write #1, Ipri(1), Ipri(2), Ipri(3)
    
'    Close #1
'   MsgBox "Input Data Copied to Bathtub worksheet"
    
'copy
    Exit Sub

Abhort:
    MsgBox "Error Writing Input Data to Excel Workbook"
    Ier = 1
    
    End Sub
Sub Read_xls(pSheetName)

'PROGRAM OVERALL APPROACH
' 1. Load bath.xla template
' 2. Read heading/label material from named ranges in bath.xla
' 3. ?? WKB disappears when done... so template is unaltered.

' NOTE THIS SUB USES GLOBAL XLSWORKBK - I'm unable to pass as an argument.


'reads input data from excel sheet "WKB" formated according to BATH.XLA

On Error GoTo Abhort
Ier = 0
Call AllZero


With XLSWorkBk.Sheets(pSheetName)
'header
'NAMED ranges MUST BE DEFINED in BATH.XLA aka WBa??
MsgBox ("Read_xls: 500: Reading Ranges from " & pSheetName)
Versx = .Range("version")
Title = .Range("title")

'dimensions
Nseg = .Range("Nseg")
NTrib = .Range("Ntrib")
Npipe = .Range("Npipe")

' Parameters & Options
With .Range("global_factors")
   For i = 1 To 4
        P(i) = .Offset(i, 2)
        Cp(i) = .Offset(i, 3)
        Next i
    End With

'Model Options
With .Range("model_options")
For i = 1 To NOptions
    Iop(i) = .Offset(i, 2)
    Next i
   End With
   
'Globals
    With .Range("calibration_factors")
    For i = 1 To NXk
        'XkName(i) = .Offset(i, 0)
        Xk(i) = .Offset(i, 2)
        CvXk(i) = .Offset(i, 3)
        Next i
        End With

'Atmospherics
    With .Range("atmos_loads")
    For i = 1 To NVariables
        'VariableName(i) = .Offset(i, 0)
        Atm(i) = .Offset(i, 2)
        CvAtm(i) = .Offset(i, 3)
         Next i
        End With
        
 'Segments
    With .Range("segment_data").Offset(0, 1)
    For i = 1 To Nseg
    'i = .Offset(1, i)
        SegName(i) = .Offset(2, i)
        Iout(i) = .Offset(3, i)
        Iag(i) = .Offset(4, i)
        Area(i) = .Offset(6, i)
        Zmn(i) = .Offset(7, i)
        Slen(i) = .Offset(8, i)
        Zmxi(i) = .Offset(9, i)
        CvZmxi(i) = .Offset(10, i)
        Zhyp(i) = .Offset(11, i)
        CvZhyp(i) = .Offset(12, i)
        Turbi(i) = .Offset(14, i)
        CvTurbi(i) = .Offset(15, i)
    
   For j = 1 To 9
        Cobs(i, j) = .Offset(14 + j * 2, i)
        CvCobs(i, j) = .Offset(14 + j * 2 + 1, i)
        Next j
        
    For j = 1 To 9
        Cal(i, j) = .Offset(33 + j * 2, i)
        CvCal(i, j) = .Offset(33 + j * 2 + 1, i)
         Next j
     
    For j = 1 To 3
        InternalLoad(i, j) = .Offset(53 + j * 2, i)
        CvInternalLoad(i, j) = .Offset(53 + j * 2 + 1, i)
        Next j
    
    'Write #1, k, Icrit(i), Targ(i)
    Next i
    
    End With
    SegName(Nseg + 1) = "AREA-WTD MEAN"

' c Tribs
    With .Range("Tributary_data").Offset(0, 1)
    For i = 1 To NTrib
        TribName(i) = .Offset(2, i)
        Iseg(i) = .Offset(3, i)
        Icode(i) = .Offset(4, i)
        Darea(i) = .Offset(5, i)
        Flow(i) = .Offset(6, i)
        CvFlow(i) = .Offset(7, i)
    
    For j = 1 To NVariables
        Conci(i, j) = .Offset(6 + 2 * j, i)
        CvCi(i, j) = .Offset(6 + 2 * j + 1, i)
        Next j
       
    For j = 1 To NCAT
       Warea(i, j) = .Offset(18 + j, i)
       Next j
     
 '   Write #1, Ecoreg(i)
    Next i
    End With

' Channels
    With .Range("transport_channels").Offset(0, 1)
    For i = 1 To Npipe
    '   i = .Offset(1, i)
        PipeName(i) = .Offset(2, i)
        Ifr(i) = .Offset(3, i)
        Ito(i) = .Offset(4, i)
        Qpipe(i) = .Offset(5, i)
        CvQpipe(i) = .Offset(6, i)
        Epipe(i) = .Offset(7, i)
        CvEpipe(i) = .Offset(8, i)
        Next i
     End With

'export categories
    With .Range("categories").Offset(0, 1)
    For i = 1 To NCAT
        i = .Offset(1, i)
       LandUseName(i) = .Offset(2, i)
       Ur(i) = .Offset(3, i)
       CvUr(i) = .Offset(4, i)

      For k = 1 To NVariables
        Uc(i, k) = .Offset(3 + k * 2, i)
        CvUc(i, k) = .Offset(3 + k * 2 + 1, i)
        Next k
      Next i
      End With

' notes
    With .Range("notes")
    For i = 1 To 10
        Note(i) = .Offset(i - 1, 0)
        Next i
    End With

MsgBox ("Read_xls: 501: Done Reading Ranges from " & pSheetName)
'Allocation
'    Write #1, "Allocation"
'    For i = 1 To 2
'        Write #1, Fmax(i), Ecod(i), Qmin(i), Pmax(i)
'        Next i
'    Write #1, Ipri(1), Ipri(2), Ipri(3)
    
'    Close #1
    'MsgBox "Input Data Saved"
    
    End With

    Exit Sub

Abhort:
    MsgBox "Error Reading Worksheet - Remember Named Ranges Required"
    Ier = 1
        
    End Sub
Sub ReadKey()

'c read key file
    Ier = 0
    With Wkb.Sheets("key")
      
    Sigma = .Range("sigma")      'number of standard errors plotted around predicted & observed values
    Tol = .Range("tolerance")    'tolerance for convergence of mass balance solution
    
'diagnostic variables
    NDiagnostics = .Range("nDiagnostics")
    Nord = NDiagnostics
    With .Range("ndiagnostics").Offset(1, 0)
    For i = 1 To NDiagnostics
        j = .Offset(i, 0)   'variable number
        Iord(i) = j
        Ilogd(j) = .Offset(i, 1)
        Cshort(j) = .Offset(i, 2)
        DiagName(j) = .Offset(i, 3)
        For k = 1 To 5
            Stat(j, k) = .Offset(i, k + 3)
            Next k
        Next i
    End With

NOptions = .Range("Noptions")
With .Range("noptions").Offset(1, 0)
    k = 0
    For i = 1 To NOptions
        k = k + 1
        Mop(i) = .Offset(k, 0)       'number of options for
        OptionName(i, 0) = .Offset(k, 1) 'option name
        IopDefault(i) = 0
        For j = 1 To Mop(i)
            k = k + 1
            OptionName(i, j) = .Offset(k, 1) 'label for selection
            If .Offset(k, 2) > 0 Then IopDefault(i) = j - 1
            Next j
        k = k + 1
        Next i
     End With
        
'coefficient labels
NXk = .Range("ncoef")
   With .Range("ncoef").Offset(1, 0)
   For i = 1 To NXk
        XkName(i) = .Offset(i, 1)
        XkDefault(i) = .Offset(i, 2)
        CvXkDefault(i) = .Offset(i, 3)
        Next i
   End With
        
   End With
   Exit Sub
        
s999:
       MsgBox "Invalid Key File"
       Ier = 1
      
       End Sub
Sub ScreenOff()
    Application.ScreenUpdating = False
End Sub
Sub ScreenOn()
    Application.ScreenUpdating = True
End Sub
Function ValidFile(s As String) As Boolean

If s = "" Or InStr(s, "*") > 0 Or InStr(UCase(s), ".BTB") = 0 Then
    ValidFile = False
    Else
    ValidFile = True
    End If

End Function

'Sub restric()
'c restrict output segments
         'If (Nseg <= 0) Then Exit Sub
         'Dummy = "INCLUDED <*> or EXCLUDED < >"
         'Call swint(Izap, Iwork, Nseg + 1)
         'Call SelSeg(Nseg + 1)
         'Call swint(Iwork, Izap, Nseg + 1)
         'Call clr(0)
'         End Sub

Sub ShowHelp(ctxtiD)
'0=contents, >0=contexid
'uses special class module downloaded from internet ----
On Error GoTo Abhort

'.CHMFile = Directory & BathtubHelpFile
'.HHWindow = "main"
    If ctxtiD <= 0 Then
        hHelp.HHDisplayContents
    Else
        hHelp.HHTopicID = ctxtiD
        hHelp.HHDisplayTopicID
    End If

Abhort:
End Sub
Sub ListOneSheet(io)
    Dim sN As String
    sN = [sheet_selected]
    Call ViewSheet(sN)
    Sheets("menu").Activate
End Sub
Sub Status(Msg As String)
'update status box
    If Msg = "Ready" Then
    frmMenu.lblStatus.BackColor = &H8000000B
    Else
    frmMenu.lblStatus.BackColor = &H8000000E
    End If
    frmMenu.lblStatus.Caption = Msg
End Sub

Sub ViewFileTextBox(fna As String)
'view a file as a text box (outputdest=2)
Dim txt As String
Dim Lstring As String

If FileExists(fna) Then
    Open fna For Input As #1
    txt = Input(LOF(1), 1)
   ' Input #1, txt
  '  Do Until EOF(1)
  '     Line Input #1, lstring
  '      wmax = MAx(wmax, Len(lstring))
  '      txt = txt & lstring & vbCrLf
  '      i = i + 1
   '     Loop
    Close #1
    
    With frmBox
    .Caption = gLSht.Name
    .txtBox.Text = txt
    .txtBox.SelStart = 0
    .Show vbModal
    End With
    
    End If
End Sub 'viewfiletextbox

'Sub ViewSheetRichTextBox(ShtName)
'view a file as a rich text box (outputdest=2)
           
'    gLSht.UsedRange.Copy
'    frmRtb.Show

    'With frmRtb
    '.rtbOutput.Text = txT
    '.rtbOutput.SelStart = 0
    '.Show
    'End With
    
'End Sub
Sub ViewFileNotepad(fna As String)
'view a text file in notepad (outputdest=1)
    If FileExists(fna) Then
        cline = "Notepad " & fna
        j = Shell(cline, 1)
        End If
End Sub

'Sub ViewSheetImage(ShtName As String)
'view sheet as image not ready
'    gLSht.UsedRange.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
'    frmPicture.Show
'End Sub
Function VerifyNumber(s) As Boolean
If IsNumeric(s) Or s = "." Or s = "-" Then
    VerifyNumber = True
    Else
    VerifyNumber = False
    End If
End Function
Function VerifyPositive(s) As Boolean
VerifyPositive = VerifyNumber(s)
If VerifyPositive Then
    If Val(s) < 0 Then VerifyPositive = False
    End If
End Function
Sub ViewSheet(SheetName As String)
   'Copy gLSHT (aka wkb.sheet (activesheet?)) to WKO.Sheetname
   'WKO has been set to wka.activeWorkbook
   SheetCopy (SheetName)
    
   If Iop(12) = 2 Then
     If Not gRunMetaModels Then
       Wka.Visible = True
       Wka.WindowState = xlNormal
       End If
   Else
   Dim fworK As String
   fworK = Directory & SheetName & ".prn"
   If FileExists(fworK) Then Kill fworK
   Wko.Save
   Wko.SaveAs FileName:=fworK, FileFormat:=xlTextPrinter, CreateBackup:=False
   Wko.Saved = True
   Wko.Close (savechanges = False)
   Wka.Workbooks.Open FileName:=Directory & BathOutXLS
   Set Wko = Wka.ActiveWorkbook
   
   If Iop(12) = 1 Then
        Call ViewFileNotepad(fworK)
        Else
        Call ViewFileTextBox(fworK)
        End If
    Kill fworK
    End If

Status ("Ready")
End Sub 'ViewSheet

Sub SheetCopy(SheetName As String)
  'this is supposed to copy Sheetname from X to WKo (gSheetOut)
   On Error Resume Next 'restart after error stmt
   'WKO is the Bathtub_output.xls workbook
   With Wko
    .Worksheets("Sheet1").Delete
    .Worksheets("Sheet2").Delete
    Set gSheetout = .Worksheets(SheetName)
    .Worksheets(SheetName).Activate
    If Err > 0 Then
        'MsgBox ("There is no Sheet in WKo called " & SheetName)
        .Worksheets.Add 'Create a New Sheet. The (AFTER) option kills the action entirely?
        .Sheets(1).Name = SheetName 'was .activesheet.name Give it the desired name
        Set gSheetout = .Sheets(1) '.ActiveSheet       'Use it
        .Worksheets("sheet3").Delete 'now this can be dumped
        Wka.ActiveWindow.DisplayGridlines = False
        End If
   End With
   'On error, bail out and let the caller handle any errors
   On Error GoTo 0 'Exit
   gSheetout.Cells.Clear 'Start with a Clean Output Sheet
    
    
  'WKB is the basic template called Bath.xla
  'MsgBox ("WKB is " & Wkb.Name)
  Wkb.Worksheets(SheetName).Activate
  If Err > 0 Then
     MsgBox ("PgmErr 22: " & SheetName & " Is not in the Template?")
     End If
  Set gLSht = Wkb.Worksheets(SheetName)
  'gLSht.Activate
  'MsgBox "Glsht is now wkb." & gLSht.Name
  'Wka.Workbooks("bath.xla.xls").Activate
  '  Sheets("Inputs").Copy Before:=Workbooks("Book1").Sheets(1)

'NOTE the global glSHT is set generally to Wkb.Workbooks("<sheetname>")

   'COPY THE Filled in Template file to output: WKO.<gSheetOut.name>
   'THIS COPY APPROACH LOOSES THE NAMED RANGES ???
   'MsgBox ("N1: GLSHT NAME IS " & gLSht.Name)
   With gLSht 'glSHT is a "filled in" Bath.Xla template
     ' .Range(.Range("A1"), .Range("A1").SpecialCells(xlCellTypeLastCell)).Name = "Print_Area"
     'SetPrintArea
     j = .UsedRange.Columns.Count
     'Save the Column Widths from the Template A column for use below
     For i = 1 To j
         x(i) = .Range("A1").Offset(0, i - 1).ColumnWidth
         Next i
    '.Range("print_area").Copy gSheetout.Range("a1") 'not sure what this does
    .UsedRange.Copy gSheetout.Range("a1") 'Copy to gsheetout @ A1
   End With 'glSHT
    
   gSheetout.Activate
   'Reset the Column Widths
   For i = 1 To j
       gSheetout.Range("a1").Offset(0, i - 1).ColumnWidth = x(i)
       Next i
   gSheetout.Range("A1").Select
   
   Wka.CutCopyMode = False
   ScreenOn
End Sub
'End of SheetCopy

'Sub RestartExcel()
'Ier = 0
'    On Error GoTo Abhort
'        Set Wko = Nothing
'        Set Wkb = Nothing
'        Set Wka = Nothing
'        Set Wka = New Excel.Application
'        Wka.Workbooks.Open FileName:=Directory & BathBook
'        Set Wkb = Wka.Workbooks(BathBook)
'        Set Hdr = Wkb.Sheets("headers")   'table headings sheet'
'        Wka.WindowState = xlMinimized
'        Wka.Workbooks.Add
'        Kill Directory & BathOutXLS
'        Wka.ActiveWorkbook.SaveAs FileName:=Directory & BathOutXLS
'        Set Wko = Wka.ActiveWorkbook       'bathtub_output.xls workbook
'        Exit Sub
'Abhort:
'    Ier = 1
'    MsgBox "Could not load Excel - can't continue"
'    Err.Clear
'End Sub


Sub LoadXLSInputApp(pFilename As String)
    Ier = 0
    On Error Resume Next
'
    'Err.Clear
    'On Error GoTo Abhort
    
    Set XLSInputApp = Nothing
    MsgBox ("load xlsinputapp")
    Set XLSInputApp = New Excel.Application
    XLSInputApp.Workbooks.Open FileName:=pFilename
    'Set XLSInputBook = XLSInputApp.Workbooks(pBookName)
    Set XLSInputSheet = XLSInputApp.Sheets(0)
   

'Abhort:
    'MsgBox "Could not load Excel (can't continue)- try closing Excel then restarting BATHTUB"
    'Ier = 1
    'Err.Clear
    
End Sub


Sub LoadExcel()
    'Dim LWrkbooks As Workbooks
    Dim Lwb As Workbook
    Ier = 0
    On Error Resume Next

    If DebugMode Then
      i = Wkb.Worksheets.Count
      MsgBox ("Wkb worksheets.count=" & Str(i))
      End If
    
    If Err = 0 And i > 0 Then 'do nothing
    Else
        Err.Clear
        On Error GoTo Abhort
            
        ' NO! DEFINE NEW EXCEL OBJECTS
        ' XLSInputApp is now set in main ONCE
        ' Set XLSInputApp = CreateObject("Excel.Application") 'excel object for input
        'NO! Set Wka = Nothing
        'NO! Set Wka = CreateObject("Excel.Application")
        'But we must CLEAR wka (close it) if it is open.
        If DebugMode Then MsgBox "DEBUG 07 Excel Version: " & Wka.Version & " " & Str(DebugCount)
        DebugCount = DebugCount + 1
        If (Wka Is Nothing) Then MsgBox ("N399: Fatal Error, Input Worksheet does not exist")
        frmMenu.VersionofExcel.Caption = Wka.Version
        'If Wka.Workbooks.Count > 0 Then
          'For Each Lwb In Wka.Workbooks
          '   MsgBox (Lwb.Name)
          '  Next Lwb
          'End If
          
        If gxla_Loaded = False Then ' Start fresh
           If DebugMode2 Then MsgBox ("Opening WKA.Workbooks " & Directory & BathBook)
           
           Wka.Workbooks.Open (Directory & BathBook)
           If DebugMode Then MsgBox ("DEBUG 08 Open Wka workbook " & BathBook & " " & Str(DebugCount))
           DebugCount = DebugCount + 1
        Else:
           'MAJOR CHANGE - CLOSE OUT ANY EXISTING BATHOUTXLS
           'This is the key to changing content of the output, we must close the
           'bathtub_output.xls before RELOADING
           If Wka.Workbooks.Count > 0 Then
             If DebugMode2 Then MsgBox "LoadExcel is closing wko"
             Wka.Workbooks(BathOutXLS).Close savechange = False
             End If
           End If
        Set Wkb = Wka.Workbooks(BathBook)
           If Wkb Is Nothing Then MsgBox ("N555: FATAL wkb is null")
           If DebugMode Then MsgBox ("DEBUG 09 set wkB to Workbooks OK" & Str(DebugCount))
           DebugCount = DebugCount + 1

        gxla_Loaded = True
       'MsgBox ("opened " & BathBook)
        'Wka.Calculation = xlCalculationManual
        ' NO NO If Wkb <> Then Wkb.Close False
       
        Set Hdr = Wkb.Sheets("headers")   'table headings sheet
        If DebugMode Then MsgBox ("DEBUG 10 set Excel Sheet Hdr OK" & Str(DebugCount))
        DebugCount = DebugCount + 1
        If Not DebugMode Then Wka.EnableEvents = True
        If DebugMode Then MsgBox ("DEBUG 10B Excel.EnableEvents Skipped" & Str(DebugCount))
        DebugCount = DebugCount + 1
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
        If DebugMode Then MsgBox ("DEBUG 11 EXCEL.WindowState xlMinimized OK" & Str(DebugCount))
        DebugCount = DebugCount + 1
        Wka.Visible = True
        
        If DebugMode Then MsgBox ("DEBUG 12 Excel Setup Done" & Str(DebugCount))
        DebugCount = DebugCount + 1
    End If
   
'bathtub_output.xls
    On Error Resume Next
    
    If DebugMode Then MsgBox ("DEBUG 13 Talking to Excel OK" & Str(DebugCount))
    DebugCount = DebugCount + 1
    i = Wko.Worksheets.Count
    If Err = 0 And i > 0 Then 'do nothing
    'MsgBox "WKO Worksheet is OK to use"
    Else
        'MsgBox "Creating Output Worksheet- this is NOT a separate workbook!"
        Err.Clear
        On Error GoTo Abend2
        'Create the new output workbook
        'Dispose of any old junk first
        If DebugMode Then MsgBox "DEBUG 13a Ready to Add Workbooks to Wka"
        If FileExists(Directory & BathOutXLS) Then Kill Directory & BathOutXLS
        Wka.Workbooks.Add
        Wka.ActiveWorkbook.SaveAs FileName:=Directory & BathOutXLS, FileFormat:=56
        If DebugMode2 Then MsgBox "N1: Loadexcel Init of bathtub_output.xls"
        Set Wko = Wka.ActiveWorkbook 'bathtub_output.xls workbook
        If DebugMode Then MsgBox "DEBUG 14 Output Workbook (WKO) Set @dbg " & Str(DebugCount)
        DebugCount = DebugCount + 1
        Wka.Application.Calculation = xlCalculationManual
    End If

    
    On Error GoTo 0
'    Status "end load"
    Exit Sub

Abend2:
    MsgBox "E102: Could not Create an Output Workbook in Excel - Bathtub Must Abort"
    Ier = 1
    Err.Clear
    Exit Sub


Abhort:
    MsgBox "E101: Failure in Sub LoadExcel (Bathtub Can't Continue)- try closing Excel then restarting BATHTUB"
    Ier = 1
    Err.Clear
End Sub

Sub SetPrintArea()
' set print area to all used cells in the current worksheet
   gLSht.UsedRange.Name = "print_area"
   End Sub
Sub ClearOutputWorkbook()
'clears output workbook without changing name
'this runs after existing case is edited
With Wko
   i = .Sheets.Count
    If i <= 0 Then Exit Sub
    Wka.DisplayAlerts = False
For j = 1 To i - 1
    .Sheets(1).Delete
    Next j
  .Sheets(1).Cells.Clear
  .Sheets(1).Name = "Sheet1"
End With
Wka.DisplayAlerts = True
End Sub

Sub ClearOutput()
'start a new output workbook
'this runs whenenever output destination is changed or new case is read
On Error Resume Next
If IsNull(Wka) Then
    If DebugMode Then MsgBox "Module1 N33: Loading Excel"
    If DebugMode2 Then MsgBox "ClearOutput is Loading Excel again"
    LoadExcel
     If Ier > 0 Then Exit Sub
     End If

Wko.Close False
If DebugMode Then MsgBox "Module1 N30: WKO Closed: " & Wko.Name
Set Wko = Nothing
On Error GoTo 0
FormUpdate
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
End Sub

Sub Backup(io)
' io=0 backup file, 1=restore backup
Dim Ofile As String

On Error GoTo Quit

    Ofile = Directory & BackupFile
If io = 0 Then
'save
        Save_btb (Ofile)
        If Ier > 0 Then GoTo Quit
    Else
'restore
        IfilE = CaseFile
        Read_btb (Ofile)
        If Ier > 0 Then GoTo Quit
        CaseFile = IfilE
        WorkingDirectory = ExtractPath(CaseFile)
    End If
    On Error GoTo 0
    Exit Sub
    
Quit:
MsgBox "Error Creating or Restoring Edit Backup File"
On Error GoTo 0
End Sub

Public Function GetFileName(flname As String) As String
    
    'Get the filename without the path or extension.
    'Input Values:
    '   flname - path and filename of file.
    'Return Value:
    '   GetFileName - name of file without the extension.
    
    Dim posn As Integer, i As Integer
    Dim Fname As String
    
    posn = 0
    'find the position of the last "\" character in filename
    For i = 1 To Len(flname)
        If (Mid(flname, i, 1) = "\") Then posn = i
    Next i

    'get filename without path
    Fname = Right(flname, Len(flname) - posn)

    'get filename without extension
    posn = InStr(Fname, ".")
        If posn <> 0 Then
            Fname = Left(Fname, posn - 1)
        End If
    GetFileName = Fname
End Function

