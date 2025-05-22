VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlot 
   Caption         =   "Bathtub Output Plot"
   ClientHeight    =   8172
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9948
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   22
   Icon            =   "frmPlot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8172
   ScaleWidth      =   9948
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLogScale 
      Caption         =   "Log Scale"
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      ToolTipText     =   "Plot observed values"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtBarWidth 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Text            =   "1"
      ToolTipText     =   "Width of error bars (number of standard errors)"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CheckBox chkObserved 
      Caption         =   "Plot Observed"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      ToolTipText     =   "Plot observed values"
      Top             =   840
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   612
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9948
      _ExtentX        =   17547
      _ExtentY        =   1080
      ButtonWidth     =   1640
      ButtonHeight    =   953
      Appearance      =   1
      HelpContextID   =   22
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Plot"
            Object.ToolTipText     =   "View plot for selected variable"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List Data"
            Object.ToolTipText     =   "List plotted data"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "Get help"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy Chart"
            Object.ToolTipText     =   "Copy chart to windows clipboard"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            Object.ToolTipText     =   "Return to program menu"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnListVar 
      Caption         =   "Variable Descriptions"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "List discriptions of model output variables that can be plotted"
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Error Bar Width:"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6255
      Left            =   120
      Top             =   1800
      Width           =   9735
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BarwidtH As Single
Dim PlotObserved As Boolean

Private Sub chkLogScale_Click()
    Combo1_click
End Sub

Private Sub chkObserved_Click()
    Combo1_click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "Plot"
    Combo1_click
    
Case "List Data"
    Combo1_click
    ViewSheet ("Plot")

Case "Help"
    ShowHelp (HelpContextID)

Case "Copy Chart"
    Clipboard.Clear
    Clipboard.SetData Image1.Picture

Case "Quit"
    Unload Me
    Status ("Ready")

End Select

End Sub
Private Sub btnlistvar_Click()
    ShowHelp (331)
End Sub
Private Sub form_initialize()
    PlotObserved = True
    txtBarWidth.Text = 1
End Sub
Private Sub Form_Load()
Combo1.Clear
For i = 1 To NDiagnostics
    Combo1.AddItem DiagName(i)
    Next i
Combo1.ListIndex = 1
End Sub

Private Sub Combo1_click()
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
    io = Combo1.ListIndex + 1
    Call ChartIt(io)
End Sub


Sub ChartIt(io)
'view chart for variable io

Dim fF As Single
Dim XBAR    As Double
Dim UpCI As Double
Dim LowCI As Double
Dim Lpic As Integer

Dim xloW As Single

If frmPlot.chkObserved.Value = 1 Then
    PlotObserved = True
    Else
    PlotObserved = False
    End If

If frmPlot.txtBarWidth.Text = "" Then
    BarwidtH = 0
    Else
    BarwidtH = Val(frmPlot.txtBarWidth.Text)
    End If

'plot_plot is where the chart is drawn, but PLOT is where the data reside
'Including the Confidence Limits (j and k aka Offset columns 9 & 10 for OBS)
Set CurrentWKChart = Wkb.Sheets("plot_plot").ChartObjects(1).Chart
    With CurrentWKChart.Axes(xlValue)
        'log scale
        If frmPlot.chkLogScale.Value = 1 Then
            .Crosses = xlAutomatic
            .MinimumScaleIsAuto = True
            .ScaleType = xlLogarithmic
            .MinorTickMark = xlOutside
        Else:
        'linear scale
            .Crosses = xlAutomatic
            .MinimumScaleIsAuto = True
            .ScaleType = xlLinear
            .MinorTickMark = xlNone
        End If
        End With

StartSheet ("Plot")

With gLSht.Range("p_label").Offset(0, -1)
    Hdr.Range("header_plot").Copy
    gLSht.Paste Destination:=.Offset(0, 0)
    .Offset(0, 1) = DiagName(io)

xloW = 10000000 'minimum value to scale log plot
For i = 1 To Nseg + 1
    .Offset(i + 3, 0) = SegName(i)
    If Cest(i, io) > 0 Then
        .Offset(i + 3, 1) = Cest(i, io)
        .Offset(i + 3, 1).NumberFormat = "0.0"
        .Offset(i + 3, 2) = Sqr(CvCest(i, io)) / Cest(i, io)
        .Offset(i + 3, 2).NumberFormat = "0.00"
        xloW = MIn(xloW, Cest(i, io))
        If BarwidtH > 0 Then
            fF = Exp(Sqr(CvCest(i, io)) / Cest(i, io) * BarwidtH)
                .Offset(i + 3, 7) = Cest(i, io) * (1 - 1 / fF)
                .Offset(i + 3, 8) = Cest(i, io) * (fF - 1)
                xloW = MIn(xloW, Cest(i, io) / fF)
                End If
        End If
        
     'IMPORTANT NOTE: Walker assumes errors are LOG-NORMALLY DISTRIBUTED AROUND MEAN
     If Cobs(i, io) > 0 Then
        .Offset(i + 3, 3) = Cobs(i, io)
        .Offset(i + 3, 3).NumberFormat = "0.0"
        .Offset(i + 3, 4) = CvCobs(i, io)
        .Offset(i + 3, 4).NumberFormat = "0.00"
        xloW = MIn(xloW, Cobs(i, io))
        If PlotObserved And BarwidtH > 0 Then
           fF = Exp(CvCobs(i, io) * BarwidtH)
                ' Lower Conf. Int Subtracted for segment i, variable io
                .Offset(i + 3, 9) = Cobs(i, io) * (1 - 1 / fF)
                ' Upper Conf. Int ADDED for segment i, variable io
                .Offset(i + 3, 10) = Cobs(i, io) * (fF - 1)
                xloW = MIn(xloW, Cobs(i, io) / fF)
                If DebugCVMode Then
                  XBAR = Cobs(i, io)
                  LowCI = XBAR - Cobs(i, io) * (1 - 1 / fF)
                  UpCI = XBAR + Cobs(i, io) * (fF - 1)
                  MsgBox "seg/Upper+ " & i & " / " & UpCI & "/" & fF
                  MsgBox "seg/Lower " & i & " / " & LowCI & "/" & fF
                  End If
                End If
        End If
        
    Next i
    End With
        
    If xloW > 0 And frmPlot.chkLogScale.Value = 1 Then
        xloW = 10 ^ Int(Log(xloW) / 2.303)
        With CurrentWKChart.Axes(xlValue)
            .MinimumScale = xloW
            .CrossesAt = xloW
            End With
        End If
    
    With gLSht
    .Range(.Range("A7"), .Range("A7").Offset(Nseg, 0)).Name = "p_segs"
    .Range(.Range("b7"), .Range("b7").Offset(Nseg, 0)).Name = "p_pred"
    If PlotObserved = True Then
        .Range(.Range("d7"), .Range("d7").Offset(Nseg, 0)).Name = "p_obs"
        Else
        .Range("A40").Name = "p_obs"
        End If
    
    Wka.Calculate
    'OLD CODE USING PICTURES
    'For Each pct In .Pictures
    '    pct.Delete
    '     Next
    
    'NEW CODE USING SHAPES COLLECTION
    Lpic = gLSht.Shapes.Count
    For i = 1 To Lpic
       'If DebugCVMode Then MsgBox "pic count:" & gLSht.Shapes.Count
       gLSht.Shapes.Item(i).Delete
       Next i
   
    If DebugMode Then MsgBox "DEBUG 16: frmPlot: About to perform CurrentWKChart.copy picture " & Str(DebugCount)
    DebugCount = DebugCount + 1
    CurrentWKChart.CopyPicture
    .Paste Destination:=.Range("A7").Offset(Nseg + 2, 0)

    Fname = Directory & "temp.gif"
    CurrentWKChart.Export FileName:=Fname, FilterName:="GIF"
    Image1.Picture = LoadPicture(Fname)
    Kill Fname
    .Range("g7:z100").ClearContents
    If Iop(12) = 2 Then .Range("j7").Offset(Nseg + 2 + 30, 0) = " "
    End With

End Sub
Private Sub txtBarWidth_change()
    If Not VerifyPositive(txtBarWidth.Text) Then txtBarWidth.Text = ""
End Sub

Private Sub txtBarWidth_LostFocus()
    Combo1_click
End Sub
