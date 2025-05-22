VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChart 
   Caption         =   "Bathtub Output Chart"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   OleObjectBlob   =   "frmChart.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub ComboBox1_Change()
    io = ComboBox1.ListIndex + 1
    Call ChartIt(io)
    LoadChart
End Sub

Sub LoadChart()

'Set CurrentCharT = Workbooks("bathtub.xls").Sheets("plotx").ChartObjects(1).Chart
    Wka.Calculate
    Set CurrentChart = Wkb.Sheets("plot").ChartObjects(1).Chart

'   Save chart as GIF
    Fname = Directory & "temp.gif"
    CurrentChart.Export FileName:=Fname, FilterName:="GIF"
   
'show the chart
    Image1.Picture = LoadPicture(Fname)

    Kill Fname
End Sub

Private Sub CommandButton1_Click()
ShowHelp (331)
End Sub

Private Sub UserForm_initialize()
ComboBox1.Clear
For i = 1 To NDiagnostics
    ComboBox1.AddItem DiagName(i)
    Next i
ComboBox1.ListIndex = 1
End Sub

'Sub SetRange(rn As String)
'    Sht.Range(rn).Cells(1, 1).Select
'    Sht.Range(Selection, Selection.Offset(0, Nseg - 1)).Name = rn
    
  '  Sht.Activate
  '  Sht.Range(rn).Range(.Offset(0, 0), .Offset(0, Nseg + 1)).Select
  '  Selection.Name = rn
 '   With Sht.Range(rn)
 '
 '   End With

'End Sub

Sub ChartIt(io)
'view chart for variable io

Set Sht = Wkb.Worksheets("plot")
'Set Rng = Sht.Range("p_label")

'.Range(Rng.Offset(0, 0), Rng.Offset(6, 50)).Value = 1
With Sht.Range("p_label")
.Range(.Cells(1, 1), .Cells(6, 50)).Value = 3

'MsgBox "hold"
.Offset(0, 0) = DiagName(io)
For i = 1 To Nseg + 1
    .Offset(1, i) = SegName(i)
    If Cest(i, io) > 0 Then
        .Offset(2, i) = Cest(i, io)
        .Offset(4, i) = Sqr(CvCest(i, io))
        Else
        .Offset(2, i).Clear
        .Offset(4, i).Clear
        End If
     If Cobs(i, io) > 0 Then
        .Offset(3, i) = Cobs(i, io)
        .Offset(5, i) = CvCobs(i, io) * Cobs(i, io)
        Else
        .Offset(3, i).Clear
        .Offset(5, i).Clear
        End If
    Next i
        
    
    .Range(.Cells(2, 2), .Cells(2, Nseg + 2)).Name = "p_segs"
    .Range(.Cells(3, 2), .Cells(3, Nseg + 2)).Name = "p_pred"
    .Range(.Cells(4, 2), .Cells(4, Nseg + 2)).Name = "p_obs"
    .Range(.Cells(5, 2), .Cells(5, Nseg + 2)).Name = "p_pred_se"
    .Range(.Cells(6, 2), .Cells(6, Nseg + 2)).Name = "p_obs_se"
       
End With
'    frmChart.Show vbModal
'    Status ("ready")

End Sub

