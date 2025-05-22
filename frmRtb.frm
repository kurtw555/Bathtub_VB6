VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRtb 
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbOutput 
      Height          =   5415
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRtb.frx":0000
   End
End
Attribute VB_Name = "frmRtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim frmt As Variant
    'For Each frmt In Array
        'Clipboard.GetFormat(frmt)
        rtbOutput.SelRTF = Clipboard.GetText(vbCFRTF)
        
   ' Exit For
   ' End If
    'Next
End Sub

