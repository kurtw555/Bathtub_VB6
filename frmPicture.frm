VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Bathtub Ouput"
   ClientHeight    =   12165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   11775
      Left            =   120
      ScaleHeight     =   11715
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim frmt As Variant
    For Each frmt In Array(vbCFBitmap, vbCFMetafile, vbCFDIB, vbCFPalette)
        If Clipboard.GetFormat(frmt) Then
        Set frmPicture.Picture1.Picture = Clipboard.GetData(frmt)
        
    Exit For
    End If
    Next
    w = Picture1.Width
    h = Picture1.Height
    MsgBox frmPicture.Width & " " & frmPicture.Height
    MsgBox w & " " & h
    'frmPicture.Width = w * 1.05
    'frmPicture.Height = h * 1.2
    frmPicture.Width = w + 400
    frmPicture.Height = h + 700
End Sub

