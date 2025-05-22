Attribute VB_Name = "Module4"
Function SheetExists(ShtName As String) As Boolean
'returns true if sheet exists in the active workbook
Dim X99 As Object
    On Error Resume Next
    Set X99 = ActiveWorkbook.Sheets(ShtName)
    If Err = 0 Then SheetExists = True Else SheetExists = False
    Err.Clear
End Function

Function FileExists(Fname As String) As Boolean
'returns true if the file exists
Dim X99 As String
    X99 = Dir(Fname)
    'MsgBox "fileexists: " & Fname & " " & ss
    If X99 <> "" Then
        FileExists = True
        Else
        FileExists = False
        End If
End Function

Function WorkbookIsOpen(wbname As String) As Boolean
'returns true if workbook is open
Dim X99 As Excel.Workbook
    On Error Resume Next
    Set X99 = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True Else WorkbookIsOpen = False
    Err.Clear
End Function

Function PathExists(Pname As String) As Boolean
'returns true if path exists
Dim X99 As String
On Error Resume Next
    X99 = GetAttr(Pname) And 0
    If Error = 0 Then PathExists = True Else PathExists = False
End Function

Function ExtractPath(Spec As String) As String
'extract pathname from filename
SpecLen = Len(Spec)
For i = SpecLen To 1 Step -1
    If Mid(Spec, i, 1) = "\" Then
        ExtractPath = Left(Spec, i - 1)
        Exit Function
        End If
    Next i
ExtractPath = ""
End Function

Function ExtractFile(Spec As String) As String
'extract file name from filename
SpecLen = Len(Spec)
For i = SpecLen To 1 Step -1
    If Mid(Spec, i, 1) = "\" Then
        ExtractFile = Mid(Spec, i + 1)
        Exit Function
        End If
    Next i
ExtractFile = ""
End Function


    Function FormatF(xx, fs)
'fixed length format
        i = Len(fs)
        s = Format(xx, fs)
        k = Len(s)
        FormatF = Space(i - k) & s
    End Function


'Sub SelectTest()
    'StartUp
    'ReadInputs
'SelectSegment3
'SelectCalib
'SelectSegment

'i = JSegs(Nseg + 1)
'i = Jtribs(n)
'MsgBox "Selected = " & i

'Call SelSegs(Nseg)

'End Sub
'Sub SelectOne(nN, choiceS, capT, helpiD, isel)
'select single value
'isel = 0
'If nN = 0 Then Exit Sub

'If nN = 1 Then
'    SelectOne = 1
'    Exit Function
'    End If

'adjust dimensions
'kd = 300
'ioff = MIn(nN, 20) * kd
'frmSelectOne.Height = ioff + 10 * kd
'frmSelectOne.ListBox1.Height = ioff
'frmSelectOne.CancelButton.top = ioff + kd
'frmSelectOne.OKButton.top = ioff + kd
'frmSelectOne.HelpButton.top = ioff + kd

'frmSelectOne.Caption = capT
'frmSelectOne.ListBox1.MultiSelect = 0
'Ihelp = helpiD
'MsgBox "setting up form"
'With frmSelectOne.ListBox1
'    .Clear
'    For i = 1 To nN
'        .AddItem Format(i, "00") & " " & choiceS(i)
'        Next i
'    End With
'frmSelectOne.Show vbModal

'MsgBox "returning from form"
'If Ichoice > 0 Then
'    isel = Ichoice
'    'MsgBox "Selected Item =" & Ichoice & " " & choiceS(Ichoice)
'    End If
''nd Sub

'Sub SelectMany(nN, choiceS, capT, helpiD, isel)
'
'For i = 1 To nN
'    isel(i) = 0
'    Next i
'If nN = 0 Then Exit Sub
'
'If nN = 1 Then
'    SelectOne = 1
'    Exit Function
'    End If

'adjust dimensions
'ioff = MIn(nN, 20) * 11
'frmSelectOne.Height = ioff + 60
'frmSelectOne.ListBox1.Height = ioff
'frmSelectOne.CancelButton.top = ioff + 11
'frmSelectOne.OKButton.top = ioff + 11
'frmSelectOne.HelpButton.top = ioff + 11
'frmSelectOne.Caption = capT
'frmSelectOne.ListBox1.MultiSelect = 1
'Ihelp = helpiD
'MsgBox ioff
'With frmSelectOne.ListBox1
'    .RowSource = ""
'    For i = 1 To nN
'        .AddItem Format(i, "00") & " " & choiceS(i)
'        Next i
'    End With
'frmSelectOne.Show
'End Sub

'Sub SelectSegment()
'select single segment
'If Nseg = 1 Then
'    jseG = 1
'    Exit Sub
'    End If

'adjust dimensions
'ioff = Nseg * 11
'SegmentForm.Height = ioff + 60
'SegmentForm.ListBox1.Height = ioff
'SegmentForm.CancelButton.Top = ioff + 11
'SegmentForm.OKButton.Top = ioff + 11
'SegmentForm.HelpButton.Top = ioff + 11

'With SegmentForm.ListBox1
'    .Clear
'    For i = 1 To Nseg
'        .AddItem Format(i, "00") & " " & SegName(i)
'        Next i
'    End With
'SegmentForm.Show
'If jseG > 0 Then MsgBox "Selected Segment =" & jseG & " " & SegName(jseG)
'End Sub
'Sub SelSegs(mx)
'c select segments to be used, sets iwork>0

'        Call pcharc(0, 0, Dummy, 32)
'        For i = 1 To mx
'          write(dlabel(i),'(i2.2,1x,a16)') i,SegName(i)
'          Next
'        dlabel(MX+1)='SEGMENT'
'        Call jselec(Iwork, mx, Dlabel, 24)
'        Call clr(0)
     
'        SelectMultipleSegments
    
'    Call SelectMany(mx, SegName, "Select Segments", 3, Iwork)
        
     '   For i = 1 To Nseg
      '      Iwork(i) = 1
      '      Next i
    
'    ms = "Selected Segments:"
'    For i = 1 To mx
'        If Iwork(i) > 0 Then ms = ms & vbCrLf & Format(i, "00") & " " & SegName(i)
''        Next i
 '   MsgBox ms
       
 '       End Sub

      
       'Sub SelVar(mx)
'c select observed variables

'       Iwork = 0
'       dummy='SELECT VARIABLES'
'       Call pcharc(0, 0, Dummy, 32)
'          For i = 1 To mx
'            Dlabel(i) = Cname(i)
'            Next
'          dlabel(mx+1)='VARIABLE'
'       Call jselec(Iwork, mx, Dlabel, 24)
'       Call clr(0)
       
       'SelectCalib
       'For i = 1 To mx
       'Iwork(i) = 0
       'Next i
       'Iwork(2) = 1
       'Iwork(3) = 1
                     
       
       'End Sub

          'Function JSegs(mx)
'select one segment
'          Call lcopy(mx, SegName, Iout, Dlabel, 0)
'          dlabel(mx+1)='SEGMENT------OUTFLOW SEG'
'          call pcharc(0,0,'SELECT SEGMENT, <ESC> TO QUIT',29)
'          jsegs = iselec(mx, Dlabel, 24)
          
          'Call SelectOne(mx, SegName, "Select Segment", 3, i)
          'JSegs = i
'          MsgBox "seg: " & SegName(i)
          'End Function

         'Function Jtribs(mx)
'c select one tributary
''$include:'net.inc'
'        Call lcopy(mx, TribName, Iseg, Dlabel, 0)
'        dlabel(mx+1)='TRIBUTARY--------SEGMENT'
'        call pcharc(0,0,'SELECT TRIBUTARY, <ESC> TO QUIT',31)
'        jtribs = iselec(mx, Dlabel, 24)
        'Call SelectOne(mx, TribName, "Select Tributary", 1, i)
        'Jtribs = i
        'MsgBox "trib: " & TribName(i)
        'End Function





