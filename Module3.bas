Attribute VB_Name = "Module3"
Sub Output_Init()
'initialize output list parameters
  
    OutputDest = 1
    ShowWarnings = True
    
    NGlobals = 4
    GlobalName(1) = "AVERAGING PERIOD (YRS)"
    GlobalName(2) = "PRECIPITATION (METERS)"
    GlobalName(3) = "EVAPORATION (METERS)"
    GlobalName(4) = "INCREASE IN STORAGE (METERS)"
    
'input concentrations
    NVariables = 5
    VariableName(1) = "CONSERVATIVE SUBST."
    VariableName(2) = "TOTAL P"
    VariableName(3) = "TOTAL N"
    VariableName(4) = "ORTHO P"
    VariableName(5) = "INORGANIC N"
    
'map components
    Imap(1) = 1
    Imap(2) = 6
    Imap(3) = 7

'mass balance type labels
    MassBalName(0) = "Observed"
    MassBalName(1) = "Predicted"
      
'type codes
    N_Type_Codes = 5
    Type_Code(1) = "Monitored Inflow"
    Type_Code(2) = "Non Point Inflow"
    Type_Code(3) = "Point Source"
    Type_Code(4) = "Reservoir Outflow"
    Type_Code(5) = "Downstream Boundary"
    'Type_Code(6) = "Internal Load"
    
 
 'term names
     NtermS = 18
     TermName(1) = "PRECIPITATION"
     TermName(2) = "EXTERNAL INFLOW"
     TermName(3) = "***EVAPORATION"
     TermName(4) = "GAUGED OUTFLOW"
     TermName(5) = "***RETENTION"
     TermName(6) = "ADVECTIVE INFLOW"
     TermName(7) = "ADVECTIVE OUTFLOW"
     TermName(8) = "DIFFUSIVE INFLOW"
     TermName(9) = "***TOTAL INFLOW"
     TermName(10) = "***TOTAL OUTFLOW"
     TermName(11) = "***STORAGE INCREASE"
     TermName(12) = "DIFFUSIVE OUTFLOW"
     TermName(13) = "NET DIFFUSIVE INFLOW"
     TermName(14) = "NET DIFFUSIVE OUTFLOW"
     TermName(15) = "INTERNAL LOAD"
     TermName(16) = "TRIBUTARY INFLOW"
     TermName(17) = "NONPOINT INFLOW"
     TermName(18) = "POINT-SOURCE INFLOW"

'C SEGMENT MASS BALANCE TERMS TO BE PRINTED FORMAT 1

    Njord = 15
    Jord(1) = 1
    Jord(2) = 15
    Jord(3) = 16
    Jord(4) = 17
    Jord(5) = 18
    Jord(6) = 6
    Jord(7) = 13
    Jord(8) = 9
    Jord(9) = 4
    Jord(10) = 7
    Jord(11) = 14
    Jord(12) = 10
    Jord(13) = 3
    Jord(14) = 11
    Jord(15) = 5
 '      data njord/15/
 '      data jord/1,15,16,17,18,6,13,9,4,7,14,10,3,11,5/

'C GROSS MASS BALANCE TERMS TO BE PRINTED

    Nkord = 14
      Kord(1) = 1
      Kord(2) = 15
      Kord(3) = 16
      Kord(4) = 17
      Kord(5) = 18
      Kord(6) = 13
      Kord(7) = 9
      Kord(8) = 4
      Kord(9) = 7
      Kord(10) = 14
      Kord(11) = 10
      Kord(12) = 3
      Kord(13) = 11
      Kord(14) = 5
            
  '   DATA NKORD/14/,KORD/1,15,16,17,18,13,9,4,7,14,10,3,11,5,1*0/

'mass balance terms format2
'  DATA MORD/8/,LORD/2,1,6,11,7,4,8,3,0,0/

    Mord = 8
    Lord(1) = 2
    Lord(2) = 1
    Lord(3) = 6
    Lord(4) = 11
    Lord(5) = 7
    Lord(6) = 4
    Lord(7) = 8
    Lord(8) = 3

    End Sub
   
Sub List_All()
   'MsgBox "minimizing Excel Template"
        realversion = CDbl(Wka.Version)
        If (realversion < 15) Then Wka.WindowState = xlMinimized
   j = 0
   For i = 0 To 9
   Select Case i
   Case 9 'reversed the order to accomodate excel .ADD behavior
    List_Inputs
   Case 8
    List_Tree
   Case 7
    List_Hydraulics
   Case 6
    List_GrossBalances
   Case 5
    If Not NoviceUser Then List_SegBalances
   Case 4
    If Not NoviceUser Then List_Terms
   Case 3
    List_Diagnostics (1)
   Case 2
    If Not NoviceUser Then List_Profiles
   Case 1
    If Not NoviceUser Then List_TTests
   Case 0
    If Not NoviceUser Then List_Fits
    
   End Select
   
   sN = gLSht.Name
   Status ("Listing: " & sN)
   'MsgBox ("About to Copy " & sN)
   SheetCopy (sN)
   Next i
   Status "Ready"
   Wka.WindowState = xlNormal
  
  End Sub 'List_All

Sub StartSheet(sN) 'sN is Sheet Name
'start new output sheet
    If DebugMode Then MsgBox ("Loading Excel from Module3, StartSheet " & sN)
    'LoadExcel
    'MsgBox "N22: StartSheet (module3) is Calling ClearOutput"
    'ClearOutput
    Status (sN)
    Wkb.Sheets(sN).Cells.Clear
    'MsgBox "Cleared Cells in " & Wkb.Name & " " & sN
    Set gLSht = Wkb.Sheets(sN)
    With gLSht
     If DebugMode Then MsgBox "StartSheet Clearing: " & Wkb.Name & "." & gLSht.Name & "  " & Title
     '.Cells.Clear
     .Range("A1") = Title
     .Range("a1", "b2").Font.Bold = True
     .Range("A2") = "File:"
     .Range("b2") = CaseFile
     End With
End Sub
Sub List_Hydraulics()

Dim atoT As Single, vtoT As Single

StartSheet ("Hydraulics")
line_no = 0
With gLSht.Range("A4")
        Hdr.Range("Header_Hydrau").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
        line_no = line_no + 3
        For i = 1 To Nseg
        line_no = line_no + 1
           Call Hydrau(i)
           .Offset(line_no, 0) = i
           .Offset(line_no, 0).HorizontalAlignment = xlCenter
           .Offset(line_no, 1) = SegName(i)
           .Offset(line_no, 2) = Iout(i)
           .Offset(line_no, 2).HorizontalAlignment = xlCenter
           .Offset(line_no, 3) = Qnet(i)
           .Offset(line_no, 3).NumberFormat = "0.0"
           For k = 3 To 8
               .Offset(line_no, 1 + k) = x(k)
               If k = 3 Then
                    .Offset(line_no, 1 + k).NumberFormat = "0.0000"
                    Else
                    .Offset(line_no, 1 + k).NumberFormat = "0.0"
                    End If
               Next k
            Next i
            

     line_no = line_no + 2
     atoT = 0
     vtoT = 0
        Hdr.Range("Header_Morpho").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
     line_no = line_no + 2
         For i = 1 To Nseg
            line_no = line_no + 1
            .Offset(line_no, 0) = i
            .Offset(line_no, 0).HorizontalAlignment = xlCenter
            .Offset(line_no, 1) = SegName(i)
             x(1) = Area(i)
             x(2) = Zmn(i)
             x(3) = Zmx(i)
             x(4) = Slen(i)
             x(5) = Area(i) * Zmn(i)
             atoT = atoT + Area(i)
             vtoT = vtoT + x(5)
             x(6) = Ratv(Area(i), Slen(i))
             x(7) = Ratv(Slen(i), x(6))
             For j = 1 To 7
                .Offset(line_no, j + 1) = x(j)
                .Offset(line_no, j + 1).NumberFormat = "0.0"
                Next j
             Next i
             line_no = line_no + 1
             .Offset(line_no, 0) = "Totals"
             .Offset(line_no, 2).NumberFormat = "0.0"
             .Offset(line_no, 3).NumberFormat = "0.0"
             .Offset(line_no, 6).NumberFormat = "0.0"
             .Offset(line_no, 2) = atoT
             .Offset(line_no, 6) = vtoT
             .Offset(line_no, 3) = Ratv(vtoT, atoT)

End With
End Sub

Sub List_GrossBalances()

Dim Da(20) As Single
Dim runofF As Single
Dim vaR As Single
Dim xL As Single
Dim cV As Single
Dim cU As Single
Dim pV As Single
Dim pV1 As Single
Dim pL As Single
Dim cvQ As Single
Dim qU As Single
Dim tW As Single

'C kPRINT=3 GROSS WATER AND NUTRIENT BALANCES

'c water balances
        If NTrib <= 0 Then Exit Sub
        'Call Balan
        Call MassBalanceTerms
        
 StartSheet ("Overall Balances")
    line_no = 0
    With gLSht.Range("A4")
        Hdr.Range("Header_gwbal").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
        .Offset(line_no + 2, 7) = P(1)
    line_no = line_no + 4
           
        For i = 1 To 18
            Da(k) = 0
            Next i
    
        For i = 1 To NTrib
           If Icode(i) <= 4 Then
           If Icode(i) = 4 Then
                k = 4
                Else
                k = 15 + Icode(i)
                End If
           Da(k) = Da(k) + Darea(i)
           qU = Flow(i)
           cvQ = CvFlow(i)
           If Icode(i) = 2 Then
                qU = FlowL(i)
                cvQ = CvFlowL(i)
                End If
           runofF = Ratv(qU, Darea(i))
           vaR = (cvQ * qU) ^ 2
         If qU > 0 Then
        line_no = line_no + 1
        .Offset(line_no, 0) = i
        .Offset(line_no, 0).HorizontalAlignment = xlCenter
        .Offset(line_no, 1) = Icode(i)
        .Offset(line_no, 1).HorizontalAlignment = xlCenter
        .Offset(line_no, 2) = Iseg(i)
        .Offset(line_no, 2).HorizontalAlignment = xlCenter
        .Offset(line_no, 3) = TribName(i)
        If Darea(i) > 0 Then .Offset(line_no, 4) = Darea(i)
        .Offset(line_no, 4).NumberFormat = "0.0"
        .Offset(line_no, 5) = qU
        .Offset(line_no, 5).NumberFormat = "0.0"
        .Offset(line_no, 6) = vaR
        .Offset(line_no, 6).NumberFormat = "0.00E+00"
        If qU > 0 Then .Offset(line_no, 7) = cvQ
        .Offset(line_no, 7).NumberFormat = "0.00"
        If runofF > 0 Then .Offset(line_no, 8) = runofF
        .Offset(line_no, 8).NumberFormat = "0.00"
        End If
        
            End If
        Next i

'C DRAINAGE AREAS
'c total external
        Da(2) = Da(16) + Da(17) + Da(18)
'c precip
        Da(1) = Area(Nseg + 1)
'C TOTAL IN
        Da(9) = Da(1) + Da(2)
'C put da error in advective outflow
        Da(7) = Da(9) - Da(4)
        Da(10) = Da(9)
        Da(5) = 0

'c Summary
        For j = 1 To Nkord
           i = Kord(j)
           If Not (i = 5 Or i = 15 Or Term(1, i) = 0) Then
           line_no = line_no + 1
           runofF = Ratv(Term(1, i), Da(i))
           cV = MIn(Ratv(Sqr(CvTerm(1, i)), Abs(Term(1, i))), 9.99)
           .Offset(line_no, 0) = TermName(i)
           If Da(i) > 0 Then .Offset(line_no, 4) = Da(i)
           .Offset(line_no, 4).NumberFormat = "0.0"
           .Offset(line_no, 5) = Term(1, i)
           .Offset(line_no, 5).NumberFormat = "0.0"
           .Offset(line_no, 6) = CvTerm(1, i)
           .Offset(line_no, 6).NumberFormat = "0.00E+00"
           .Offset(line_no, 7) = cV
           .Offset(line_no, 7).NumberFormat = "0.00"
           If runofF > 0 Then .Offset(line_no, 8) = runofF
           .Offset(line_no, 8).NumberFormat = "0.00"
           End If
           
            Next j
       
'c gross nutrient balances
      
        For iC = 1 To 3
            If Iop(iC) > 0 Then
        
        line_no = line_no + 3
        Hdr.Range("Header_Gross").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
                       
        .Offset(line_no, 4) = MassBalName(Iop(11))
        .Offset(line_no + 1, 4) = VariableName(Imap(iC))
        line_no = line_no + 3
        tW = Term(iC + 1, 9)         'total load
        Tf = Term(1, 9)              'total flow
        
        For i = 1 To NTrib
            If Icode(i) < 5 And Iseg(i) > 0 Then
        '    If (Icode(i) >= 5 Or Iseg(i) <= 0) Then Next i
            jseG = Iseg(i)
            If Icode(i) = 2 Then
                cU = Concil(i, Imap(iC))
                cV = CvCil(i, Imap(iC))
                qU = FlowL(i)
                cvQ = CvFlowL(i)
                Else
                cU = Conci(i, Imap(iC))
                cV = CvCi(i, Imap(iC))
                qU = Flow(i)
                cvQ = CvFlow(i)
                End If
            'If Icode(i) = 4 And Iop(11) = 1 Then
            '    cU = Cest(jseG, iC)
                If Icode(i) = 4 Then
                    cU = Cuse(Cobs(jseG, iC), Cest(jseG, iC), Iop(11))
                    cV = Ratv(Sqr(CvCest(jseG, iC)), cU)
                    End If
           xL = qU * cU
           vaR = (cV ^ 2 + cvQ ^ 2) * xL ^ 2
           Export = Ratv(xL, Darea(i))
           cV = Ratv(Sqr(vaR), xL)
           pL = Ratv(xL, tW)
           pV1 = Ratv(vaR, CvTerm(iC + 1, 9))
           If xL > 0 Then
           line_no = line_no + 1
            .Offset(line_no, 0) = i
            .Offset(line_no, 0).HorizontalAlignment = xlCenter
            .Offset(line_no, 1) = Icode(i)
            .Offset(line_no, 1).HorizontalAlignment = xlCenter
            .Offset(line_no, 2) = Iseg(i)
            .Offset(line_no, 2).HorizontalAlignment = xlCenter
            .Offset(line_no, 3) = TribName(i)
            .Offset(line_no, 4) = xL
            .Offset(line_no, 4).NumberFormat = "0.0"
                If pL > 0 And Icode(i) <> 4 Then
                    .Offset(line_no, 5) = pL
                    .Offset(line_no, 5).NumberFormat = "0.0%"
                    End If
                .Offset(line_no, 6) = vaR
                .Offset(line_no, 6).NumberFormat = "0.00E+00"
                If pV1 > 0 And Icode(i) <> 4 Then
                    .Offset(line_no, 7) = pV1
                    .Offset(line_no, 7).NumberFormat = "0.0%"
                    End If
                .Offset(line_no, 8) = cV
                .Offset(line_no, 8).NumberFormat = "0.00"
            If cU > 0 Then .Offset(line_no, 9) = cU
            .Offset(line_no, 9).NumberFormat = "0.0"
            If Export > 0 Then .Offset(line_no, 10) = Export
            .Offset(line_no, 10).NumberFormat = "0.0"
            End If
            End If
            Next i
'summary terms
         For j = 1 To Nkord
          i = Kord(j)
          If i <> 3 And Term(iC + 1, i) <> 0 Then
          cU = 0
          Export = 0
          cV = 0
          If Not (i = 5 Or i = 15) Then
            cU = Ratv(Term(iC + 1, i), Term(1, i))
            Export = Ratv(Term(iC + 1, i), Da(i))
            End If
         cV = Ratv(Sqr(CvTerm(iC + 1, i)), Abs(Term(iC + 1, i)))
         cV = MIn(cV, 9.999)
         pL = Ratv(Term(iC + 1, i), tW)
         If i <= 2 Or i = 9 Or i = 13 Or i > 14 Then
            pV1 = Ratv(CvTerm(iC + 1, i), CvTerm(iC + 1, 9))
            Else
            pV1 = 0
            End If
         line_no = line_no + 1
        .Offset(line_no, 0) = TermName(i)
        .Offset(line_no, 4) = Term(iC + 1, i)
         .Offset(line_no, 4).NumberFormat = "0.0"
        If pL > 0 Then .Offset(line_no, 5) = pL
         .Offset(line_no, 5).NumberFormat = "0.0%"
        .Offset(line_no, 6) = CvTerm(iC + 1, i)
         .Offset(line_no, 6).NumberFormat = "0.00E+00"
        If pV1 > 0 Then .Offset(line_no, 7) = pV1
         .Offset(line_no, 7).NumberFormat = "0.0%"
        .Offset(line_no, 8) = cV
         .Offset(line_no, 8).NumberFormat = "0.00"
        If cU > 0 Then .Offset(line_no, 9) = cU
         .Offset(line_no, 9).NumberFormat = "0.0"
        If Export > 0 Then .Offset(line_no, 10) = Export
          .Offset(line_no, 10).NumberFormat = "0.0"

        End If
        Next j

 'C MASS BALANCE STATISTICS
        For i = 1 To 8
            x(i) = 0
            Next i

'C OVERFLOW RATE M/YR  BASED ON NET INFLOW
        If Area(Nseg + 1) > 0 Then x(1) = MAx(((Term(1, 9) - Term(1, 3)) / Area(Nseg + 1)), 0)

'C RESIDENCE TIME (YRS)
        If x(1) > 0 Then x(2) = Zmn(Nseg + 1) / x(1)

'C MEAN POOL CONC
'        x(3) = Cobs(Nseg + 1, iC)
        'If iO > 1 Then x(3) = Cest(Nseg + 1, ic)  ' ?????
'        If Iop(11) = 1 Or x(3) <= 0 Then x(3) = Cest(Nseg + 1, iC) ' fixed 4/1/2004
        x(3) = Cuse(Cobs(Nseg + 1, iC), Cest(Nseg + 1, iC), Iop(11))
        
'C NUTRIENT RESIDENCE TIME
        x(4) = Ratv(Area(Nseg + 1) * Zmn(Nseg + 1) * x(3), Term(iC + 1, 9))

'C TURNOVER RATIO
        x(5) = Ratv(P(1), x(4))

'C RETENTION COEF
        x(6) = Ratv(Term(iC + 1, 5), Term(iC + 1, 9))
        
        
        line_no = line_no + 2
        .Offset(line_no, 1) = "Overflow Rate (m/yr)"
        .Offset(line_no, 4).NumberFormat = "0.0"
        .Offset(line_no, 4) = x(1)
        
        line_no = line_no + 1
        .Offset(line_no, 1) = "Hydraulic Resid. Time (yrs)"
        .Offset(line_no, 4) = x(2)
        .Offset(line_no, 4).NumberFormat = "0.0000"
        
        line_no = line_no + 1
        .Offset(line_no, 1) = "Reservoir Conc (mg/m3)"
        .Offset(line_no, 4) = x(3)
        .Offset(line_no, 4).NumberFormat = "0"
        
        line_no = line_no - 2
        .Offset(line_no, 6) = "Nutrient Resid. Time (yrs)"
        .Offset(line_no, 9) = x(4)
        .Offset(line_no, 9).NumberFormat = "0.0000"
        
        line_no = line_no + 1
        .Offset(line_no, 6) = "Turnover Ratio"
        .Offset(line_no, 9) = x(5)
        .Offset(line_no, 9).NumberFormat = "0.0"
        
        line_no = line_no + 1
        .Offset(line_no, 6) = "Retention Coef."
        .Offset(line_no, 9) = x(6)
        .Offset(line_no, 9).NumberFormat = "0.000"

        
        End If
        Next iC
End With

End Sub
  
  Sub Dumpa(io)
  
  'save mass balance solution matrix - debugging
   
  With Wkb.Sheets("matrix").Range("A1")
   If io = 0 Then
        ioff = 0
        .Cells.Clear
        '.Range("A1:z999").ClearContents
        Else
        ioff = io * 12
        End If
    
  For i = 1 To Nseg
  For j = 1 To Nseg + 1
    If io < 2 Then
    .Offset(i + ioff, j - 1) = A(i, j)
    Else
    .Offset(i + ioff, j - 1) = Q(i, j)
    End If
  Next j
  Next i
  
  End With
End Sub
  
Sub List_Terms()

'C IPRINT=5 SEGMENT WATER AND MASS BALANCES - FORMAT 2
    
    Dim k As Integer
    Dim jc As Integer
    Dim iC As Integer
    Dim i As Integer

'terms
       
       If NTrib <= 0 Then Exit Sub
        
        StartSheet ("Summary Balances")
        
        With gLSht.Range("A3")
        line_no = 0
        For jc = 1 To 4
            iC = jc - 1
        
        If Not (iC > 0 And Iop(iC) <= 0) Then
        
        For i = 1 To 10
            y(i) = 0
            Next i
        
        If iC = 0 Then
        
        Lord(8) = 3         'water balance - uses evaporation in last slot
        line_no = line_no + 1
        Hdr.Range("Header_Wbal").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
        .Offset(line_no, 5) = P(1) 'Averaging Period
        
        Else

        Lord(8) = 5          'mass balance - uses net retention in last slot
        line_no = line_no + 2
        Hdr.Range("Header_Mbal").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
        .Offset(line_no, 8) = VariableName(Imap(iC))
        .Offset(line_no, 3) = MassBalName(Iop(11))
        End If
        line_no = line_no + 2
        
        For jseG = 1 To Nseg
          k = MAx(iC, 1)
          Call Balas(jseG, k, Iop(11))
          For i = 1 To Mord
             j = Lord(i)
             If iC = 0 Then
                        x(i) = Qt(j)
                Else
                        x(i) = Bt(j)
                End If
                Next i
        
        
        If iC > 0 Then
            x(8) = Bt(5) - Bt(15)   'retention - internal load
            'x(7) = Bt(12) - Bt(8)   'net diffusive output
            x(7) = Bt(14) - Bt(13)
            Else
            x(7) = Exch(jseG)
            End If
        
        If Iout(jseG) = 0 Then y(5) = y(5) + x(5)
        For i = 1 To Mord
                If (i <> 5) Then y(i) = y(i) + x(i)
                Next i
        line_no = line_no + 1
            .Offset(line_no, 0) = jseG
            .Offset(line_no, 0).HorizontalAlignment = xlCenter
            .Offset(line_no, 1) = SegName(jseG)
        For i = 1 To Mord
            .Offset(line_no, i + 1) = x(i)
            .Offset(line_no, i + 1).NumberFormat = "0"
            Next i
        
        Next jseG
        
        y(3) = 0
        y(7) = 0
        line_no = line_no + 1
        .Offset(line_no, 0) = "Net"
        For i = 1 To Mord
           .Offset(line_no, i + 1) = y(i)
           .Offset(line_no, i + 1).NumberFormat = "0"
            Next i
        
        End If
        Next jc
      
        End With
End Sub

Sub List_SegBalances()

'C IPRINT=4  WATER AND NUTRIENT BALANCES BY SEGMENT
    Dim iC As Integer
    Dim i As Integer
    Dim cU As Single
    Dim j As Integer
    
        If NTrib <= 0 Then Exit Sub
        StartSheet ("Segment Balances")
        
        With gLSht.Range("A4")
                
        line_no = 0
        .Offset(line_no, 0) = "Segment Mass Balance Based Upon " & MassBalName(Iop(11)) & " Concentrations"
        .Offset(line_no, 0).Font.Bold = True
    
        For jseG = 1 To Nseg
        If Izap(jseG) = 0 Then

        For iC = 1 To 3
        If Iop(iC) > 0 Then
        Call Balas(jseG, iC, Iop(11))
        
        line_no = line_no + 2
        Hdr.Range("Header_segbal").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
            .Offset(line_no, 2) = VariableName(Imap(iC))
            .Offset(line_no, 5) = jseG
            .Offset(line_no, 5).HorizontalAlignment = xlCenter
            .Offset(line_no, 6) = SegName(jseG)
            line_no = line_no + 2

' 393    WRITE(NOUT,301) VariableName(IC),JSEG,SegName(JSEG)
' 301    FORMAT(' COMPONENT: ',A8,20x,
'     * '  SEGMENT:',I3,1X,a16/24x,
'     * '     --- FLOW ---       --- LOAD ---         CONC'/
'     * ' ID  T LOCATION         ',
'     * '     HM3/YR      %       KG/YR      %       MG/M3')

'C EXTERNAL LOAD COMPONENTS

        Tf = Qt(9)
        tW = Bt(9)
        For i = 1 To NTrib
           If Not (Iseg(i) <> jseG Or Icode(i) >= 5) Then
           If Icode(i) = 2 Then
                cU = Concil(i, Imap(iC))
                qU = FlowL(i)
                Else
                cU = Conci(i, Imap(iC))
                qU = Flow(i)
                End If
'           If Icode(i) > 3 And Iop(11) = 1 Then cU = Cest(jseG, iC)
          If Icode(i) > 3 Then cU = Cuse(Cobs(jseG, iC), Cest(jseG, iC), Iop(11))  '4/11/2004
           
           xL = qU * cU
            If xL > 0 Then
            line_no = line_no + 1
            .Offset(line_no, 0) = i
            .Offset(line_no, 0).HorizontalAlignment = xlCenter
            .Offset(line_no, 1) = Icode(i)
            .Offset(line_no, 1).HorizontalAlignment = xlCenter
            .Offset(line_no, 2) = TribName(i)
            .Offset(line_no, 3) = qU
            .Offset(line_no, 3).NumberFormat = "0.0"
            .Offset(line_no, 4) = Ratv(qU, Tf)
            .Offset(line_no, 4).NumberFormat = "0.0%"
            .Offset(line_no, 5) = xL
            .Offset(line_no, 5).NumberFormat = "0.0"
            .Offset(line_no, 6) = Ratv(xL, tW)
            .Offset(line_no, 6).NumberFormat = "0.0%"
            If cU > 0 Then .Offset(line_no, 7) = cU
            .Offset(line_no, 7).NumberFormat = "0"
            End If
        End If
        Next i

'C SUMMARY TERMS

        For j = 1 To Njord
           i = Jord(j)
           If Not (j < Njord And Bt(i) = 0# And Qt(i) = 0#) Then
           cU = 0
           If (i <> 5 And i <> 15) Then cU = Ratv(Bt(i), Qt(i))
            line_no = line_no + 1
            .Offset(line_no, 0) = TermName(i)
            .Offset(line_no, 3) = Qt(i)
            .Offset(line_no, 3).NumberFormat = "0.0"
            .Offset(line_no, 4) = Ratv(Qt(i), Tf)
            .Offset(line_no, 4).NumberFormat = "0.0%"
            .Offset(line_no, 5) = Bt(i)
            .Offset(line_no, 5).NumberFormat = "0.0"
            .Offset(line_no, 6) = Ratv(Bt(i), tW)
            .Offset(line_no, 6).NumberFormat = "0.0%"
            If cU > 0 Then
                    .Offset(line_no, 7) = cU
                    .Offset(line_no, 7).NumberFormat = "0"
                    End If
            End If
            Next j

        Call Hydrau(jseG)
        
'        WRITE(NOUT,333) x'(3),x(4),zmn(jseg)
 '333    FORMAT(
  '   * ' RESID. TIME =',F8.3,' YRS, OVERFLOW RATE =',F8.1,' M/YR',
   '  * ', DEPTH =',F6.1,' M')

        line_no = line_no + 2
        .Offset(line_no, 0) = "Hyd. Residence Time ="
        .Offset(line_no, 3) = x(3)
        .Offset(line_no, 3).NumberFormat = "0.0000"
        .Offset(line_no, 4) = " yrs"
        
        line_no = line_no + 1
        .Offset(line_no, 0) = "Overflow Rate ="
        .Offset(line_no, 3) = x(4)
        .Offset(line_no, 3).NumberFormat = "0.0"
        .Offset(line_no, 4) = " m/yr"
        
        line_no = line_no + 1
        .Offset(line_no, 0) = "Mean Depth ="
        .Offset(line_no, 3) = Zmn(jseG)
        .Offset(line_no, 3).NumberFormat = "0.0"
        .Offset(line_no, 4) = " m"
        

        End If
        Next iC
        End If
        Next jseG
        End With
End Sub

Sub List_TTests()

'C IPRINT=6 PRINT OBSERVED AND PREDICTED WQ

        line_no = 0
        StartSheet ("T tests")
        
        With gLSht.Range("A4")
        
        Hdr.Range("Header_ttest1").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
        line_no = line_no + 3
                       
        For ii = 1 To Nseg + 1
        If ii = 1 Then
            i = Nseg + 1
            Else
            i = ii - 1
            End If
        
        If Izap(i) = 0 And Not (Nseg = 1 And ii = 1) Then
        
        line_no = line_no + 2
        Hdr.Range("Header_ttest2").Copy
        gLSht.Paste Destination:=.Offset(line_no, 0)
        If i <= Nseg Then .Offset(line_no, 1) = i
        .Offset(line_no, 1).HorizontalAlignment = xlCenter
        .Offset(line_no, 2) = SegName(i)
        line_no = line_no + 2

        nord1 = 13
        For l = 1 To nord1
          j = Iord(l)
        If Cobs(i, j) > 0 And Cest(i, j) > 0 Then         'list only paired results
          For k = 1 To 4
                x(k) = 0
                Next k
          cV = Ratv(Sqr(CvCest(i, j)), Cest(i, j))
          If Cest(i, j) > 0 And Cobs(i, j) > 0 Then
            x(1) = Cobs(i, j) / Cest(i, j)
            y(1) = CvCobs(i, j)
            y(2) = Stat(j, 3)
            y(3) = Sqr(cV ^ 2 + CvCobs(i, j) ^ 2)
            y(4) = Log(x(1))
          For k = 1 To 3
             If y(k) > 0 Then x(k + 1) = y(4) / y(k)
             Next k
          If Iop(9) <= 0 Then x(4) = 0
          End If
           line_no = line_no + 1
          .Offset(line_no, 0) = DiagName(j)
          If Cobs(i, j) > 0 Then
            .Offset(line_no, 1) = Cobs(i, j)
            .Offset(line_no, 1).NumberFormat = "0.0"
            .Offset(line_no, 2) = CvCobs(i, j)
            .Offset(line_no, 2).NumberFormat = "0.00"
             End If
          .Offset(line_no, 3) = Cest(i, j)
          .Offset(line_no, 3).NumberFormat = "0.0"
          .Offset(line_no, 4) = cV
          .Offset(line_no, 4).NumberFormat = "0.00"
          For k = 1 To 4
            If x(1) > 0 And x(k) <> 0 Then
                    .Offset(line_no, 4 + k) = x(k)
                    .Offset(line_no, 4 + k).NumberFormat = "0.00"
                    End If
            Next k
            End If
            Next l
 
        End If
        Next ii
        End With
End Sub

Sub List_Diagnostics(io As Integer)

'c iprint=7 diagnostics  io=0 predicted, io=1 predicted & observed

line_no = 0
StartSheet ("Diagnostics")

If io = 0 Then
    Status ("Predicted")
    Else
    Status ("Predicted & Observed")
    End If

With gLSht.Range("A4")
    If io = 1 Then
        .Offset(0, 0) = "Predicted & Observed Values Ranked Against CE Model Development Dataset"
        Else
        .Offset(0, 0) = "Predicted Values Ranked Against CE Model Development Dataset"
        End If
        .Offset(0, 0).Font.Bold = True
        
        If Nseg = 1 Then
            kseG = 1
            Else
            kseG = Nseg + 1
            End If
        
      For ii = 1 To kseG
        If ii = 1 Then
            i = kseG
            Else
            i = ii - 1
            End If

      If Izap(i) = 0 Then
        
        line_no = line_no + 2
        If io = 0 Then
            Hdr.Range("header_diag_0").Copy
            Else
            Hdr.Range("Header_diag").Copy
            End If
        gLSht.Paste Destination:=.Offset(line_no, 0)
        
        .Offset(line_no, 1) = i
        .Offset(line_no, 1).HorizontalAlignment = xlCenter
        .Offset(line_no, 2) = SegName(i)
        line_no = line_no + 2
        
        For l = 1 To Nord
          j = Iord(l)
          If Not (Cobs(i, j) <= 0 And Cest(i, j) <= 0) Then
   
   'ranking
          If j > 19 And j < 26 Then     'chla freq
                k = 4
            ElseIf j = 26 Then     'tsip
                k = 2
            ElseIf j = 27 Then     'tsi chla
                k = 4
            ElseIf j = 28 Then     'tsi secchi
                k = 5
            Else
                k = j
            End If
          p1 = 0
          If Cobs(i, k) > 0 Then Call Rank(Cobs(i, k), Stat(k, 1), Stat(k, 2), t, p1)
          p2 = 0
          If Cest(i, k) > 0 Then Call Rank(Cest(i, k), Stat(k, 1), Stat(k, 2), t, p2)
          If j = 28 Then                 'reverse ranking for
            If p1 > 0 Then p1 = 1 - p1
            If p2 > 0 Then p2 = 1 - p2
            End If
        
        line_no = line_no + 1
          .Offset(line_no, 0) = DiagName(j)
        
        If Cest(i, j) > 0 Then
                .Offset(line_no, 1) = Cest(i, j)
                If CvCest(i, j) > 0 Then .Offset(line_no, 2) = Sqr(CvCest(i, j)) / Cest(i, j)
                If p2 > 0 Then .Offset(line_no, 3) = p2
                .Offset(line_no, 3).NumberFormat = "0.0%"
                .Offset(line_no, 1).NumberFormat = "0.0"
                .Offset(line_no, 2).NumberFormat = "0.00"
                End If
                        
        If Cobs(i, j) > 0 And io > 0 Then
                .Offset(line_no, 4) = Cobs(i, j)
                If CvCobs(i, j) > 0 Then .Offset(line_no, 5) = CvCobs(i, j)
                If p1 > 0 Then .Offset(line_no, 6) = p1
                .Offset(line_no, 6).NumberFormat = "0.0%"
                .Offset(line_no, 4).NumberFormat = "0.0"
                .Offset(line_no, 5).NumberFormat = "0.00"
                End If
       
        

            End If
        
        Next l
        End If

        Next ii

End With
End Sub

Sub Rank(x1, z1, z2, t, perc)
    perc = 0.5
    t = 0
    If x1 = z1 Then Exit Sub
    perc = 0
    If z2 * z1 <= 0 Or x1 <= 0 Then Exit Sub
    t = Log(x1 / z1) / z2
    perc = Application.WorksheetFunction.NormSDist(t)
End Sub

Sub List_Profiles()

'C IPRINT=8 PRINT PROFILES
Dim kseG As Integer
Dim sN As String
Dim kU As Integer
Dim io As Integer

line_no = 0
StartSheet ("Profiles")
With gLSht.Range("A4")

If Nseg = 1 Then
    kseG = 1
    Else
    kseG = Nseg + 1
    End If

.Offset(line_no, 0) = "Segment"
.Offset(line_no, 1) = " Name"
.Offset(line_no, 1).Font.Bold = True
.Offset(line_no, 0).Font.Bold = True
.Offset(line_no, 0).HorizontalAlignment = xlCenter
For j = 1 To kseG
    line_no = line_no + 1
    If j = Nseg + 1 Then
        .Offset(line_no, 0) = "Mean"
        .Offset(line_no, 0).HorizontalAlignment = xlRight
        Else
        .Offset(line_no, 0) = j
        End If
        .Offset(line_no, 1) = " " & SegName(j)
    Next j

line_no = line_no + 1

For io = 1 To 5
        Select Case io
        Case 1
           sN = "PREDICTED CONCENTRATIONS:"
        Case 2
           sN = "OBSERVED CONCENTRATIONS:"
         Case 3
           sN = "OBSERVED/PREDICTED RATIOS:"
         Case 4
           sN = "OBSERVED STANDARD ERRORS"
         Case 5
           sN = "PREDICTED STANDARD ERRORS"
          End Select
           
          line_no = line_no + 2
            .Offset(line_no, 0) = sN
            .Offset(line_no, 0).Font.Bold = True

line_no = line_no + 1

'With .Range(.Offset(0, 0), .Offset(0, kseG + 1)).Font
'    .Bold = True
'    .Underline = True
'End With

For i = 1 To kseG
    .Offset(line_no, 0) = "Variable  Segment-->"
    If i = Nseg + 1 Then
        .Offset(line_no, i + 1) = "Mean"
        .Offset(line_no, i + 1).HorizontalAlignment = xlRight
        Else
        .Offset(line_no, i + 1) = i
        End If
        .Offset(line_no, 0).Font.Bold = True
        .Offset(line_no, i + 1).Font.Bold = True
        .Offset(line_no, 0).Font.Underline = True
        .Offset(line_no, i + 1).Font.Underline = True
       Next i

'       nord1 = 10
        nord1 = NDiagnostics
        For j = 1 To nord1
           i = Iord(j)
           kU = 0
           For k = 1 To kseG
             x(k) = 0
             
             Select Case io
             Case 1
                x(k) = Cest(k, i)
             Case 2
                x(k) = Cobs(k, i)
             Case 3
                x(k) = Ratv(Cobs(k, i), Cest(k, i))
             Case 4
                x(k) = CvCobs(k, i) * Cobs(k, i)
             Case 5
                 x(k) = Sqr(CvCest(k, i))
             End Select
            If x(k) > 0 Then kU = 1
            Next k

            If kU > 0 Then
                line_no = line_no + 1
                .Offset(line_no, 0) = DiagName(i)
                For k = 1 To kseG
                    If x(k) > 0 Then .Offset(line_no, k + 1) = x(k)
                    .Offset(line_no, k + 1).NumberFormat = "0.0"
                    Next k
             End If

        Next j

    Next io

End With

End Sub

        Sub List_Verify()
'c verify segment balances

        If Iop(1) > 0 Then
                j = 1
                Else
                j = 2
                End If

       StartSheet ("verify")
        line_no = 0
        With gLSht.Range("A4")
            Hdr.Range("header_verify").Copy
            gLSht.Paste Destination:=.Offset(0, 0)
            .Offset(0, 4) = VariableName(j)
            line_no = line_no + 2

        For i = 1 To Nseg
          'Call Balas(i, j, 2)
          Call Balas(i, j, 1)
          wbe = Ratv(Qt(5), Qt(9))
          be = Ratv(Bt(5), Bt(9))
            line_no = line_no + 1
            .Offset(line_no, 0) = i
            .Offset(line_no, 0).HorizontalAlignment = xlCenter
            .Offset(line_no, 1) = SegName(i)
            .Offset(line_no, 2) = Qt(9)
            .Offset(line_no, 2).NumberFormat = "0"
            .Offset(line_no, 3) = Qt(5)
            .Offset(line_no, 3).NumberFormat = "0"
            .Offset(line_no, 4) = wbe
            .Offset(line_no, 4).NumberFormat = "0.0%"
            .Offset(line_no, 5) = Bt(9)
            .Offset(line_no, 5).NumberFormat = "0"
            .Offset(line_no, 6) = Bt(5)
            .Offset(line_no, 6).NumberFormat = "0"
            .Offset(line_no, 7) = be
            .Offset(line_no, 7).NumberFormat = "0.0%"
            .Offset(line_no, 8) = Qadv(i)
            .Offset(line_no, 8).NumberFormat = "0"
            Next i

        End With

'          write(nout,12) i,SegName(i),qt(9),qt(5),wbe,bt(9),bt(5),be,
'     &      qadv(i)
' 12     format(i3,1x,a8,f11.2,f10.2,f6.2,f11.2,f10.2,f6.2,f10.2)
     
        End Sub
Function StrSp(j, s)
'returns string of fixed length j
    StrSp = Left(s & Space(j), j)
End Function

Sub List_Tree()
'List Segment & Tributary Network

SegName(0) = "Out of Reservoir"
StartSheet ("Segment Network")
line_no = 0
With gLSht.Range("A4")

.Offset(line_no, 0) = "Segment & Tributary Network"
.Offset(line_no, 0).Font.Bold = True

For i = 1 To Nseg

line_no = line_no + 2
.Offset(line_no, 0) = "--------Segment:"
.Offset(line_no, 0).HorizontalAlignment = xlRight
.Offset(line_no, 1) = i
.Offset(line_no, 1).HorizontalAlignment = xlCenter
.Offset(line_no, 2) = SegName(i)

line_no = line_no + 1
.Offset(line_no, 0) = "Outflow Segment:"
.Offset(line_no, 0).HorizontalAlignment = xlRight
.Offset(line_no, 1) = Iout(i)
.Offset(line_no, 1).HorizontalAlignment = xlCenter
.Offset(line_no, 2) = SegName(Iout(i))

For j = 1 To NTrib
    If Iseg(j) = i Then
    line_no = line_no + 1
    .Offset(line_no, 0) = "Tributary:"
    .Offset(line_no, 0).HorizontalAlignment = xlRight
    .Offset(line_no, 1) = j
    .Offset(line_no, 1).HorizontalAlignment = xlCenter
    .Offset(line_no, 2) = TribName(j)
    .Offset(line_no, 3) = "Type:"
    .Offset(line_no, 3).HorizontalAlignment = xlRight
    .Offset(line_no, 4) = Type_Code(Icode(j))
    End If
    Next j
    
For j = 1 To Npipe

    If Ifr(j) = i Then
    line_no = line_no + 1
    .Offset(line_no, 0) = "Channel:"
    .Offset(line_no, 0).HorizontalAlignment = xlRight
    .Offset(line_no, 1) = j
    .Offset(line_no, 1).HorizontalAlignment = xlCenter
    .Offset(line_no, 2) = PipeName(j)
    .Offset(line_no, 3) = " To Segment:"
    .Offset(line_no, 3).HorizontalAlignment = xlRight
    .Offset(line_no, 4) = Format(Ito(j), "00") & " " & SegName(Ito(j))
    End If
    Next j

Next i
End With
End Sub

Sub List_Tree2()
'List Segment & Tributary Network

Dim txt As String

SegName(0) = "Out of Reservoir"
txt = Title & vbCrLf
txt = txt & "Segment & Tributary Network" & vbCrLf & vbCrLf

For i = 1 To Nseg
txt = txt & "----Segment: " & Format(i, "00") & " " & SegName(i) & vbCrLf
txt = txt & "Outflow Seg: " & Format(Iout(i), "00") & " " & SegName(Iout(i)) & vbCrLf
For j = 1 To NTrib
    If Iseg(j) = i Then txt = txt + "  Tributary: " & Format(j, "00") & " " & StrSp(20, TribName(j)) & "   Type: " & Type_Code(Icode(j)) & vbCrLf
    Next j
For j = 1 To npipes
    If Ifr(j) = i Then txt = txt + "    Channel: " & Format(j, "00") & " " & StrSp(20, PipeName(j)) & "    To Segment: " & Format(Ito(j)) & " " & SegName(Ito(j)) & vbCrLf
    Next j
txt = txt + vbCrLf
Next i

With frmBox
    .txtBox.Text = txt
    .txtBox.SelStart = 0
    .Show vbModal
End With

End Sub
Sub List_inss()
    Status ("Inputs")
    Set gLSht = Wkb.Sheets("Inputs")
    Save_xls                 'save inputs to wkb
    If Ier > 0 Then Exit Sub
'    Set gLSht = Wkb.Worksheets("inputs")
    ViewSheet ("inputs")
    Status ("Ready")

End Sub
Sub List_Inputs()
Dim i As Long
Dim j As Long
Dim k As Long
StartSheet ("Case Data")

line_no = 0
With gLSht.Range("A3")

Hdr.Range("header_in_top").Copy .Offset(line_no, 0)

    For i = 1 To 10
        line_no = line_no + 1
        .Offset(line_no, 1) = Note(i)
        .Offset(line_no, 1).WrapText = False
        Next i
   
'global variables
   line_no = line_no + 2
   k = line_no
   For i = 1 To 4
        line_no = line_no + 1
        .Offset(line_no, 2) = P(i)
        .Offset(line_no, 3) = Cp(i)
        Next i

'atmospheric loads
    line_no = line_no + 2
    For i = 1 To NVariables
        line_no = line_no + 1
        .Offset(line_no, 2).Value = Atm(i)
        .Offset(line_no, 3).Value = CvAtm(i)
         Next i
    
'Model Options
    line_no = k
    For i = 1 To NOptions
        line_no = line_no + 1
        .Offset(line_no, 9) = Iop(i)
        .Offset(line_no, 9).HorizontalAlignment = xlCenter
        .Offset(line_no, 10) = OptionName(i, Iop(i) + 1)
        Next i

'segment list
    line_no = line_no + 4
    For i = 1 To Nseg
       line_no = line_no + 1
    .Offset(line_no, 0) = i
    .Offset(line_no, 0).HorizontalAlignment = xlCenter
    .Offset(line_no, 1) = SegName(i)
    .Offset(line_no, 3) = Iout(i)
    .Offset(line_no, 4) = Iag(i)
    .Offset(line_no, 5) = Area(i)
    .Offset(line_no, 6) = Zmn(i)
    .Offset(line_no, 7) = Slen(i)
    .Offset(line_no, 8) = Zmx(i)
    .Offset(line_no, 9) = CvZmx(i)
    .Offset(line_no, 10) = Zhyp(i)
    .Offset(line_no, 11) = CvZhyp(i)
    .Offset(line_no, 12) = Turbi(i)
    .Offset(line_no, 13) = CvTurbi(i)
    For j = 1 To 3
        .Offset(line_no, 12 + j * 2) = InternalLoad(i, j)
        .Offset(line_no, 12 + j * 2 + 1) = CvInternalLoad(i, j)
        Next j
    Next i
    line_no = line_no + 2

'segment observed water quality
  Hdr.Range("header_in_obswq").Copy .Offset(line_no, 0)
  line_no = line_no + 2
  For i = 1 To Nseg
    line_no = line_no + 1
    .Offset(line_no, 0) = i
    .Offset(line_no, 0).HorizontalAlignment = xlCenter
    For j = 1 To 9
        .Offset(line_no, -1 + j * 2) = Cobs(i, j)
        .Offset(line_no, 0 + j * 2) = CvCobs(i, j)
        Next j
    Next i
    
    line_no = line_no + 2

'calibration factors
Hdr.Range("header_in_calibfactors").Copy .Offset(line_no, 0)
  line_no = line_no + 2
  For i = 1 To Nseg
    line_no = line_no + 1
    .Offset(line_no, 0) = i
    .Offset(line_no, 0).HorizontalAlignment = xlCenter
    For j = 1 To 9
        .Offset(line_no, -1 + j * 2) = Cal(i, j)
        .Offset(line_no, 0 + j * 2) = CvCal(i, j)
        Next j
    Next i

'tributaries
line_no = line_no + 2
Hdr.Range("header_in_tribs").Copy .Offset(line_no, 0)
    line_no = line_no + 2
    For i = 1 To NTrib
    line_no = line_no + 1
    .Offset(line_no, 0) = i
    .Offset(line_no, 0).HorizontalAlignment = xlCenter
    .Offset(line_no, 1) = TribName(i)
    .Offset(line_no, 3) = Iseg(i)
    .Offset(line_no, 4) = Icode(i)
    .Offset(line_no, 5) = Darea(i)
    .Offset(line_no, 6) = Flow(i)
    .Offset(line_no, 7) = CvFlow(i)
    For j = 1 To NVariables
        .Offset(line_no, 6 + 2 * j) = Conci(i, j)
        .Offset(line_no, 7 + 2 * j) = CvCi(i, j)
        Next j
        Next i

'trib nonpoint areas
    t = 0
    For i = 1 To NTrib
    For j = 1 To NCAT
     t = t + Warea(i, j)
     Next j
     Next i
    
 If t > 0 Then
 line_no = line_no + 2
 Hdr.Range("header_in_areas").Copy .Offset(line_no, 0)
    line_no = line_no + 2
    For i = 1 To NTrib
    line_no = line_no + 1
    .Offset(line_no, 0) = i
    .Offset(line_no, 0).HorizontalAlignment = xlCenter
    .Offset(line_no, 1) = TribName(i)
    For j = 1 To NCAT
        .Offset(line_no, 2 + j) = Warea(i, j)
        Next j
     Next i
    End If
    
' Channels

    If Npipe > 0 Then
    line_no = line_no + 2
    Hdr.Range("header_in_transport").Copy .Offset(line_no, 0)
    line_no = line_no + 2
    For i = 1 To Npipe
        line_no = line_no + 1
        .Offset(line_no, 0) = i
        .Offset(line_no, 0).HorizontalAlignment = xlCenter
        .Offset(line_no, 1) = PipeName(i)
        .Offset(line_no, 3) = Ifr(i)
        .Offset(line_no, 4) = Ito(i)
        .Offset(line_no, 5) = Qpipe(i)
        .Offset(line_no, 6) = CvQpipe(i)
        .Offset(line_no, 7) = Epipe(i)
        .Offset(line_no, 8) = CvEpipe(i)
        Next i
     End If

'export categories
    t = 0
    For i = 1 To NCAT
        t = t + Ur(i)
        Next i
    
    If t > 0 Then
    line_no = line_no + 2
    Hdr.Range("header_in_export").Copy .Offset(line_no, 0)
    line_no = line_no + 2
    For i = 1 To NCAT
        line_no = line_no + 1
       .Offset(line_no, 0) = i
       .Offset(line_no, 14).HorizontalAlignment = xlCenter
       .Offset(line_no, 1) = LandUseName(i)
       .Offset(line_no, 3) = Ur(i)
       .Offset(line_no, 4) = CvUr(i)
       For k = 1 To NVariables
            .Offset(line_no, 3 + k * 2) = Uc(i, k)
            .Offset(line_no, 4 + k * 2) = CvUc(i, k)
            Next k
      Next i
      End If

'coefficients
    line_no = line_no + 2
    Hdr.Range("header_in_last").Copy .Offset(line_no, 0)
    For i = 1 To NXk
        line_no = line_no + 1
        .Offset(line_no, 3).Value = Xk(i)
        .Offset(line_no, 4).Value = CvXk(i)
        Next i

    End With
    End Sub

