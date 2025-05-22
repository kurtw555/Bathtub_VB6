Attribute VB_Name = "Module5"
'calibration routines
'variables accessible only to calibration module procedures
    
    Dim txt As String
    Dim ncoM As Integer
    Dim pcoM(50) As Single
    Dim xicoM(NSMAX) As Single
    Dim xT(NSMAX) As Single
    Dim jR As Integer
    Dim kC As Integer
    Dim nU As Integer
    Dim nuC As Integer
    Dim nD As Integer
    Dim deL As Single
    Dim iT As Integer
    Dim iuK(4) As Integer
    Dim iuC(5) As Integer
    Dim Jzap(NSMAX) As Integer
    

    Const nbuZ As Integer = 1
    Const buZ As Single = 0.1
    Const iwT As Integer = 0
    Const ptoL As Single = 0.01
    Const deL1 As Single = 0.01
    Const ptoL2 As Single = 0.001
    Const deL2 As Single = 0.001
    Const deL3 As Single = 0.1
    Const pdmaX As Single = 10

'1       'number of restarts'
'.1      'randomization range on restarts'
'0       'weighting scheme'
'.01     '.01 initial solution tolerance'
'.01     '.01 initial derivative increment'
'.001    '.001 final solution tolerance'
'.001    '.001 final derivative increment'
'.1      'variance increment'
'143     'random number seed'
'10.      'max increment was 8'

'fortran menu
'  c Solve Global  29
'       Call GlobalCalib(2)

'c Solve Regional  30
'       Call LocalCalib(1)

'c Solve Local  31
'       Call LocalCalib(0)
'c
'c List Each  33
'       Call List_Fits

'c List Summary  34
'       Call ListCalibFactors

'c Reset Global  36
'       Call GlobalCalib(0)
'c
'c Reset Local  37
'       Call LocalCalib(2)
'c
'c Calibrate Exchange  38
'       Call ExCo
         
'-------------------------
'        Sub Calib_Menu()
'control calibrations
'        StartUp
'        Run
'Top:
'        io = InputBox("enter 0=quit,1=list fits,2=solve global,3=list calib factors,4=solve local,5=reset global,6=reset local")
'        Select Case io
'            Case 0
'        Exit Sub
'            Case 1
'        List_Fits
'            Case 2
'        GlobalCalib (2)
'        List_Fits
'            Case 3
'        ListCalibFactors
'            Case 4
'        LocalCalib (0)
'        MsgBox "hold"
'        List_Fits
'        Case 5
'        GlobalCalib (0)
'        Model
'        ListCalibFactors
'        Case 6
'        LocalCalib (2)
'        Model
'        ListCalibFactors
'
'        End Select
'        GoTo Top
'        End Sub
       
        Sub List_Fits()
'c list model coefficients & calibration statistics
      If Icalc = 0 Then Exit Sub
          line_no = 0
          For i = 1 To 4
              If Iop(i) > 0 Then
                Call ListFitOne(i)
                End If
              Next i
                
        End Sub
        Sub ListFitOne(io)

'c list observed and predicted concs
'c io=1 cons, 2=tp, 3=tn, 4=chla, 5=secchi
        
'       line_no = 0
        

        If line_no = 0 Then
            StartSheet ("calibrations")
            Else
            line_no = line_no + 2
            End If
        With gLSht.Range("A3")
            line_no = line_no + 1
            Hdr.Range("header_calib").Copy
            gLSht.Paste Destination:=.Offset(line_no, 0)
        
        If Iop(io) > 0 Then
        Call Oss(io, 0, fF, r2, nobs)
            .Offset(line_no, 2) = DiagName(io)
            .Offset(line_no, 2).Font.Bold = True
            .Offset(line_no, 4) = r2
            .Offset(line_no, 4).NumberFormat = "0.00"
            .Offset(line_no + 1, 3) = Xk(io)
            .Offset(line_no + 1, 3).NumberFormat = "0.00"
            .Offset(line_no + 1, 5) = CvXk(io)
            .Offset(line_no + 1, 5).NumberFormat = "0.00"
        line_no = line_no + 3
        

        If Nseg = 1 Then
            m = 1
            Else
            m = Nseg + 1
            End If
        
        For i = 1 To m
          If Cobs(i, io) > 0 And Cest(i, io) > 0 And Izap(i) = 0 Then
            x(1) = Sqr(MAx(0, CvCest(i, io))) / Cest(i, io)
            x(2) = Log(Cobs(i, io) / Cest(i, io))
            x(3) = Sqr(CvCobs(i, io) ^ 2 + x(1) ^ 2)
            x(4) = Ratv(x(2), x(3))
        line_no = line_no + 1
        .Offset(line_no, 0) = i
        .Offset(line_no, 0).HorizontalAlignment = xlCenter
        .Offset(line_no, 1) = Iag(i)
        .Offset(line_no, 1).HorizontalAlignment = xlCenter
        .Offset(line_no, 2) = SegName(i)
        If i <= Nseg Then
            .Offset(line_no, 3) = Cal(i, io)
            .Offset(line_no, 3).NumberFormat = "0.00"
            .Offset(line_no, 4) = CvCal(i, io)
            .Offset(line_no, 4).NumberFormat = "0.00"
            End If
        .Offset(line_no, 5) = Cest(i, io)
        .Offset(line_no, 5).NumberFormat = "0.0"
        .Offset(line_no, 6) = x(1)
        .Offset(line_no, 6).NumberFormat = "0.00"
        If Cobs(i, io) > 0 Then
            .Offset(line_no, 7) = Cobs(i, io)
            .Offset(line_no, 7).NumberFormat = "0.0"
            .Offset(line_no, 8) = CvCobs(i, io)
            .Offset(line_no, 8).NumberFormat = "0.00"
            .Offset(line_no, 9) = x(2)
            .Offset(line_no, 9).NumberFormat = "0.00"
            .Offset(line_no, 10) = x(3)
            .Offset(line_no, 10).NumberFormat = "0.00"
            .Offset(line_no, 11) = x(4)
            .Offset(line_no, 11).NumberFormat = "0.00"
            End If
        

          End If
          Next i
       End If
       End With
       
        End Sub
   
        Sub Oss(js, iwT, osst, r2, nobs)
' weighted objective function for one variable, segment kc
' residual mean square - adjusted for number of parameters
        Dim f1 As Single
        Dim f2 As Single
        'Dim r2 As Single
        Dim f0 As Single
        Dim yo As Single
        Dim yye As Single
        
          osst = 0
          f0 = 0
          f1 = 0
          f2 = 0
                For j = 1 To Nseg
                  If (Cest(j, js) > 0 And Cobs(j, js) > 0 And Jzap(j) <= 0) Then
                  Weight = 1
                  If (iwT = 2 Or iwT = 4) Then Weight = Weight * Area(j)
                  If (iwT >= 3) Then Weight = Ratv(Weight, MAx(CvCobs(j, js), 0.001))
                  Weight = Sqr(Weight)
                  yo = Log(Cobs(j, js)) * Weight
                  yye = Log(Cest(j, js)) * Weight
                  osst = osst + (yye - yo) ^ 2
                  f0 = f0 + 1
                  f1 = f1 + yo
                  f2 = f2 + yo ^ 2
                  End If
                Next j
                f2 = f2 - Ratv(f1 * f1, f0)
                r2 = 1 - Ratv(osst, f2)
                nobs = f0
          
          End Sub

'    Sub Reset_Zap()
'        For i = 1 To Nseg
'           Izap(i) = 0
'           Next i
'   End Sub
      
       Sub GlobalCalib(jo)

'global calibration ?
'c jo=0 reset
'c jo=1 simultaneous
'c jo=2 series  <---- only one activated & debugged
'c jo=3 custom
'c jo=4 series custom
    
'!!!    common/opt/jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)

Select Case jo
'c Reset
Case 0
          For k = 1 To 5
                Xk(k) = 1
                Next k
             Icalc = 0
             'Call ListCalibFactors
             Exit Sub

Case 1
            nuC = 5
            nU = 0
            For k = 1 To 5
              iuC(k) = k
              If k = 1 And Nseg > 1 Then
                  nU = nU + 1
                  iuK(nU) = 13
                ElseIf Iop(k) > 0 And k <> 5 Then
                  nU = nU + 1
                  iuK(nU) = k - 1
                  End If
              Next k
           If nU > 0 Then Call LeastSquaresGlobal

Case 2

'c calibrate in series jo=2
'  SelectCalibVariables
'  If IndSum(Icoef, 4) = 0 Then Exit Sub
    
    txt = ""
    For k = 1 To 4
          If frmCalibration.chkVariable(k - 1) Then
               nU = 1
               iuK(1) = k                      'calibrated coefficients
               nuC = 1
               iuC(1) = k                       'calibrated concentrations
               Call LeastSquaresGlobal
               If Ier > 0 Then
                    MsgBox "Fatal error occured during calibration"
                    Icalc = 0
                    Exit Sub
                    End If
               End If
           Next k
           
'List_Fits   'list calibration results  not functional

   Case 3
'c jo=3 custom
        'Call SelVar(5)
        nuC = 0
        For i = 1 To 5
           If Iwork(i) > 0 Then
               nuC = nuC + 1
               iuC(nuC) = i
               End If
           Next i
        If nuC = 0 Then GoTo S99

       ' SelectCalibVariables
        nU = 0
        For i = 1 To 4
           If Icoef(i) > 0 Then
              nU = nU + 1
              iuK(nU) = i - 1
              If iuK(nU) < 1 Then iuK(nU) = 13
              If Xk(iuK(nU)) <= 0 Then Xk(iuK(nU)) = 1#
              End If
            Next i
        If nU = 0 Then GoTo S99
        Call LeastSquaresGlobal
    
    End Select
        
'99      Izap = 0
S99:
        End Sub

        Sub LeastSquaresGlobal()
'c calibrate global factors

'!!!        common/opt/jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)

'c k=1 exchange cobs(1)-->xk(13)
'c k=2 tp       cobs(2)-->xk(1)
'c k=3 tn       cobs(3)-->xk(2)
'c k=4 chla     cobs(4)-->xk(3)
    
    If Icalc = 0 Then
        Call Calcon
        If Ier > 0 Then Exit Sub
        End If
    
    Status ("Calibrating")
'c assign parameters
        For i = 1 To nU
           Xp(i) = Xk(iuK(i))
'c Log - transform
           If Xp(i) <= 0 Then Xp(i) = 1
           Xp(i) = Log(Xp(i))
           Yp(i) = Xp(i)
           Next i
             
'header for calibration progress window
       
       CalibHeader
       kC = 0   'this tells mincon that it is a global calibration
       Call MinCon
       txt = txt & vbCrLf
       'MsgBox "Calibration Complete"
        
    'Call Model
    'If Ier <= 0 Then
    ''    WriteCase
    '    Status ("Ready")
    '    Else
    '    Status ("Errors")
    '    End If
    End Sub
Sub CalibHeader()
'header for calibration progress window
       txt = txt & "Calibrated Variable:" & DiagName(iuC(1)) & vbCrLf
       'For i = 1 To nU
       '    txT = txT & VariableName(iuC(i)) & vbCrLf
       '    Next i
      'txT = txT & "Coefficients Calibrated:" & vbCrLf
      ' For i = 1 To nU
      '     txT = txT & XkName(iuK(i)) & vbCrLf
      '     Next i
       txt = txt & "Iter    Resid SS Coefficients..." & vbCrLf
End Sub
    
    Sub Dfunc(pp, dd, f0)

'c calculate derivatives numerically - designed for log-scale parameters

'    common /opt/jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)
'    DIMENSION pp(1), dd(1)

'       if(itrace > 0) write(23,23) 'p:',(pp(j),j=1,nu)
' 23    format(1x,a4,20f7.2)
    
    
    f0 = Func(pp)
    For j = 1 To nU
        plast = pp(j)
        If pp(j) = pdmaX Then
            pp(j) = pp(j) - deL
            Else
            pp(j) = pp(j) + deL
            End If
        fF = Func(pp)
        dd(j) = (fF - f0) / (pp(j) - plast)
        pp(j) = plast
       Next j
    nD = nD + 1
    'if(itrace > 0) write(23,23) 'd:', (dd(j),j=1,nu)

    If nD = 1 Then fF = 0
    
    txt = txt & FormatF(nD, "####") & FormatF(f0, "######.0000")
        For k = 1 To nU
            txt = txt & FormatF(Exp(pp(k)), " ####.00")
            Next k
        txt = txt & vbCrLf
    
    With frmCalibration.txtCalib
        .Text = txt
        .SelStart = 0
    End With

    fF = f0
    
    End Sub
    
    Function Func(pp)

'    common /opt/jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)
'    DIMENSION pp(1)

'c constrain values
       For i = 1 To nU
            pp(i) = MAx(MIn(pp(i), pdmaX), -pdmaX)
            Next i

'c swap in current parameter values
'   if(itrace > 0) write(23,24) (pp(i),i=1,nu)
' 24     format(' p(f):',20f7.2)

'c assign parameters
        For i = 1 To nU
           ppu = Exp(pp(i))
           If kC = 0 Then
             Xk(iuK(i)) = ppu
           Else
            If jR = 0 Then
'c by segment
              Cal(Jsg(i), kC) = ppu
              Else
'c by segment group
                For j = 1 To Nseg
                  If Jsg(i) = Iag(j) And Izap(j) = 0 Then Cal(j, kC) = ppu
                  Next j
              End If
             End If
           Next i

'c run model
        Call Model

'c calculate objective function
        Func = 0
        For i = 1 To nuC
          Call Oss(iuC(i), iwT, fF, r2, nobs)
          Func = Func + fF
          Next i
        
        End Function

'    Sub ResetLocalCalib(io)
'c reset local calibration factors
'c 0 = all, 1 = tp, 2 = tn, 3 = chla, 4 = secchi, 5 = oxygen, 6 = exchange

    'Iwork = 1
    'Dummy = "SELECT SEGMENTS TO BE RESET"
    
    
'    Call SelectMany(Nseg, SegName, "Select Segments to be Reset", 0, Iwork)
'    For j = 1 To 6
'      If Not (io > 0 And j <> io) Then
'    For i = 1 To Nseg
'      If Iwork(i) > 0 Then
'            Cal(i, io) = 1
'c CvCal(i, io) = 0#
'            End If
'      Next i
'    Icalc = 0
'    End If
'    Next j
'
'    End Sub

Function IndSum(iv, m)
'sum index
    IndSum = 0
    For i = 1 To m
    If iv(i) > 0 Then IndSum = IndSum + 1
    Next i
End Function
   
    Sub LocalCalib(io)
    
'c control least-squares solutions
'c jr=0 by segment, jr=1 by group, jr=2 reset

'   common/opt/jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)

        'jR = io
'        If jr <> 2 And Icalc <= 0 Then
'             Call helpm(56, i, j)
'             Exit Sub
'             End If

'c select coefficients
'        SelectCalibVariables
'        If IndSum(Icoef, 4) <= 0 Then Exit Sub

'c select segments
'        Call SelectMany(Nseg, SegName, "Select Segments to be Calibrated", 0, Iwork)
'        If IndSum(Iwork, Nseg) <= 0 Then Exit Sub
'        For j = 1 To Nseg
'            Izap(j) = 1 - Iwork(j)
'            Next j

'c calibrate
   '     For j = 1 To 4
    '       iuK(j) = Icoef(j)
    '       Next j
        
    txt = ""
        For j = 1 To 4
               If frmCalibration.chkVariable(j - 1) Then
               Call LeastSquaresLocal(j, io)
               If Ier > 0 Then
                    MsgBox "Calibration Failed"
                    Icalc = 0
                    Exit Sub
                    End If
                End If
            Next j
        
        End Sub

    Sub LeastSquaresLocal(jc, kr)
'c least-squares solution for segment-specific values
'c jc=1 exchange, jc=2 p decay, jc=3 n decay, jc=4 chla
'c jr=0 each segment, jr=1 by lake region,jr=3 region

'    common /opt/ jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)

        jR = kr
        
'c set observed values
        nuC = 1
        iuC(1) = jc

'c calibrate coefficients (cal(i,kc))
        kC = jc
  '     If jc = 1 Then
  '         kC = 6
  '         Else
  '         kC = jc - 1
  '         End If

'c Reset
'      If jR = 2 Then
'           For i = 1 To Nseg
'             Cal(i, kC) = 1
'             Next i
'            Icalc = 0
'            Exit Sub
'            End If

'      If Iop(jc) <= 0 Then Exit Sub

'c set coefficients - select segments
    nU = 0
    For i = 1 To Nseg

'c reject if zapped
'          If Izap(i) > 0 Then cycle

'c dont use last segment if chloride
'          If Iout(i) = 0 And jc = 1 Then cycle

'c dont use segments without observed concs
'      If Cobs(i, jc) <= 0 Then cycle

'c dont use segments calibrated to zero
'          If Cal(i, kc) <= 0 Then cycle
   
    Jzap(i) = 1
    If frmCalibration.List1.Selected(i - 1) = True And Cobs(i, jc) > 0 And Cal(i, kC) > 0 Then
    Jzap(i) = 0
    
    If Iout(i) > 0 Or jc <> 1 Then
'c by segment group
         If jR > 0 Then
               For j = 1 To nU
                    If Jsg(j) = Iag(i) Then
                       Cal(i, kC) = Exp(Xp(j))
                       GoTo s35
                       End If
                Next j
           nU = nU + 1
           Jsg(nU) = Iag(i)
            Else
'c by segment
            nU = nU + 1
            Jsg(nU) = i
            End If
'???
        Xp(nU) = Log(Cal(i, kC))
        Yp(nU) = Xp(nU)
    End If
    End If

s35:
    Next i
    
    If nU = 0 Then
            MsgBox "There are no observed water quality data for the selected segments"
            Icalc = 0
            Else
   
    CalibHeader
    Call MinCon
            
'   MsgBox "calibration complete"
'    SetXlw (1)
    
     End If
'       Reset_Zap
        
     End Sub

Sub ListCalibFactors()

'c list calibration factors
    If Icalc = 0 Then Exit Sub
    
        line_no = 0
        Sheets("CalibFactors").Activate
        With Range("A1")
        .Range("A1:z5000").Clear
        Range("header_listcal").Copy
        ActiveSheet.Paste Destination:=.Offset(0, 0)
 
'       write(nout,13)
' 13     format(' Calibration Factors:'/
'     & ' Seg Grp Label           ',
'     &       '  Exchange   Total P   Total N    Chl-a')
        line_no = line_no + 3
        .Offset(line_no, 0) = "Global"
        For i = 1 To 3
            .Offset(line_no, i + 3) = Xk(i)
            .Offset(line_no, i + 3).NumberFormat = "0.00"
            Next i
'         write(nout,14) p(6),(xk(i),i=1,3)
' 14     format(' Global ',17x,4f10.4)
        line_no = line_no + 1
        .Offset(line_no, 0) = "CV:"
         For i = 1 To 3
            .Offset(line_no, i + 3) = CvXk(i)
            .Offset(line_no, i + 3).NumberFormat = "0.00"
            Next i
'        write(nout,19) cvxk(13),(cvxk(i),i=1,3)
' 19      format(22x,'Cv:',4f10.4)
         
         For i = 1 To Nseg
'          write(nout,15) i,iag(i),SegName(i),cal(i,6),(cal(i,j),j=1,3)
'          write(nout,19) cvcal(i,6),(cvcal(i,j),j=1,3)
' 15     format(2i4,1x,a16,4f10.4)
        line_no = line_no + 1
        .Offset(line_no, 0) = i
        .Offset(line_no, 0).HorizontalAlignment = xlCenter
        .Offset(line_no, 1) = Iag(i)
        .Offset(line_no, 1).HorizontalAlignment = xlCenter
        .Offset(line_no, 2) = SegName(i)
        .Offset(line_no, 3) = Cal(i, 1)
        .Offset(line_no, 3).NumberFormat = "0.00"
        .Offset(line_no, 4) = Cal(i, 2)
        .Offset(line_no, 4).NumberFormat = "0.00"
        .Offset(line_no, 5) = Cal(i, 3)
        .Offset(line_no, 5).NumberFormat = "0.00"
        .Offset(line_no, 6) = Cal(i, 4)
        .Offset(line_no, 6).NumberFormat = "0.00"
          Next i
            
        line_no = line_no + 2
        .Offset(line_no, 0) = "R-Squared:"
        
        For i = 1 To 4
          Call Oss(i, 0, fF, xx, nobs)
          .Offset(line_no, i + 2) = xx
          .Offset(line_no, i + 2).NumberFormat = "0.00"
          Next i
            
'        write(nout,16) (100.*x(i),i=1,4)
' 16     format(/' R-Squared: ',13x,10f10.1)
'        MsgBox "hold"
    
        End With
        End Sub

  Sub ExCo()
 
'c back-solve for exchange terms

'c check
'      if(icalc <= 0) then
'                Call helpm(56, i, j)
'        Return
'        End If

    Dim isym(100) As String
    

      If Iop(1) <= 0 Then Return
      Icalc = 0
      line_no = 0
      
      Sheets("work").Activate
      With Sheets("work").Range("A1")

'c left-hand side = external inputs + advective inputs - advective outputs
S56:
        For i = 1 To Nseg

'c cancel rows - outflow segments + zero calib
       If Iout(i) = 0 Or Cal(i, 1) = 0 Then
         isym(i) = "*"
           Else
         isym(i) = " "
         Cest(i, 1) = Cobs(i, 1)
           End If
'c target
     Yp(i) = Cest(i, 1)

'c mass-balance terms
        Call Balas(i, 1, 2)      '??????

'c left-hand side = external inputs + net advective inputs
    Xp(i) = Bt(1) - Bt(11) + Bt(2) + Bt(6) - Bt(7) - Bt(4)

    For j = 1 To Nseg
          A(j, i) = 0
          Next j

'c concentration gradient on right hand side
    j = Iout(i)
       If j > 0 Then
        delta = Cest(i, 1) - Cest(j, 1)
        A(i, i) = delta
        A(j, i) = -delta
        End If
        Next i

'c eliminate segments with null diagonal term
      For i = 1 To Nseg
        If A(i, i) = 0 Or Cal(i, 1) = 0 Then
        For j = 1 To Nseg
            A(j, i) = 0
            A(i, j) = 0
            Next j
        Xp(i) = 0
        A(i, i) = 1
        End If
        Next i

'c solve model
'      Call blin(Dx, A, Xp, Nseg, Ml, Mu, Isym, Ier)
    For i = 1 To Nseg
        A(i, Nseg + 1) = Xp(i)
        Next i
    Call SolveLinear
    For i = 1 To Nseg
        Xp(i) = A(i, Nseg + 1)
        Next i

     If Ier > 0 Then
     MsgBox "solution failed - resetting coefficients"
'       write(nout,*) 'solution failed - resetting coefficients'
        Exit Sub
        End If

'c find most negative solution & constrain to zero
      XMIN = 0#
      imin = 0
      For i = 1 To Nseg
        If Xp(i) < XMIN Then
         XMIN = Xp(i)
         imin = i
         End If
         Next i

 '1         format(21x,
 '    &       '  DIFFUSIVE EXCHANGE TERMS (HM3/YR)',
 '    &       ' CONS SUBSTANCE CONCS' /
 '    &       ' SEGMENT              ',
 '    &  'EXCHANGE (HM3/YR)     CALIB    OBS    EST OBS-EST')
                  
      For i = 1 To Nseg
'c if solution <0, set =0.
       Xp(i + Nseg) = Exch(i)
       If i = imin Then
         Cal(i, 1) = 0
'c if solution >0, recalibrate
       ElseIf Xp(i) > 0 Then
        Cal(i, 1) = Cal(i, 1) * Ratv(Xp(i), Exch(i))
        End If
       Next i

'cc     call masbal(0,iter)
'cc     call masbal(1,iter)
        Call Model
'        If Ier > 0 Then
'            Call ResetLocalCalib(6)
'            Exit Sub
'            End If
      
      Range("header_exchange").Copy
      ActiveSheet.Paste Destination:=.Offset(line_no, 0)
'c Results
'       write(nout,1)
       For i = 1 To Nseg
'cc     write(nout,2) i,iout(i),SegName(i),xp(i+nseg),xp(i),cal(i,6),
'cc     &        cobs(i,1),yp(i),cest(i,1),iwork(i)
        line_no = line_no + 1
        .Offset(line_no, 0) = i
        .Offset(line_no, 1) = Iout(i)
        .Offset(line_no, 2) = SegName(i)
        .Offset(line_no, 3) = Exch(i)

        .Offset(line_no, 4) = Cal(i, 1)
        .Offset(line_no, 5) = Cobs(i, 1)
        .Offset(line_no, 6) = Cest(i, 1)
        .Offset(line_no, 7) = Cobs(i, 1) - Cest(i, 1)
        .Offset(line_no, 8) = isym(i)
        For k = 3 To 7
            .Offset(line_no, k).NumberFormat = "0.0"
            Next k

' write(nout,2) i,iout(i),SegName(i),exch(i),cal(i,6),
'     &        cobs(i,1),cest(i,1),cobs(i,1)-cest(i,1),iwork(i)
' 2         format(2i3,1x,a16,f16.2,f10.4,3f7.2,1x,a1)
       Next i
      line_no = line_no + 1
      .Offset(line_no, 0) = "* = constrained > 0 or excluded from calibration"
' write(nout,22)
' 22       format(
'     &    ' * = constrained >=0 or excluded from calibration')
      'MsgBox "hold"
      
      If imin > 0 Then GoTo S56

'c successful
      MsgBox "calibration successful"
      Call Model
  '    If Ier <= 0 Then Icalc = 1
      End With
      
      End Sub
  
'  Sub Rmat(io)        'disabled
  
'c unit response to increase in load
'c io=1 cs, io=2 total p
'$include:'net.inc'
'        character*3 iunit(2) /'ppb','%'/
'        if(icalc <= 0 or iop(io) <= 0) return
'        for i=1,nseg
'           Ysave(i) = Cest(i, io)
'           end do
'        write(nout,8) cname(io), term(io+1,9)
' 8        format(' Variable: ',a16,', Total Load =',f12.1)
'        dfac=xinp(0,'enter load increase to be tested ?   ')
'        if(dfac <= 0.) return
'        ik=iinp(0,'enter output units < 0=ppb, 1=% > ?   ')

'        Key = Blank
'        Call oswap(0, nout, Key)
'        write(nout,12) cname(io),iunit(ik+1),dfac
' 12     format(
'     & ' Load Transfer Matrix, Variable = ',a16/
'     & ' Increase in Conc (',a3,')  for',f12.1,
'     & ' Increase in Load')
'         write(nout,11) (i,i=1,nseg)
' 11      format(
'     &   22x,'Response In Segments --->'/
'     &   ' Tested Segments    ',15i4/20x,15i4/20x,15i4)
'        write(nout,7) (ifix(ysave(i)),i=1,nseg)
'  7      format(
'     &   ' Baseline Concs:    ',15i4/20x,15i4/20x,15i4)
'        NTrib = NTrib + 1
'        Flow(N) = 0.000001
'        Conci(N, Imap(io)) = dfac / Flow(N)
'        Icode(N) = 1
'        for i=1,nseg
'          Iseg(N) = i
'          Call Model(Iter)
'          for j=1,nseg
'             X(j) = Cest(j, io) - Ysave(j)
'             if(ik > 0) x(j)=100.*ratv(x(j),ysave(j))
'             end do
'       write(nout,13) i,SegName(i),(ifix(x(j)),j=1,nseg)
'13     format(i3,1x,a16,15i4/20x,15i4/20x,15i4)
'       end do
'        NTrib = NTrib - 1
'        Icalc = 0
'        Call Model
'        Icalc = 1
'        End Sub

Function Runi1()
    Runi1 = Application.WorksheetFunction.rand()
    End Function

Sub MinCon()

'c control nonlinear search
'c kc=0 global, >0 segment

'    common/opt/jr,kc,nu,nuc,nd,del,it,iwt,pdmax,iuk(4),iuc(5)
    'integer*4 i4

'    open(23,file='mincon.dat',status='old')
'    read(23,*) nbuz,dummy,buz,dummy,iwt,dummy,
'     &         ptol,dummy,del1,dummy,
'     &         ptol2,dummy,del2,dummy,del3,dummy,
'     &         i4,dummy,pdmax,dummy
'    Close (23)

    nD = 0
S33:
    If nU <= 0 Then Exit Sub
    
    Icalc = 0
    deL = deL1
    
'c first search for global scale factor giving best overall solution
'c searches exp(-5) to exp(+5)
        
        fmin = 1E+20
        For i = 1 To nU
         Yp(i) = Xp(i)
         Next i
        For k = 1 To 11
          For i = 1 To nU
             Xp(i) = Yp(i) + k - 6
             Next i
          fF = Func(Xp)
        
'        line_no = line_no + 1
'        .Offset(line_no, 0) = "map:"
'        .Offset(line_no, 1) = k
'        .Offset(line_no, 2) = Ff
'          MsgBox "map: " & k & " " & ff
'c        write(*,*) 'map:',k,ff
          
          If fF < fmin Then
                km = k
                fmin = fF
                End If
           Next k
        For i = 1 To nU
            Xp(i) = Yp(i) + km - 6
            Next i

'c Search
'        write(nout,16)
' 16     format(' Press <ESC> to Quit'/
'     &  ' Run Der  Resid SS   Coefficients......')

'c solution
    fmin = 1E+20
'c        call clrkey
'c        call bufset(0)
    For iT = 1 To nbuZ
'c check for <esc>
'        Call inkeyf(iasc, ikey)
'         if(ikey = 1) goto 333
'c randomize start
    For i = 1 To nU
      If iT > 1 Then Xp(i) = Yp(i) + (Runi1() - 0.5) * buZ
      Next i
'    if(itrace > 0) write(23,123) 'res:',it,(xp(i),i=1,nu)
' 123    format(1x,a4,i4,15f10.3)

    nD = 0
'c solution
    Call DFPMI2(Xp, nU, ptoL, itj, freT)

'    if(itrace > 0) write(23,123) 'sol:',it,(xp(i),i=1,nu)
    If freT < fmin Then
          fmin = freT
          For i = 1 To nU
            Yp(i) = Xp(i)
            Next i
          End If

    Next iT

'c one for good measure
    nD = 0
    deL = deL2
    Call DFPMI2(Yp, nU, ptoL2, itj, freT)

S333:
'        Iop(9) = 0    '???????
'        Call Eran(0)   '?????
        Call Model

'c calculate variance of parameter estimates
'c based upon quadratic approximation of response surface
'c f0 = fret
'c        call oss(kc,iwt,ff,r2,nobs)
'c dof = amax1(1#, nobs - nu)
'c        for i=1,nu
'c p0 = Yp(i)
'c Yp(i) = amin1(p0 + del3, pdmax)
'c p1 = Yp(i)
'c f1 = Func(Yp)
'c                yp(i)=amax1(p0-del3,-pdmax)c
'c f2 = Func(Yp)
'c       p2=yp(i)c
'c Yp(i) = p0
'c       v1=ratv((p1-p0)**2,f1-f0)
'c       v2=ratv((p2-p0)**2,f2-f0)
'c Xp(i) = sqrt(amax1(nu * f0 / dof * (v1 + v2) / 2#, 0#))
'c         write(nout,17) i,p0,f0,f1,f2,xp(i)
'c 17      format(' i,p,f0,f1,f2,cv:',i3,5g10.4)
'c           end do
'c        write(nout,15) (xp(i),i=1,nu)
'c 15       format(' error cvs:',8x,5g10.4/(19x,5g10.4))

'c map error sum of squares for each parameter
'c del4 = 0.1
'c         for i=1,nu
'c           write(nout,*) 'parameter cv =',xp(i)
'c p0 = Yp(i)
'c           for k=1,9
'c Yp(i) = p0 + del4 * (k - 5)
'c f1 = Func(Yp)
'c              write(nout,134) i,yp(i),exp(yp(i)),f0,f1,f1-f0
'c 134         format(i4,5g10.4)
'c enddo
'c Yp(i) = p0
'c           end do

    End Sub


      Sub DFPMI2(p2, n2, ftoL, iteR2, freT)
'c modified by www
'c convergence test based on parameter increment
'      PARAMETER (NMAX=50,ITMAX=200,EPS=1.E-10)
'      Parameter (DPMAX = 0.3)
'      DIMENSION P(N),psave(NMAX),
'     &HESSIN(NMAX,NMAX),XI(NMAX),G(NMAX),DG(NMAX),
'     *HDG(NMAX)
    nmax = 50
    itmax = 200
    eps = 0.0000000001
    dpmax = 0.3
    Dim psave(50) As Single
    Dim hessin(50, 50) As Single
    Dim xi(50) As Single
    Dim g(50) As Single
    Dim dg(50) As Single
    Dim hdg(50) As Single
    
'c not needed since dfunc returns function
'c FP = Func(P)
      Call Dfunc(p2, g, fP)
      
      For i = 1 To n2
      psave(i) = p2(i)
      For j = 1 To n2
          hessin(i, j) = 0
        Next j
        hessin(i, i) = 1
        xi(i) = -g(i)
        Next i

      For itS = 1 To itmax
        iteR2 = itS
        Call LinMin(p2, xi, n2, freT)

'c check for maximum delta p
    For i = 1 To n2
       dp = MAx(MIn(p2(i) - psave(i), dpmax), -dpmax)
       p2(i) = psave(i) + dp
       Next i

'c old convergence test
'c        IF(2.*ABS(FRET-FP).LE.FTOL*(ABS(FRET)+ABS(FP)+EPS))RETURN
'c new convergence test
     For i = 1 To n2
       If (Abs(p2(i) - psave(i)) > ftoL) Then GoTo S122
       Next i
       Exit Sub
S122:
        fP = freT
        For i = 1 To n2
            psave(i) = p2(i)
          dg(i) = g(i)
          Next i

'c freT = Func(P)
        Call Dfunc(p2, g, freT)
        For i = 1 To n2
          dg(i) = g(i) - dg(i)
            Next i

        For i = 1 To n2
          hdg(i) = 0
          For j = 1 To n2
            hdg(i) = hdg(i) + hessin(i, j) * dg(j)
            Next j
            Next i

        faC = 0
        faE = 0
        For i = 1 To n2
          faC = faC + dg(i) * xi(i)
          faE = faE + dg(i) * hdg(i)
          Next i

'c check for negative
        faC = 1 / faC
        faD = 1 / faE
        For i = 1 To n2
          dg(i) = faC * xi(i) - faD * hdg(i)
          Next i

        For i = 1 To n2
          For j = 1 To n2
            hessin(i, j) = hessin(i, j) + faC * xi(i) * xi(j) - faD * hdg(i) * hdg(j) + faE * dg(i) * dg(j)
            Next j
            Next i

        For i = 1 To n2
          xi(i) = 0
          For j = 1 To n2
            xi(i) = xi(i) - hessin(i, j) * g(j)
            Next j
            Next i
'24    CONTINUE
      Next itS
      MsgBox "too many iterations in DFPMIN"
      
      End Sub

      Sub LinMin(p2, xi, n2, freT)
'      PARAMETER (NMAX=50,TOL=1.E-4)
'      EXTERNAL F1DIM
'      DIMENSION P(N), XI(N)
'      COMMON /F1COM/ NCOM,PCOM(NMAX),XICOM(NMAX)
      nmax = NSMAX
      Tol2 = 0.0001
      
      ncoM = n2
      For j = 1 To n2
        pcoM(j) = p2(j)
        xicoM(j) = xi(j)
        Next j
      aX = 0
      xx = 1
      bX = 2
      Call MNBRAK(aX, xx, bX, fA, fX, fB, F1dimX)
      freT = Brent(aX, xx, bX, F1dimX, Tol2, XMIN)
      For j = 1 To n2
        xi(j) = XMIN * xi(j)
        p2(j) = p2(j) + xi(j)
        Next j
      End Sub
     
      
      Function Sign2(aa, bb)
 'was SIGN transfers sign of bb to aa and assigns results to sign2
      If bb < 0 Then
        Sign2 = -Abs(aa)
        Else
        Sign2 = Abs(aa)
        End If
        End Function
            
      Sub MNBRAK(aX, bX, cX, fA, fB, fC, funcXZ)
'      PARAMETER (GOLD=1.618034, TINY=1.E-20)
'c was glimit=100
'      Parameter (GLIMIT = 10#)
        glimit = 10
        gold = 1.618034
        tiny = 1E-20
      
     ' fA = Func(aX)
     ' fB = Func(bX)
     fA = F1dim(aX)
     fB = F1dim(bX)
      If (fB > fA) Then
        dDum = aX
        aX = bX
        bX = dDum
        dDum = fB
        fB = fA
        fA = dDum
      End If
      cX = bX + gold * (bX - aX)
      fC = F1dim(cX)
S1:
      If (fB >= fC) Then
        r = (bX - aX) * (fB - fC)
        q2 = (bX - cX) * (fB - fA)
        u = bX - ((bX - cX) * q2 - (bX - aX) * r) / (2 * Sign2(MAx(Abs(q2 - r), tiny), q2 - r))
        uliM = bX + glimit * (cX - bX)
        If ((bX - u) * (u - cX) > 0) Then
          fU = F1dim(u)
          If (fU < fC) Then
            aX = bX
            fA = fB
            bX = u
            fB = fU
            GoTo S1
          ElseIf (fU > fB) Then
            cX = u
            fC = fU
            GoTo S1
          End If
          u = cX + gold * (cX - bX)
          fU = F1dim(u)
        ElseIf ((cX - u) * (u - uliM) > 0) Then
          fU = F1dim(u)
          If (fU < fC) Then
            bX = cX
            cX = u
            u = cX + gold * (cX - bX)
            fB = fC
            fC = fU
            fU = F1dim(u)
          End If
        ElseIf ((u - uliM) * (uliM - cX) >= 0) Then
          u = uliM
          fU = F1dim(u)
        Else
          u = cX + gold * (cX - bX)
          fU = F1dim(u)
        End If
        aX = bX
        bX = cX
        cX = u
        fA = fB
        fB = fC
        fC = fU
        GoTo S1
      End If

      End Sub

      Function Brent(aX, bX, cX, F, Tol3, XMIN)
'      PARAMETER (ITMAX=100,CGOLD=.3819660,ZEPS=1.0E-10)
    itmax = 100
    cgold = 0.381966
    zeps = 0.0000000001
      Dim a2 As Single
      
      a2 = MIn(aX, cX)
      b = MAx(aX, cX)
      v = bX
      w = v
      x1 = v
      e2 = 0
      fX = F1dim(x1)
      Fv2 = fX
      Fw = fX
      For iteR2 = 1 To itmax
        xM = 0.5 * (a2 + b)
        Tol1 = Tol3 * Abs(x1) + zeps
        Tol2 = 2# * Tol1
        If (Abs(x1 - xM) <= (Tol2 - 0.5 * (b - a2))) Then GoTo S3
        If (Abs(e2) > Tol1) Then
          r = (x1 - w) * (fX - Fv2)
          q2 = (x1 - v) * (fX - Fw)
          p2 = (x1 - v) * q2 - (x1 - w) * r
          q2 = 2 * (q2 - r)
          If q2 > 0 Then p2 = -p2
          q2 = Abs(q2)
          Etemp = e2
          e2 = d
          If (Abs(p2) >= Abs(0.5 * q2 * Etemp) Or p2 <= q2 * (a2 - x1) Or p2 >= q2 * (b - x1)) Then GoTo S1
          d = p2 / q2
          u = x1 + d
          If (u - a2 < Tol2 Or b - u < Tol2) Then d = Sign2(Tol1, xM - x1)
          GoTo S2
        End If
S1:
       If (x1 >= xM) Then
          e2 = a2 - x1
        Else
          e2 = b - x1
        End If
        d = cgold * e2
S2:
       If (Abs(d) >= Tol1) Then
          u = x1 + d
        Else
          u = x1 + Sign2(Tol1, d)
        End If
        fU = F1dim(u)
        If (fU <= fX) Then
          If (u >= x1) Then
            a2 = x1
          Else
            b = x1
          End If
          v = w
          Fv2 = Fw
          w = x1
          Fw = fX
          x1 = u
          fX = fU
        Else
          If (u < x1) Then
            a2 = u
          Else
            b = u
          End If
          If (fU <= Fw Or w = x1) Then
            v = w
            Fv2 = Fw
            w = u
            Fw = fU
          ElseIf (fU <= Fv2 Or v = x1 Or v = w) Then
            v = u
            Fv2 = fU
          End If
        End If
        Next iteR2

      MsgBox "Brent optimization routine exceed maximum iterations"
S3:
      XMIN = x1
      Brent = fX
      
      End Function
    Function F1dim(x2)
'      Parameter (NMAX = 50)
'      COMMON /F1COM/ NCOM,PCOM(NMAX),XICOM(NMAX)
 
      
      For j = 1 To ncoM
        xT(j) = pcoM(j) + x2 * xicoM(j)
        Next j
      F1dim = Func(xT)
      End Function
