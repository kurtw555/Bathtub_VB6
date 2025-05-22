Attribute VB_Name = "Module7"
'seg & tributary edits
    Sub SegmentEdit(jseG, io)

'c io=0 delete, io=1 insert, io=2 copy, a segment
             
        Select Case io

'c delete segment
        Case 0
'c first update segment references
            If Nseg <= 1 Then Exit Sub
               For i = 1 To Nseg
                    If Iout(i) = jseG Then Iout(i) = Iout(jseG)
                    Next i
               
               For i = 1 To Nseg
                 If (Iout(i) > jseG) Then Iout(i) = Iout(i) - 1
                 Next i

'c now copy segments after deleted one
                For i = jseG To Nseg - 1
                   Call SegCopy(i + 1, i)
                   Next i
                Call SegZero(Nseg)
                Nseg = Nseg - 1
                
'c update tributary refs
              For i = 1 To NTrib
              If Iseg(i) = jseG Then
                    MsgBox "Tributary " & i & " " & TribName(i) & " Assigned to Segment 0"
                    Iseg(i) = 0
                  ElseIf (Iseg(i) > jseG) Then
                    Iseg(i) = Iseg(i) - 1
                  End If
               Next i

'c update pipe references
           For i = 1 To Npipe
                 If Ifr(i) = jseG Or Ito(i) = jseG Then
                    MsgBox "Channel " & i & " " & PipeName(i) & " Assigned to Segment 0"
                    Ifr(i) = 0
                    Ito(i) = 0
                    End If
                 If (Ifr(i) > jseG) Then Ifr(i) = Ifr(i) - 1
                 If (Ito(i) > jseG) Then Ito(i) = Ito(i) - 1
                 Next i

'c insert segment
        Case 1

'c first update segment references
            If Nseg >= NSMAX - 1 Then Exit Sub
             For i = 1 To Nseg
                   If (Iout(i) > jseG) Then Iout(i) = Iout(i) + 1
                   Next i

'c insert a new segment after jseg
               j = Nseg + 1
               For i = jseG + 1 To Nseg
                   j = j - 1
                    Call SegCopy(j, j + 1)
                    Next i
  '             Call SegCopy(NSMAX, jseg + 1)
               Call SegZero(jseG + 1)
               SegName(jseG + 1) = "????????"
               Nseg = Nseg + 1

'c update trib references
               For i = 1 To NTrib
                    If (Iseg(i) > jseG) Then Iseg(i) = Iseg(i) + 1
                Next i

'c update pipe references
         For i = 1 To Npipe
               If (Ifr(i) > jseG) Then Ifr(i) = Ifr(i) + 1
               If (Ito(i) > jseG) Then Ito(i) = Ito(i) + 1
               Next i

            End Select

             End Sub

           Sub SegCopy(i, j)
           
'c copy segment data from segment i to segment j

            Area(j) = Area(i)
            Zmn(j) = Zmn(i)
            Zmx(j) = Zmxi(i)
            SegName(j) = SegName(i)
            Iout(j) = Iout(i)
            Iag(j) = Iag(i)
            Zhyp(j) = Zhyp(i)
            Slen(j) = Slen(i)
            CvZmxi(j) = CvZmxi(i)
            CvZhyp(j) = CvZhyp(i)
            Turbi(j) = Turbi(i)
            CvTurbi(j) = CvTurbi(i)
            Targ(j) = Targ(i)

            For k = 1 To 7
                CvCal(j, k) = CvCal(i, k)
                Cal(j, k) = Cal(i, k)
                InternalLoad(j, k) = InternalLoad(i, k)
                CvInternalLoad(j, k) = CvInternalLoad(i, k)
                Next k

            For k = 1 To 9
                Cobs(j, k) = Cobs(i, k)
                CvCobs(j, k) = CvCobs(i, k)
                Next k
            
            End Sub
            Sub SegZero(j)
'c zero data for segment j

            Area(j) = 0
            Zmn(j) = 0
            Zmxi(j) = 0
            SegName(j) = ""
            Iout(j) = 0
            Iag(j) = 1
            Zhyp(j) = 0
            Slen(j) = 0
            CvZmxi(j) = 0
            CvZhyp(j) = 0
            Turbi(j) = 0
            CvTurbi(j) = 0
            Targ(j) = 0
            Izap(j) = 0

            For k = 1 To 7
                InternalLoad(j, k) = 0
                CvInternalLoad(j, k) = 0
                Next k

            For k = 1 To 9
                CvCal(j, k) = 0
                Cal(j, k) = 1
                Cobs(j, k) = 0
                CvCobs(j, k) = 0
                Next k
            
            End Sub
             Sub TribEdit(jtriB, io)
             
'c delete, insert, copy, a tribtary

             
'             write(*,11) (i,tname(i),i=1 to Ntrib)
' 11          format(' CURRENT TRIBUTARY LIST:'/ 3(i3,' = ',a14))
'             write(*,*)
'             If (io <= 0) Then
'                write(*,*) 'SELECT TRIBUTARY  TO DELETE'
'             else if(io = 1) then
'                write(*,*) 'INSERT NEW TRIBUTARY AFTER '
'             Else
'                write(*,*) 'SELECT TRIBUTARY TO COPY FROM '
'             End If

'             jtrib=iinp(0,'enter tributary , <0> to quit ?   ')
'             if(inbet(jtrib,1 to Ntrib) <= 0) return
             
       Select Case io

'c delete trib
       Case 0
                If NTrib <= 1 Then Exit Sub
                For i = jtriB To NTrib - 1
                    Call TribCopy(i + 1, i)
                    Next i
                Call TribZero(NTrib)
                NTrib = NTrib - 1
       Case 1

'c insert a new tributary after jtrib
               If NTrib >= NTMAX - 1 Then Exit Sub
               j = NTrib + 1
               For i = jtriB + 1 To NTrib
                   j = j - 1
                   Call TribCopy(j, j + 1)
                   Next i
       '        Area(NSMAX) = 0
       '        Zmn(NSMAX) = 0
       '        Call Tcopy(NTMAX, jtrib + 1)
               Call TribZero(jtriB + 1)
               TribName(jtriB + 1) = "????????"
               NTrib = NTrib + 1

'c copy a number of tribs

'              write(*,41) jtrib
'  41          format(' copy trib ',i3, ' to tribs  A thru B')
'              j1=iinp(0,'enter first output trib A ?   ')
'              if(inbet(j1 to 1 to ntmax) <= 0) return
'              j2=iinp(0,'enter last output trib B ?   ')
'              if(inbet(j2,1 to ntmax) <= 0 or j2 < j1) return

'              For i = j1 To j2
'                 Call Tcopy(Jtrib, i)
'                 Next i
'              NTrib = MAx(n, j2)

            End Select
             
'             Call clr(0)
'             write(*,*)
'     &       'NOW EDIT TRIBUTARY DATA TO CORRECT SEGMENT REFERENCES'
'             i = ihold(0)
'             Call clr(0)
'             Call edit(4)
'             GoTo 10
             End Sub

           Sub TribCopy(i, j)
'c copy trib data from trib i to trib j
            
            Darea(j) = Darea(i)
            TribName(j) = TribName(i)
            Iseg(j) = Iseg(i)
            Icode(j) = Icode(i)
            Flow(j) = Flow(i)
            CvFlow(j) = CvFlow(i)
            Ecoreg(j) = Ecoreg(i)

            For k = 1 To NCAT
                Warea(j, k) = Warea(i, k)
                Next k

            For k = 1 To 7
               Conci(j, k) = Conci(i, k)
               CvCi(j, k) = CvCi(i, k)
               Next k
    
            End Sub
        Sub TribZero(j)
'reset all data for trib j
            Darea(j) = 0
            TribName(j) = ""
            Iseg(j) = 1
            Icode(j) = 1
            Flow(j) = 0
            CvFlow(j) = 0
            Ecoreg(j) = 0

            For k = 1 To NCAT
                Warea(j, k) = 0
                Next k

            For k = 1 To 7
               Conci(j, k) = 0
               CvCi(j, k) = 0
               Next k
        End Sub
        Sub PipeZero(k)
'reset all data for pipe j
        
            PipeName(k) = ""
            Ito(k) = 0
            Ifr(k) = 0
            Qpipe(k) = 0
            CvQpipe(k) = 0
            Epipe(k) = 0
            CvEpipe(k) = 0
        
        End Sub
        
        Sub PipeCopy(i, j)
'copy data from pipe i to pipe j
        
            PipeName(j) = PipeName(i)
            Ito(j) = Ito(i)
            Ifr(j) = Ifr(i)
            Qpipe(j) = Qpipe(i)
            CvQpipe(j) = CvQpipe(i)
            Epipe(j) = Epipe(i)
            CvEpipe(j) = CvEpipe(i)
        
        End Sub
       
Sub AllZero()
'zero all input values

'reset output variables
    Call Ycopy_In
    For i = 1 To Nye
        Cye(i) = 0
        Next i
    Call Ycopy_Out

'reset all input variables
    For i = 0 To NTMAX
        Call TribZero(i)
        Next i
    For i = 0 To NSMAX
        Call SegZero(i)
        Next i
    For i = 0 To NPMAX
        Call PipeZero(i)
        Next i
        
    For i = 0 To NOPptions
        Iop(i) = IopDefault(i)
        Next i
      
    For i = 0 To NXk
        Xk(i) = XkDefault(i)
        CvXk(i) = CvXkDefault(i)
        Next i
    
    For i = 1 To NCAT
        Ur(i) = 0
        CvUr(i) = 0
        For j = 1 To NVariables
           Uc(i, j) = 0
           CvUc(i, j) = 0
           Next j
        Next i
    
    Nseg = 1
    NTrib = 1
    Npipe = 0
    
    SegName(1) = "Segment 1"
    TribName(1) = "Trib 1"
    
    For i = 1 To NGlobals
        P(i) = 0
        Cp(i) = 0
        Next i
    P(1) = 1   'period length
        
    For i = 1 To 10
        Note(i) = ""
        Next i
    
    Title = ""
    ErrTxt = ""
    Nmsg = 0
    Icalc = 0
    
End Sub
