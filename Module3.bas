Attribute VB_Name = "Module3"

Function Min(Dades)

Min = 1E+99
For x = 1 To UBound(Dades) - 1
    If Dades(x) < Min Then Min = Dades(x)
Next x
End Function

Function Tanh(Q)
    Tanh = (exp(Q) - exp(-Q)) / (exp(Q) + exp(-Q))
End Function
Function pacient_sa() As Variant

        ' Basal Parameters
        Const BW As Double = 78 ' Kg
        Const ib As Single = 25.49 ' pmol/l
        Const EGPb As Single = 1.92 ' mg/kg/min
        Const gb As Single = 91.76 ' mg/dl
        ' Meals
        Const Tdose1 As Double = (24 + 6) * 60
        Const Tdose2 As Double = (24 + 12) * 60
        Const Tdose3 As Double = (24 + 20) * 60
        Const tsim As Double = (24 * 2) * 60 ' simulation time
        Const tMig = 1
        
        Const dt As Single = 0.003 '0.01
        
        
        'Const Tdose4 As double = (Tdose3 + tsim) / 2
        Dim TDose(5) As Double
        TDose(0) = 0
        TDose(1) = Tdose1
        TDose(2) = Tdose2
        TDose(3) = Tdose3
        TDose(4) = tsim
        
        Const Dose1 As Double = 45000
        Const Dose2 As Double = 0 '70000
        Const Dose3 As Double = 0 '70000
        Dim Dose(5) As Double
        Dose(0) = 0
        Dose(1) = Dose1
        Dose(2) = Dose2
        Dose(3) = Dose3
        Dose(4) = 0
        
        '  Plasma meal rate of appearance parameters
        Const kabs As Single = 0.0568 'min^-1
        Const kmax As Single = 0.0558 'min^-1
        Const kmin As Single = 0.008 'kmin min^-1
        Const f As Single = 0.9 'dimensionless
        Const b As Single = 0.82 'dimensionless
        Const d As Single = 0.01 'dimensionless

        Const Vg As Single = 1.88 'Glucose Volume in dl/Kg
        Const Vi As Single = 0.05 'Insulin Volume in L/kg

        Const Gpb As Single = gb * Vg 'amount of glucose in the plasma compartment in the ss
        Const ipb As Single = ib * Vi 'pmol/kg amount of plasma insulin
        ' Excretion
        Const ke1 As Single = 0.0005 'min^-1
        Const ke2 As Double = 339 'mg/kg
        'Utilization Parameters
        Const Vmx As Single = 0.047 'mg/kg/min per pmol/L
        Const Vmx_normal As Single = 0.047
        Const km0 As Single = 225.59 'mg/kg
        Const K2 As Single = 0.079 'min^-1
        Const K1 As Single = 0.065 'min^-1
        Const Fsnc As Single = 1 'mg/kg/min

        Const Gtb As Single = (Fsnc - EGPb + K1 * Gpb) / K2 'mg/kg
        Const Vm0 As Single = (EGPb - Fsnc) * (km0 + Gtb) / Gtb 'mg/kg/min
        Const Rdb As Single = EGPb 'mg/kg/min
        Const PCRb As Single = Rdb / gb 'dl/kg/min
        'Insulin Action Parameter
        Const p2U = 0.0331 'min^-1
        ' Plasma Insulin & Insulin Secretion
        Const h As Single = gb 'mg/dl equals basal glucose concnetration
        Const K As Single = 2.28 'pmol/kg/(mg/dl)
        Const beta As Single = 0.11 'pmol/kg/min/(mg/dl)
        Const K_normal As Single = 2.28 'pmol/kg/(mg/dl)
        Const beta_normal As Single = 0.11 'pmol/kg/min/(mg/dl)
        Const alpha As Single = 0.05 'min^-1
        Const gamma As Single = 0.5 'min^-1
        Const m1 As Single = 0.19 'min^-1
        Const m2 As Single = 0.484 'min^-1
        Const m5 As Single = 0.0304 'min*kg/pmol
        Const HEb As Single = 0.6 'dimensionless
        Const m4 As Single = 2 / 3 * m2 * HEb 'min^-1
        Const ilb As Single = ipb * (m4 + m2) / m1 'pmol/kg
        Const m30 As Single = HEb * m1 / (1 - HEb)
        Const SRb As Single = ipb * m4 + ilb * m30 'pmol/kg/min
        Const ipo As Single = SRb / gamma 'pmol/kg
        Dim m6 As Single
        m6 = m5 * SRb + HEb
        ' Production Parameters
        Const ki As Single = 0.0079 'min^1
        Const kp2 As Single = 0.0021 'min^-1
        Const kp3 As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp3_normal As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp4 As Single = 0.0618 'mg/kg/min/(pmol/kg)
        Const kp1 As Single = EGPb + kp2 * Gpb + kp3 * ib + kp4 * ipo 'mg/kg/min

        'Attenzione:
        'ys(0)=stomaco1 (q11 in SAAM) Qsto1
        'ys(1)=stomaco2 (q12  in SAAM) Qsto2
        'ys(2)=intestino (q13 in SAAM) Qgut
        'ys(3)=glucosio plasmatico+organi insulino-indipendenti (q1 in
        'SAAM) Gp
        'ys(4)=glucosio+organi insulino dipendenti (q2 in SAAM) Gt
        'ys(5)=insulina labile in pmol/kg (q6 in SAAM) Y
        'ys(6)=insulina in Vena Porta in pmol/kg (q15 in SAAM) Ipo
        'ys(7)=insulina nel fegato in pmol/kg (q14 in SAAM) Il
        'ys(8)=insulina plasmatica in pmol/kg (q7 in SAAM) Ip
        'ys(9)=azione insulinica in pmol/L (q4 in SAAM) X
        'ys(10)=insulina ritardata 1 in pmol/L (q9 in SAAM) I1
        'ys(11)=insulina ritardata 2 in pmol/L(q10 in SAAM) Idel

        Dim glupt(tsim / tMig), gluit(tsim / tMig), inspt(tsim / tMig), insrt(tsim / tMig), insportt(tsim / tMig), actiont(tsim / tMig) As Double
        Dim escrez(tsim / dt), prodt(tsim / tMig), utilt(tsim / tMig), asst(tsim / tMig), secrt(tsim / tMig), escrt(tsim / tMig) As Double
        Dim t0 As Double: t0 = 0
        Dim conta As Double, y0(12) As Double, Tspan() As Double, w As Double
        Dim ys(tsim / dt, 12) As Double, last As Double
        last = 0
        Dim dery(12) As Double, ind As Double
        Dim dosekempt As Double
        dosekempt = 90000 'mg
        
        Dim escr, aa, cc, qsto, kgut, m3, HE As Double
        For ind = 0 To (tsim / dt) - 1
            For ind2 = 0 To 11
                ys(ind, ind2) = 0   'Inicializa derivada en cada instatne de tiempo de simulacion y en cada variable
            Next
        Next
        For j = 0 To 3 Step 1        'Recorre las 4 dosis: la incial es 0 y 3 ingestas
            w = 0
            Erase Tspan
            'ReDim Tspan((Tdose(j + 1) - Tdose(j)) / dt - 1)
            For num = TDose(j) To TDose(j + 1) Step dt
                ReDim Preserve Tspan(w)
                Tspan(w) = num                  'Almacena los instantes de tiempo de la simulacion inter-dosis
                w = w + 1
            Next
            If j = 0 Then
                'Initialize variables
                ys(0, 0) = Dose(0)
                ys(0, 1) = 0
                ys(0, 2) = 0
                ys(0, 3) = Gpb
                ys(0, 4) = Gtb
                ys(0, 5) = 0
                ys(0, 6) = ipo
                ys(0, 7) = ilb
                ys(0, 8) = ipb
                ys(0, 9) = 0
                ys(0, 10) = ib
                ys(0, 11) = ib
            Else
                ys(last, 0) = ys(last, 0) + Dose(j)
            End If
            If Dose(j) > 0 Then
                dosekempt = ys(last, 0) + ys(last, 1)
            End If
            For ind = 0 To UBound(Tspan) - 1 '.Length '- 1     'Recorrre todos los instantes de tiempo del periodo inter-dosis
                ' Renal Excretion:
                If ys(last, 3) > ke2 Then
                    escr = ke1 * (ys(last, 3) - ke2)
                Else
                    escr = 0
                End If
                ' Plasma Glucose
                dery(3) = Max(kp1 - kp2 * ys(last, 3) - kp3 * ys(last, 11) - kp4 * ys(last, 6), 0) + Max(kabs * ys(last, 2) * f / BW, 0) - Fsnc - K1 * ys(last, 3) + K2 * ys(last, 4) - escr
                dery(4) = K1 * ys(last, 3) - K2 * ys(last, 4) - (Vm0 + Vmx * ys(last, 9)) / (km0 + ys(last, 4)) * ys(last, 4)
                ' Absorption
                dery(0) = -kmax * ys(last, 0)
                qsto = ys(last, 0) + ys(last, 1)
                If (Dose(1) = 0 And Dose(1) = 0 And Dose(2) = 0 And Dose(3) = 0 And Dose(4) = 0 And Dose(5) = 0) Or j = 0 Then
                    dery(1) = 0
                    dery(2) = 0
                Else
                    aa = 5 / 2 / (1 - b) / dosekempt
                    cc = 5 / 2 / d / dosekempt
                    kgut = kmin + (kmax - kmin) / 2 * (Tanh(aa * (qsto - b * dosekempt)) - Tanh(cc * (qsto - d * dosekempt)) + 2)
                    dery(1) = kmax * ys(last, 0) - ys(last, 1) * kgut
                    dery(2) = kgut * ys(last, 1) - kabs * ys(last, 2)
                End If
                ' Plasma insulin and secretion
                If (beta * (ys(last, 3) / Vg - h)) > -SRb Then
                    dery(5) = -alpha * ys(last, 5) + alpha * beta * (ys(last, 3) / Vg - h)
                Else
                    dery(5) = -alpha * ys(last, 5) - alpha * SRb
                End If
                If (dery(3) > 0) And (ys(last, 3) / Vg > h) Then
                    dery(6) = ys(last, 5) - gamma * ys(last, 6) + K * dery(3) / Vg + SRb
                Else
                    dery(6) = ys(last, 5) - gamma * ys(last, 6) + SRb
                End If
                m6 = HEb + m5 * SRb
                HE = -m5 * gamma * ys(last, 6) + m6
                m3 = m1 * HE / (1 - HE)
                dery(7) = gamma * ys(last, 6) - (m1 + m3) * ys(last, 7) + m2 * ys(last, 8)
                dery(8) = m1 * ys(last, 7) - (m2 + m4) * ys(last, 8)
                ' Insulin action
                dery(9) = -p2U * ys(last, 9) + p2U * (ys(last, 8) / Vi - ib)
                ' Production (delayed insulin)
                dery(10) = ki * (ys(last, 8) / Vi - ys(last, 10))
                dery(11) = ki * (ys(last, 10) - ys(last, 11))
                For ind2 = 0 To 11 '******************EULER*********************
                    ys(last + 1, ind2) = ys(last, ind2) + dt * dery(ind2)
                Next
                last = last + 1
                DoEvents
            Next
        Next
        ' Plotting variables, average over 10 Minutes
        For ind = 0 To (tsim / tMig) - 1
            glupt(ind) = ys(ind * tMig / dt, 3) * dt / tMig / Vg
            gluit(ind) = ys(ind * tMig / dt, 4) * dt / tMig
            inspt(ind) = ys(ind * tMig / dt, 8) * dt / tMig / Vi
            insrt(ind) = ys(ind * tMig / dt, 1) * dt / tMig
            insportt(ind) = ys(ind * tMig / dt, 6) * dt / tMig
            actiont(ind) = ys(ind * tMig / dt, 9) * dt / tMig
            ' assorbimento in mg/(Kg*min)
            asst(ind) = kabs * ys(ind * tMig / dt, 2) * f / BW * dt / tMig
            For ind2 = 1 To tMig / dt - 1
                ' glup=plasmatic glucose in mg/dl
                glupt(ind) = glupt(ind) + (ys(ind * tMig / dt + ind2, 3)) * dt / tMig / Vg
                ' organi insulino dipendenti in mg/kg
                gluit(ind) = gluit(ind) + (ys(ind * tMig / dt + ind2, 4)) * dt / tMig
                ' insp=plasmatic insulin and insr=remote insulin in pmol/L
                inspt(ind) = inspt(ind) + (ys(ind * tMig / dt + ind2, 8)) * dt / tMig / Vi
                insrt(ind) = insrt(ind) + (ys(ind * tMig / dt + ind2, 1)) * dt / tMig
                insportt(ind) = insportt(ind) + (ys(ind * tMig / dt + ind2, 6)) * dt / tMig
                ' action insulin in pmol/L
                actiont(ind) = actiont(ind) + ys(ind * tMig / dt + ind2, 9) * dt / tMig
                ' assorbimento in mg/(Kg*min)
                asst(ind) = asst(ind) + kabs * ys(ind * tMig / dt + ind2, 2) * f / BW * dt / tMig
            Next
            
            '***************************************************
            '  SALIDA POR PANTALLA
            '***************************************************
'            If ind Mod 2 = 0 Then
'                cv = cv + 1
'                Chart1.Column = 1
'                Chart1.RowCount = cv
'                Chart1.Row = cv
 '               Chart1.RowLabel = ind * 10
 '               Chart1.Data = glupt(ind)
                
 '               Chart1.Column = 2
  '              Chart1.Data = 0 ' inspt(ind)
  '          End If
                    
            'Chart1.Series("Series1").Points.AddXY(ind * 10, glupt(ind))
            'Debug.Print (glupt(ind))
            ' prod in mg/(Kg*min)
            prodt(ind) = Max(kp1 - kp2 * glupt(ind) * Vg - kp3 * insrt(ind) - kp4 * insportt(ind), 0)
            ' util in mg/(Kg*min)
            utilt(ind) = Fsnc + (Vm0 + Vmx * actiont(ind)) / (km0 + gluit(ind)) * gluit(ind)
            ' secr in pmol/(kg*min)
            secrt(ind) = gamma * insportt(ind)
            'Debug.Print(gluit(ind))
        Next
        'Chart1.Titles.Add ("Healthy person [Glucose] plasma")
        ' escrezione renale
        For ind = 0 To last - 1
            If ys(ind, 3) > ke2 Then
                escrez(ind) = ke1 * (ys(ind, 3) * Vg - ke2)
            Else
                escrez(ind) = 0
            End If
            
        Next
        
        'Return glupt
        
        mm = FreeFile
        Open "\\VBOXSVR\Compartit_Windows\diab\pacient_sa.txt" For Output As mm
        'For Index = 1 To last - 1 Step 100
        For Index = 0 To UBound(glupt) Step 1
            Print #mm, Index, glupt(Index), inspt(Index)
          Next Index
        Close #mm
        'return glupt
  pacient_sa = glupt
  
        
End Function
Function diabetic_random(ab, cd, ef, gh, poblacio, cel, contador) As Variant
'diabetic_random(dos, tres, quatre, cinc, poblacio1, linia, poblacio1)
'Private Sub Command1_Click()
       '***************************
        '   DM1
        '***************************
        '----- Implant ----
        Dim K1_mA As Double, K1_pI As Double, K1_ins_in As Double, K1_ins_ex As Double, K1_N As Double
        Dim New_mA As Double, New_pI As Double, New_ins_in As Double, New_ins_ex As Double, New_N As Double
        Dim Kglu As Double
        
        Const gamma As Single = 1 * 10 ^ -8 / 0.1
        Const w1 As Single = 1 * 10 ^ -6.5
        Const zz As Double = 3
        Const timescale As Single = 0.1
        Const d_insi As Single = 0 ';0.03/timescale'; % degradacio ins interna
        Const d_inse As Single = 0 '0.004 / timescale '; % degradacio ins externa
        Const r As Single = 0.1 / timescale '
        Const Vcel As Single = 1 / timescale '; % volum celula
        Const Vext As Single = 100 / timescale '; % volum medi
        Const kk As Single = 10 '100 / timescale '; % volum medi
        Const tMig = 1
        
        Dim G1 As Double '
        G1 = 10
        Dim d_mA As Single '
        d_mA = ab ' 0.009 / timescale   '; % missatger pro ins
        Dim KpI As Single
        KpI = cd ' 0.1 / timescale '; % mA->pI
        Dim Kins As Single
        Kins = ef '1 / timescale '; % pins->in
        Dim Kex As Single
        Kex = gh '0.1 / timescale   ';% % secrecio ins a fora
         

        Dim tamany As Single
        tamany = poblacio '5000 '0 '5000'0
        Dim mA As Double '= 0
        Dim pI As Double '= 0
        Dim ins_in As Double '= 0
        Dim ins_ex As Double '= 0
        Dim N As Double '= 0.1
        N = 0.001
        Dim rate_ins As String  ' 0


        Const status As Double = 1
        Const openclosed As Double = 0
        Const Tdose1 As Double = (24 + 6) * 60 'min
        Const Tdose2 As Double = (24 + 12) * 60 'min
        Const Tdose3 As Double = (24 + 20) * 60  'min
        Const irrb As Double = 0

        Const Dose1 As Double = 75000  'mg
        Const Dose2 As Double = 0 '70000 'mg
        Const Dose3 As Double = 0 '70000 'mg

        Const BW As Double = 78 'kg
        Const ib As Single = 0 '41.32 'pmol/l
        Const EGPb As Single = 2.4 'mg/kg/min
        Const gb As Double = 180 'mg/dl
        ' tempi
        Const dt As Single = 0.0015 ' time step min
        Const IncT As Double = dt
        Const tsim As Double = (24 * 2) * 60 'min ¿? son 150 hores
        'Const Tdose4=(Tdose3+tsim)/2 'min
        
        Dim TDose(4) As Double
        TDose(0) = 0
        TDose(1) = Tdose1
        TDose(2) = Tdose2
        TDose(3) = Tdose3
        TDose(4) = tsim

        ' Dosi
        Dim Dose(5) As Double
        Dose(0) = 0
        Dose(1) = Dose1
        Dose(2) = Dose2
        Dose(3) = Dose3
        Dose(4) = 0

        'Dosi insulina
        Const Ipumpb As Double = 1
        Const dose_ins1 As Double = 3
        Const dose_ins2 As Double = 5
        Const dose_ins3 As Double = 5
        Dim dose_i(4) As Double
        dose_i(0) = 0
        dose_i(1) = 0 ' dose_ins1 * 6000
        dose_i(2) = 0 ' dose_ins2 * 6000
        dose_i(3) = 0 ' dose_ins3 * 6000
        dose_i(4) = 0
        
        '  Plasma meal rate of appearance parameters
        Const kabs As Single = 0.0568 'min^-1
        Const kmax As Single = 0.0558 'min^-1
        Const kmin As Single = 0.008 'kmin min^-1
        Const f As Single = 0.9 'dimensionless
        Const b As Single = 0.82 'dimensionless
        Const d As Single = 0.01 'dimensionless

        Const Vg As Single = 1.88 'Glucose Volume in dl/Kg
        Const Vi As Single = 0.05 'Insulin Volume in L/kg

        Const Gpb As Single = gb * Vg 'amount of glucose in the plasma compartment in the ss
        Const ipb As Single = 3 'ib * Vi 'pmol/kg amount of plasma insulin
        ' Excretion
        Const ke1 As Single = 0.0005 'min^-1
        Const ke2 As Double = 339 'mg/kg
        'Utilization Parameters
        Const Vmx As Single = 0.047 'mg/kg/min per pmol/L
        Const Vmx_normal As Single = 0.047
        Const km0 As Single = 225.59 'mg/kg
        Const K2 As Single = 0.079 'min^-1
        Const K1 As Single = 0.065 'min^-1
        Const Fsnc As Single = 1 'mg/kg/min

        Const Gtb As Single = (Fsnc - EGPb + K1 * Gpb) / K2 'mg/kg
        Const Vm0 As Single = (EGPb - Fsnc) * (km0 + Gtb) / Gtb 'mg/kg/min
        Const Rdb As Single = EGPb 'mg/kg/min
        Const PCRb As Single = Rdb / gb 'dl/kg/min
        'Insulin Action Parameter
        Const p2U = 0.0331 'min^-1
        ' Plasma Insulin & Insulin Secretion
        Const m1 As Single = 0.19 'min^-1
        Const m2 As Single = 0.484 'min^-1
        Const m5 As Single = 0.0304 'min*kg/pmol
        Const HEb As Single = 0.6 'dimensionless
        Const m4 As Single = 2 / 3 * m2 * HEb 'min^-1
        Const m3 As Double = HEb * m1 / (1 - HEb)
        Const ilb As Double = ipb * m2 / (m1 + m3) 'pmol/kg
        Const SRb As Double = ipb * m4 + ilb * m3 'pmol/kg/min
        ' Production Parameters
        Const ki As Single = 0.0079 'min^1
        Const kp2 As Single = 0.0021 'min^-1
        Const kp3 As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp3_normal As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp1 As Single = EGPb + kp2 * Gpb + kp3 * ib 'mg/kg/min
        ' PID
        Const Kp As Single = 2.5 'pmol/min per mg/dl (da Steil 2006 Kp=0.025 U/h per mg/dl)
        Const TI As Double = 450 'min da Steil uomo 2006
        Const TD As Double = 66 'min da Steil uomo 2006
        ' insulin infusion
        Const ksc As Single = 0.01 '0.03; %min^-1 =1/tau2 di Steil 2006: tau2=33.5 min in media -> k=0.03 min^-1
        Const kd As Single = 0.0164
        Const ka1 As Single = 0.0018
        Const ka2 As Single = 0.0182
        Const isc1_ss As Double = irrb / (kd + ka1)
        Const isc2_ss As Double = (kd / ka2) * isc1_ss
        
        ' Cell implant
        Dim gamma1 As Double: gamma1 = 0.0033
        Dim alpha1 As Double: alpha1 = 12.2
        Dim gamma2 As Double: gamma2 = 0.00054
        Dim n1 As Double: n1 = 2
        Dim n2 As Double: n2 = 1.5
        Dim omega1 As Double: omega1 = 10
        Dim omega2 As Double: omega2 = 50
        Dim kappa1 As Double: kappa1 = 0.00385
        Dim kappa2 As Double: kappa2 = 0.0002
        Const FacS As Double = 15

        ' Attenzione:
        '  ys(0)=stomaco1 (q11 in SAAM)
        '  ys(1)=stomaco2 (q12  in SAAM)
        '  ys(2)=intestino (q13 in SAAM)
        '  ys(3)=glucosio plasmatico+organi insulino-indipendenti (q1 in SAAM)
        '  ys(4)=glucosio+organi insulino dipendenti (q2 in SAAM)

        'NUOVO       ys(5)=glucosio nel sottocute in mg/kg
        'NUOVO       ys(6)=insulina NON-MONOMERICA nel sottocute Isc1 in pmol/kg

        '  ys(7)=insulina nel fegato in pmol/kg (q14 in SAAM)
        '  ys(8)=insulina plasmatica in pmol/kg (q7 in SAAM)
        '  ys(9)=azione insulinica in pmol/L (q4 in SAAM)
        '  ys(10)=insulina ritardata 1 in pmol/L (q9 in SAAM)
        '  ys(11)=insulina ritardata 2 in pmol/L(q10 in SAAM)

        ' NUOVISSIMO  ys(12)=insulina MONOMERICA nel sottocute Isc2 in pmol/kg

        ' RENAL EXCRETION;
        
        Temps = 0

        Dim glupt(tsim / tMig), gluit(tsim / tMig), inspt(tsim / tMig), insrt(tsim / tMig), actiont(tsim / tMig) As Double
        Dim escrez(tsim / dt), prodt(tsim / tMig), utilt(tsim / tMig), asst(tsim / tMig), secrt(tsim / tMig), escrt(tsim / tMig), absit(tsim / tMig) As Double
        Dim insulina(tsim / dt) As Double '(tsim / tmig)
        Dim t0 As Double: t0 = 0
        Dim Tspan() As Double, w As Double
        
        Dim ys(tsim / dt, 16) As Double, last As Double: last = 0
        Dim dery(16) As Double, ind As Double
        Dim dosekempt As Double: dosekempt = 90000 'mg
        Dim escr, aa, cc, qsto, kgut, HE As Double
        For j = 0 To 3 Step 1
            w = 0
            Erase Tspan
            'ReDim Tspan((Tdose(j + 1) - Tdose(j)) / dt - 1)
            For num = TDose(j) To TDose(j + 1) Step dt
                ReDim Preserve Tspan(w)
                Tspan(w) = num
                w = w + 1
            Next
            If j = 0 Then

                ' Initialize state variables
                ys(0, 0) = Dose(0)
                ys(0, 1) = 0
                ys(0, 2) = 0
                ys(0, 3) = Gpb
                ys(0, 4) = Gtb
                ys(0, 5) = gb 'Gpb
                ys(0, 6) = isc1_ss
                ys(0, 7) = ilb
                ys(0, 8) = ipb
                ys(0, 9) = 0
                ys(0, 10) = ib
                ys(0, 11) = ib
                ys(0, 12) = isc2_ss
                ys(0, 13) = 0
                ys(0, 14) = 0
                ys(0, 15) = 0
                ys(0, 16) = 0
            Else
                ys(last, 0) = ys(last, 0) + Dose(j) 'diabetico j>1
                ys(last, 6) = ys(last, 6) + dose_i(j) / BW
            End If
            If Dose(j) > 0 Then
                dosekempt = ys(last, 0) + ys(last, 1)
            End If
            For ind = 0 To UBound(Tspan) - 1 '.Length - 1
                ' RENAL EXCRETION
                If ys(last, 3) > ke2 Then
                    escr = ke1 * (ys(last, 3) - ke2)
                Else
                    escr = 0
                End If
                ' PLASMA GLUCOSE
                dery(3) = Max((kp1 - kp2 * ys(last, 3) - kp3 * ys(last, 11)), 0) + Max(kabs * ys(last, 2) * f / BW, 0) - Fsnc - K1 * ys(last, 3) + K2 * ys(last, 4) - escr
                dery(4) = K1 * ys(last, 3) - K2 * ys(last, 4) - (Vm0 + Vmx * ys(last, 9)) / (km0 + ys(last, 4)) * ys(last, 4)
                ' ASSORBIMENTO
                dery(0) = -kmax * ys(last, 0)
                qsto = ys(last, 0) + ys(last, 1)
                If (Dose(0) = 0 And Dose(1) = 0 And Dose(2) = 0 And Dose(3) = 0 And Dose(4) = 0 And Dose(5) = 0) Or j = 0 Then
                    dery(1) = 0
                    dery(2) = 0
                Else
                    aa = 5 / 2 / (1 - b) / dosekempt
                    cc = 5 / 2 / d / dosekempt
                    kgut = kmin + (kmax - kmin) / 2 * (Tanh(aa * (qsto - b * dosekempt)) - Tanh(cc * (qsto - d * dosekempt)) + 2)
                    dery(1) = kmax * ys(last, 0) - ys(last, 1) * kgut
                    dery(2) = kgut * ys(last, 1) - kabs * ys(last, 2)
                End If
                ' Ginterstizio (ritardo=10 min rispetto al plasma)
                dery(5) = -0.1 * ys(last, 5) + 0.1 * (ys(last, 3) / Vg)
                'INSULIN) in SC pmol/Kg

                '********************************************
                '* El modelo trabaja en mg/Kg y el codigo del
                '* device lo hace en mM. La conversión es
                '* 1mM=0.018 gr/dL * Vg dL/Kg
                '********************************************
'omega1 = 20
'omega1 = 13 '15
                dery(13) = gamma1 * (0.018 * Vg * ys(last, 5)) ^ n1 / (1 + alpha1 * (0.018 * Vg * ys(last, 5) / omega1) ^ n1) - gamma2 * ys(last, 13) ^ n2 / (1 + (ys(last, 13) / omega2) ^ n2) ' glucosa sottocute
                dery(14) = gamma2 * ys(last, 13) ^ n1 / (1 + (ys(last, 13) / omega2) ^ n2) - kappa1 * ys(last, 14)
                'dery(15) = kappa1 * ys(last, 14) - kappa2 * ys(last, 15) ' consumption ADDED
                'dery(6) = kappa2 * ys(last, 15) - (kd + ka1) * ys(last, 6) 'insulin from implant comes here

'INTRODUIR UNA FUNCION DE L IMPLANT input=ys 5 ouput=Ipumpb
                '  ---
                ' eva = ys(last, 5)
                   Kglu = (gamma * (ys(last, 5) / 2) ^ zz) / (1 + w1 * (ys(last, 5) / 2) ^ zz)
                    '--- Eq's ---
                    K1_mA = (Kglu * G1 - d_mA * mA) * IncT
                    K1_pI = (KpI * mA - Kins * pI) * IncT
                    K1_ins_in = (Kins * pI - Kex * ins_in) * IncT
                    K1_ins_ex = (Kex * N * ins_in * (Vcel / Vext)) * IncT
                    K1_N = (r * N * (1 - N / kk)) * IncT
                    
                    New_mA = mA + K1_mA
                    New_pI = pI + K1_pI
                    New_ins_in = ins_in + K1_ins_in
                    New_ins_ex = K1_ins_ex
                    New_N = N + K1_N
                            
                   
                    ' ----
                'Call implant(ys(last, 5), mA, pI, ins_in, ins_ex, N)
                 rate_ins = tamany * ins_ex
                insulina(last) = rate_ins
                dery(6) = -(kd + ka1) * ys(last, 6) + rate_ins   'Version original
'dery(6) = FacS * kappa1 * ys(last, 14) - (kd + ka1) * ys(last, 6)  'Cambio mio
'dery(15) = FacS * kappa1 * ys(last, 14)
                dery(12) = kd * ys(last, 6) - ka2 * ys(last, 12)

                ' PLASMA(INSULIN)
                dery(7) = -(m1 + m3) * ys(last, 7) + m2 * ys(last, 8)
                dery(8) = m1 * ys(last, 7) - (m2 + m4) * ys(last, 8) + ka1 * ys(last, 6) + ka2 * ys(last, 12)
                ' INSULIN ACTION in pmol/L
                dery(9) = -p2U * ys(last, 9) + p2U * (ys(last, 8) / Vi - ib)
                ' PRODUCTION Delayed Insulin for PRODUCTION in pmol/L
                dery(10) = ki * (ys(last, 8) / Vi - ys(last, 10))
                dery(11) = ki * (ys(last, 10) - ys(last, 11))
                ' EULER
                For ind2 = 0 To 15
                    ys(last + 1, ind2) = ys(last, ind2) + dt * dery(ind2)
                Next
                    
                    mA = New_mA
                    pI = New_pI
                    ins_in = New_ins_in
                    ins_ex = New_ins_ex
                    N = New_N
    
                
'cv = cv + 1
'Chart1.Column = 1
'Chart1.RowCount = cv
'Chart1.Row = cv
'Chart1.RowLabel = ind * 10
'Chart1.Data = ys(last, 6)
                
                last = last + 1
                DoEvents
            Next
            DoEvents

        Next

        ' Plotting variables, average over 10 Minutes
        For ind = 0 To (tsim / tMig) - 1
            DoEvents
            glupt(ind) = ys(ind * tMig / dt, 3) * dt / tMig / Vg
            gluit(ind) = ys(ind * tMig / dt, 4) * dt / tMig
            inspt(ind) = ys(ind * tMig / dt, 8) * dt / tMig / Vi
            insrt(ind) = ys(ind * tMig / dt, 11) * dt / tMig
            actiont(ind) = ys(ind * tMig / dt, 9) * dt / tMig
            absit(ind) = (ka1 * ys(ind * tMig / dt, 6) + ka2 * ys(ind * tMig / dt, 12)) * dt / tMig
            ' assorbimento in mg/(Kg*min)
            asst(ind) = kabs * ys(ind * tMig / dt, 2) * f / BW * dt / tMig
            For ind2 = 1 To tMig / dt - 1
                ' glup=plasmatic glucose in mg/dl
                glupt(ind) = glupt(ind) + (ys(ind * tMig / dt + ind2, 3)) * dt / tMig / Vg
                ' organi insulino dipendenti in mg/kg
                gluit(ind) = gluit(ind) + (ys(ind * tMig / dt + ind2, 4)) * dt / tMig
                ' insp=plasmatic insulin and insr=remote insulin in pmol/L
                inspt(ind) = inspt(ind) + (ys(ind * tMig / dt + ind2, 8)) * dt / tMig / Vi
                insrt(ind) = insrt(ind) + (ys(ind * tMig / dt + ind2, 11)) * dt / tMig
                actiont(ind) = actiont(ind) + ys(ind * tMig / dt + ind2, 9) * dt / tMig
                ' assorbimento in mg/(Kg*min)
                asst(ind) = asst(ind) + kabs * ys(ind * tMig / dt + ind2, 2) * f / BW * dt / tMig
                ' assorbimento s.c. di insulina in pmol/(kg*min)
                absit(ind) = absit(ind) + (ka1 * ys(ind * tMig / dt + ind2, 6) + ka2 * ys(ind * tMig / dt + ind2, 12)) * dt / tMig
            Next

            '***************************************************
            '  SALIDA POR PANTALLA
            '***************************************************
            'If ind Mod 2 = 0 Then
             '   cv = cv + 1
              '  Chart1.Column = 1
               ' Chart1.RowCount = cv
                'Chart1.Row = cv
            '    Chart1.RowLabel = ind * tmig
             '   Chart1.Data = glupt(ind)
                
              '  Chart1.Column = 2
               ' Chart1.Data = inspt(ind)
            'End If
            
            
            'Chart1.Series("Series1").Points.AddXY(ind * tmig, glupt(ind))
            'Debug.Print (glupt(ind))
            ' prod in mg/(Kg*min)
            prodt(ind) = Max(kp1 - kp2 * glupt(ind) * Vg - kp3 * insrt(ind), 0)                       ' util in mg/(Kg*min)
            utilt(ind) = Fsnc + (Vm0 + Vmx * actiont(ind)) / (km0 + gluit(ind)) * gluit(ind)
        Next
        'Chart1.Titles.Add ("D. Type I open loop [Glucose] plasma")
        ' escrezione renale
        For ind = 0 To last - 1
            If ys(ind, 3) > ke2 Then
                escrez(ind) = ke1 * (ys(ind, 3) * Vg - ke2)
            Else
                escrez(ind) = 0
            End If
        Next
        
        diabetic_random = glupt
        
        '----- guardar en document glupt(ind)
'        mm = FreeFile
 '       Open "E:\glucose_" & contador & ".txt" For Output As mm
 '       'For Index = 1 To last - 1 Step 100
 '       For Index = 0 To UBound(glupt) Step 1
 '           Print #mm, Index, glupt(Index), inspt(Index)
 '           'Print #mm, Index, ys(Index, 5), ys(Index, 6), ys(Index, 8), Kglu
 '         Next Index
 '       Close #mm
    
      '  mm = FreeFile
    '    Open "\\VBOXSVR\Compartit_Windows\diab\celula_" & cel & "glucose_SUB_" & contador & ".txt" For Output As mm
     '   'For Index = 1 To last - 1 Step 100
     '   For Index = 0 To last - 1 Step 1000 'UBound(glupt) Step 1
      '      'Print #mm, Index * dt, glupt(Index), inspt(Index), rate_ins
      '      Print #mm, Index * dt, ys(Index, 5), insulina(Index)
      '    Next Index
      '  Close #mm
        
        '----- guardar en document glupt(ind)
       ' mm = FreeFile
       ' Open "\\VBOXSVR\Compartit_Windows\diab\celula_" & cel & "glucose_" & contador & ".txt" For Output As mm
      '  'For Index = 1 To last - 1 Step 100
      '  For Index = 0 To UBound(glupt) Step 1
        '    Print #mm, Index, glupt(Index), inspt(Index), insulina(Index)
        '  Next Index
       ' Close #mm
    
        
End Function

Function Max(a, b)
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Sub diabetic(poblacio)
'Private Sub Command1_Click()
       '***************************
        '   DM1
        '***************************
        '----- Implant ----
        Dim K1_mA As Double, K1_pI As Double, K1_ins_in As Double, K1_ins_ex As Double, K1_N As Double
        Dim New_mA As Double, New_pI As Double, New_ins_in As Double, New_ins_ex As Double, New_N As Double
        Dim Kglu As Double
        
        Const gamma As Single = 1 * 10 ^ -8 / 0.1
        Const w1 As Single = 1 * 10 ^ -6.5
        Const zz As Double = 3
        Const timescale As Single = 0.1
        Const d_insi As Single = 0 ';0.03/timescale'; % degradacio ins interna
        Const d_inse As Single = 0 '0.004 / timescale '; % degradacio ins externa
        Const r As Single = 0.1 / timescale '
        Const Vcel As Single = 1 / timescale '; % volum celula
        Const Vext As Single = 100 / timescale '; % volum medi
        Const kk As Single = 10 '100 / timescale '; % volum medi
        Const tMig = 1
        
        Dim G1 As Double '
        G1 = 10
        Dim d_mA As Single '
        d_mA = 0.009 / timescale  '; % missatger pro ins
        Dim KpI As Single
        KpI = 0.1 / timescale  '; % mA->pI
        Dim Kins As Single
        Kins = 1 / timescale  '; % pins->in
        Dim Kex As Single
        Kex = 0.1 / timescale  ';% % secrecio ins a fora
         

        Dim tamany As Double
        tamany = poblacio '5000 '0 '5000'0
        Dim mA As Double '= 0
        Dim pI As Double '= 0
        Dim ins_in As Double '= 0
        Dim ins_ex As Double '= 0
        Dim N As Double '= 0.1
        N = 0.001
        Dim rate_ins As String  ' 0


        Const status As Double = 1
        Const openclosed As Double = 0
        Const Tdose1 As Double = (24 + 6) * 60 'min
        Const Tdose2 As Double = (24 + 12) * 60 'min
        Const Tdose3 As Double = (24 + 20) * 60  'min
        Const irrb As Double = 0

        Const Dose1 As Double = 75000  'mg
        Const Dose2 As Double = 0 '70000 'mg
        Const Dose3 As Double = 0 '70000 'mg

        Const BW As Double = 78 'kg
        Const ib As Single = 0 '41.32 'pmol/l
        Const EGPb As Single = 2.4 'mg/kg/min
        Const gb As Double = 180 'mg/dl
        ' tempi
        Const dt As Single = 0.0015 ' time step min
        Const IncT As Double = dt
        Const tsim As Double = (24 * 2) * 60 'min ¿? son 150 hores
        'Const Tdose4=(Tdose3+tsim)/2 'min
        
        Dim TDose(4) As Double
        TDose(0) = 0
        TDose(1) = Tdose1
        TDose(2) = Tdose2
        TDose(3) = Tdose3
        TDose(4) = tsim

        ' Dosi
        Dim Dose(5) As Double
        Dose(0) = 0
        Dose(1) = Dose1
        Dose(2) = Dose2
        Dose(3) = Dose3
        Dose(4) = 0

        'Dosi insulina
        Const Ipumpb As Double = 1
        Const dose_ins1 As Double = 3
        Const dose_ins2 As Double = 5
        Const dose_ins3 As Double = 5
        Dim dose_i(4) As Double
        dose_i(0) = 0
        dose_i(1) = 0 ' dose_ins1 * 6000
        dose_i(2) = 0 ' dose_ins2 * 6000
        dose_i(3) = 0 ' dose_ins3 * 6000
        dose_i(4) = 0
        
        '  Plasma meal rate of appearance parameters
        Const kabs As Single = 0.0568 'min^-1
        Const kmax As Single = 0.0558 'min^-1
        Const kmin As Single = 0.008 'kmin min^-1
        Const f As Single = 0.9 'dimensionless
        Const b As Single = 0.82 'dimensionless
        Const d As Single = 0.01 'dimensionless

        Const Vg As Single = 1.88 'Glucose Volume in dl/Kg
        Const Vi As Single = 0.05 'Insulin Volume in L/kg

        Const Gpb As Single = gb * Vg 'amount of glucose in the plasma compartment in the ss
        Const ipb As Single = 3 'ib * Vi 'pmol/kg amount of plasma insulin
        ' Excretion
        Const ke1 As Single = 0.0005 'min^-1
        Const ke2 As Double = 339 'mg/kg
        'Utilization Parameters
        Const Vmx As Single = 0.047 'mg/kg/min per pmol/L
        Const Vmx_normal As Single = 0.047
        Const km0 As Single = 225.59 'mg/kg
        Const K2 As Single = 0.079 'min^-1
        Const K1 As Single = 0.065 'min^-1
        Const Fsnc As Single = 1 'mg/kg/min

        Const Gtb As Single = (Fsnc - EGPb + K1 * Gpb) / K2 'mg/kg
        Const Vm0 As Single = (EGPb - Fsnc) * (km0 + Gtb) / Gtb 'mg/kg/min
        Const Rdb As Single = EGPb 'mg/kg/min
        Const PCRb As Single = Rdb / gb 'dl/kg/min
        'Insulin Action Parameter
        Const p2U = 0.0331 'min^-1
        ' Plasma Insulin & Insulin Secretion
        Const m1 As Single = 0.19 'min^-1
        Const m2 As Single = 0.484 'min^-1
        Const m5 As Single = 0.0304 'min*kg/pmol
        Const HEb As Single = 0.6 'dimensionless
        Const m4 As Single = 2 / 3 * m2 * HEb 'min^-1
        Const m3 As Double = HEb * m1 / (1 - HEb)
        Const ilb As Double = ipb * m2 / (m1 + m3) 'pmol/kg
        Const SRb As Double = ipb * m4 + ilb * m3 'pmol/kg/min
        ' Production Parameters
        Const ki As Single = 0.0079 'min^1
        Const kp2 As Single = 0.0021 'min^-1
        Const kp3 As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp3_normal As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp1 As Single = EGPb + kp2 * Gpb + kp3 * ib 'mg/kg/min
        ' PID
        Const Kp As Single = 2.5 'pmol/min per mg/dl (da Steil 2006 Kp=0.025 U/h per mg/dl)
        Const TI As Double = 450 'min da Steil uomo 2006
        Const TD As Double = 66 'min da Steil uomo 2006
        ' insulin infusion
        Const ksc As Single = 0.01 '0.03; %min^-1 =1/tau2 di Steil 2006: tau2=33.5 min in media -> k=0.03 min^-1
        Const kd As Single = 0.0164
        Const ka1 As Single = 0.0018
        Const ka2 As Single = 0.0182
        Const isc1_ss As Double = irrb / (kd + ka1)
        Const isc2_ss As Double = (kd / ka2) * isc1_ss
        
        ' Cell implant
        Dim gamma1 As Double: gamma1 = 0.0033
        Dim alpha1 As Double: alpha1 = 12.2
        Dim gamma2 As Double: gamma2 = 0.00054
        Dim n1 As Double: n1 = 2
        Dim n2 As Double: n2 = 1.5
        Dim omega1 As Double: omega1 = 10
        Dim omega2 As Double: omega2 = 50
        Dim kappa1 As Double: kappa1 = 0.00385
        Dim kappa2 As Double: kappa2 = 0.0002
        Const FacS As Double = 15

        ' Attenzione:
        '  ys(0)=stomaco1 (q11 in SAAM)
        '  ys(1)=stomaco2 (q12  in SAAM)
        '  ys(2)=intestino (q13 in SAAM)
        '  ys(3)=glucosio plasmatico+organi insulino-indipendenti (q1 in SAAM)
        '  ys(4)=glucosio+organi insulino dipendenti (q2 in SAAM)

        'NUOVO       ys(5)=glucosio nel sottocute in mg/kg
        'NUOVO       ys(6)=insulina NON-MONOMERICA nel sottocute Isc1 in pmol/kg

        '  ys(7)=insulina nel fegato in pmol/kg (q14 in SAAM)
        '  ys(8)=insulina plasmatica in pmol/kg (q7 in SAAM)
        '  ys(9)=azione insulinica in pmol/L (q4 in SAAM)
        '  ys(10)=insulina ritardata 1 in pmol/L (q9 in SAAM)
        '  ys(11)=insulina ritardata 2 in pmol/L(q10 in SAAM)

        ' NUOVISSIMO  ys(12)=insulina MONOMERICA nel sottocute Isc2 in pmol/kg

        ' RENAL EXCRETION;
        
        Temps = 0

        Dim glupt(tsim / tMig), gluit(tsim / tMig), inspt(tsim / tMig), insrt(tsim / tMig), actiont(tsim / tMig) As Double
        Dim escrez(tsim / dt), prodt(tsim / tMig), utilt(tsim / tMig), asst(tsim / tMig), secrt(tsim / tMig), escrt(tsim / tMig), absit(tsim / tMig) As Double
        Dim insulina(tsim / dt) As Double '(tsim / tmig)
        Dim t0 As Double: t0 = 0
        Dim Tspan() As Double, w As Double
        
        Dim ys(tsim / dt, 16) As Double, last As Double: last = 0
        Dim dery(16) As Double, ind As Double
        Dim dosekempt As Double: dosekempt = 90000 'mg
        Dim escr, aa, cc, qsto, kgut, HE As Double
        For j = 0 To 3 Step 1
            w = 0
            Erase Tspan
            'ReDim Tspan((Tdose(j + 1) - Tdose(j)) / dt - 1)
            For num = TDose(j) To TDose(j + 1) Step dt
                ReDim Preserve Tspan(w)
                Tspan(w) = num
                w = w + 1
            Next
            If j = 0 Then

                ' Initialize state variables
                ys(0, 0) = Dose(0)
                ys(0, 1) = 0
                ys(0, 2) = 0
                ys(0, 3) = Gpb
                ys(0, 4) = Gtb
                ys(0, 5) = gb 'Gpb
                ys(0, 6) = isc1_ss
                ys(0, 7) = ilb
                ys(0, 8) = ipb
                ys(0, 9) = 0
                ys(0, 10) = ib
                ys(0, 11) = ib
                ys(0, 12) = isc2_ss
                ys(0, 13) = 0
                ys(0, 14) = 0
                ys(0, 15) = 0
                ys(0, 16) = 0
            Else
                ys(last, 0) = ys(last, 0) + Dose(j) 'diabetico j>1
                ys(last, 6) = ys(last, 6) + dose_i(j) / BW
            End If
            If Dose(j) > 0 Then
                dosekempt = ys(last, 0) + ys(last, 1)
            End If
            For ind = 0 To UBound(Tspan) - 1 '.Length - 1
                ' RENAL EXCRETION
                If ys(last, 3) > ke2 Then
                    escr = ke1 * (ys(last, 3) - ke2)
                Else
                    escr = 0
                End If
                ' PLASMA GLUCOSE
                dery(3) = Max((kp1 - kp2 * ys(last, 3) - kp3 * ys(last, 11)), 0) + Max(kabs * ys(last, 2) * f / BW, 0) - Fsnc - K1 * ys(last, 3) + K2 * ys(last, 4) - escr
                dery(4) = K1 * ys(last, 3) - K2 * ys(last, 4) - (Vm0 + Vmx * ys(last, 9)) / (km0 + ys(last, 4)) * ys(last, 4)
                ' ASSORBIMENTO
                dery(0) = -kmax * ys(last, 0)
                qsto = ys(last, 0) + ys(last, 1)
                If (Dose(0) = 0 And Dose(1) = 0 And Dose(2) = 0 And Dose(3) = 0 And Dose(4) = 0 And Dose(5) = 0) Or j = 0 Then
                    dery(1) = 0
                    dery(2) = 0
                Else
                    aa = 5 / 2 / (1 - b) / dosekempt
                    cc = 5 / 2 / d / dosekempt
                    kgut = kmin + (kmax - kmin) / 2 * (Tanh(aa * (qsto - b * dosekempt)) - Tanh(cc * (qsto - d * dosekempt)) + 2)
                    dery(1) = kmax * ys(last, 0) - ys(last, 1) * kgut
                    dery(2) = kgut * ys(last, 1) - kabs * ys(last, 2)
                End If
                ' Ginterstizio (ritardo=10 min rispetto al plasma)
                dery(5) = -0.1 * ys(last, 5) + 0.1 * (ys(last, 3) / Vg)
                'INSULIN) in SC pmol/Kg

                '********************************************
                '* El modelo trabaja en mg/Kg y el codigo del
                '* device lo hace en mM. La conversión es
                '* 1mM=0.018 gr/dL * Vg dL/Kg
                '********************************************
'omega1 = 20
'omega1 = 13 '15
                dery(13) = gamma1 * (0.018 * Vg * ys(last, 5)) ^ n1 / (1 + alpha1 * (0.018 * Vg * ys(last, 5) / omega1) ^ n1) - gamma2 * ys(last, 13) ^ n2 / (1 + (ys(last, 13) / omega2) ^ n2) ' glucosa sottocute
                dery(14) = gamma2 * ys(last, 13) ^ n1 / (1 + (ys(last, 13) / omega2) ^ n2) - kappa1 * ys(last, 14)
                'dery(15) = kappa1 * ys(last, 14) - kappa2 * ys(last, 15) ' consumption ADDED
                'dery(6) = kappa2 * ys(last, 15) - (kd + ka1) * ys(last, 6) 'insulin from implant comes here

'INTRODUIR UNA FUNCION DE L IMPLANT input=ys 5 ouput=Ipumpb
                '  ---
                ' eva = ys(last, 5)
                   Kglu = (gamma * (ys(last, 5) / 2) ^ zz) / (1 + w1 * (ys(last, 5) / 2) ^ zz)
                    '--- Eq's ---
                    'K1_mA = (Kglu * G1 - d_mA * mA) * IncT
                   ' K1_pI = (KpI * mA - Kins * pI) * IncT
                    'K1_ins_in = (Kins * pI - Kex * ins_in) * IncT
                   ' K1_ins_ex =
                   ' K1_N = (r * N * (1 - N / kk)) * IncT
                    
                    New_mA = mA + (Kglu * G1 - d_mA * mA) * IncT
                    New_pI = pI + (KpI * mA - Kins * pI) * IncT
                    New_ins_in = ins_in + (Kins * pI - Kex * ins_in) * IncT
                    New_ins_ex = (Kex * N * ins_in * (Vcel / Vext)) * IncT
                    New_N = N + (r * N * (1 - N / kk)) * IncT
                            
                   
                    ' ----
                'Call implant(ys(last, 5), mA, pI, ins_in, ins_ex, N)
                 rate_ins = tamany * ins_ex
                insulina(last) = rate_ins
                dery(6) = -(kd + ka1) * ys(last, 6) + rate_ins   'Version original
'dery(6) = FacS * kappa1 * ys(last, 14) - (kd + ka1) * ys(last, 6)  'Cambio mio
'dery(15) = FacS * kappa1 * ys(last, 14)
                dery(12) = kd * ys(last, 6) - ka2 * ys(last, 12)

                ' PLASMA(INSULIN)
                dery(7) = -(m1 + m3) * ys(last, 7) + m2 * ys(last, 8)
                dery(8) = m1 * ys(last, 7) - (m2 + m4) * ys(last, 8) + ka1 * ys(last, 6) + ka2 * ys(last, 12)
                ' INSULIN ACTION in pmol/L
                dery(9) = -p2U * ys(last, 9) + p2U * (ys(last, 8) / Vi - ib)
                ' PRODUCTION Delayed Insulin for PRODUCTION in pmol/L
                dery(10) = ki * (ys(last, 8) / Vi - ys(last, 10))
                dery(11) = ki * (ys(last, 10) - ys(last, 11))
                ' EULER
                For ind2 = 0 To 15
                    ys(last + 1, ind2) = ys(last, ind2) + dt * dery(ind2)
                Next
                    
                    mA = New_mA
                    pI = New_pI
                    ins_in = New_ins_in
                    ins_ex = New_ins_ex
                    N = New_N
    
                
'cv = cv + 1
'Chart1.Column = 1
'Chart1.RowCount = cv
'Chart1.Row = cv
'Chart1.RowLabel = ind * 10
'Chart1.Data = ys(last, 6)
                
                last = last + 1
                DoEvents
            Next
            DoEvents

        Next

        ' Plotting variables, average over 10 Minutes
        For ind = 0 To (tsim / tMig) - 1
            DoEvents
            glupt(ind) = ys(ind * tMig / dt, 3) * dt / tMig / Vg
            gluit(ind) = ys(ind * tMig / dt, 4) * dt / tMig
            inspt(ind) = ys(ind * tMig / dt, 8) * dt / tMig / Vi
            insrt(ind) = ys(ind * tMig / dt, 11) * dt / tMig
            actiont(ind) = ys(ind * tMig / dt, 9) * dt / tMig
            absit(ind) = (ka1 * ys(ind * tMig / dt, 6) + ka2 * ys(ind * tMig / dt, 12)) * dt / tMig
            ' assorbimento in mg/(Kg*min)
            asst(ind) = kabs * ys(ind * tMig / dt, 2) * f / BW * dt / tMig
            For ind2 = 1 To tMig / dt - 1
                ' glup=plasmatic glucose in mg/dl
                glupt(ind) = glupt(ind) + (ys(ind * tMig / dt + ind2, 3)) * dt / tMig / Vg
                ' organi insulino dipendenti in mg/kg
                gluit(ind) = gluit(ind) + (ys(ind * tMig / dt + ind2, 4)) * dt / tMig
                ' insp=plasmatic insulin and insr=remote insulin in pmol/L
                inspt(ind) = inspt(ind) + (ys(ind * tMig / dt + ind2, 8)) * dt / tMig / Vi
                insrt(ind) = insrt(ind) + (ys(ind * tMig / dt + ind2, 11)) * dt / tMig
                actiont(ind) = actiont(ind) + ys(ind * tMig / dt + ind2, 9) * dt / tMig
                ' assorbimento in mg/(Kg*min)
                asst(ind) = asst(ind) + kabs * ys(ind * tMig / dt + ind2, 2) * f / BW * dt / tMig
                ' assorbimento s.c. di insulina in pmol/(kg*min)
                absit(ind) = absit(ind) + (ka1 * ys(ind * tMig / dt + ind2, 6) + ka2 * ys(ind * tMig / dt + ind2, 12)) * dt / tMig
            Next

            '***************************************************
            '  SALIDA POR PANTALLA
            '***************************************************
            'If ind Mod 2 = 0 Then
             '   cv = cv + 1
              '  Chart1.Column = 1
               ' Chart1.RowCount = cv
                'Chart1.Row = cv
            '    Chart1.RowLabel = ind * tmig
             '   Chart1.Data = glupt(ind)
                
              '  Chart1.Column = 2
               ' Chart1.Data = inspt(ind)
            'End If
            
            
            'Chart1.Series("Series1").Points.AddXY(ind * tmig, glupt(ind))
            'Debug.Print (glupt(ind))
            ' prod in mg/(Kg*min)
            prodt(ind) = Max(kp1 - kp2 * glupt(ind) * Vg - kp3 * insrt(ind), 0)                       ' util in mg/(Kg*min)
            utilt(ind) = Fsnc + (Vm0 + Vmx * actiont(ind)) / (km0 + gluit(ind)) * gluit(ind)
        Next
        'Chart1.Titles.Add ("D. Type I open loop [Glucose] plasma")
        ' escrezione renale
        For ind = 0 To last - 1
            If ys(ind, 3) > ke2 Then
                escrez(ind) = ke1 * (ys(ind, 3) * Vg - ke2)
            Else
                escrez(ind) = 0
            End If
        Next
        
        
        '----- guardar en document glupt(ind)
        mm = FreeFile
        
        Open "\\VBOXSVR\Compartit_Windows\glucose_" & contador & ".txt" For Output As mm
        'For Index = 1 To last - 1 Step 100
        For Index = 0 To UBound(glupt) Step 1
            Print #mm, Index, glupt(Index), inspt(Index)
            'Print #mm, Index, ys(Index, 5), ys(Index, 6), ys(Index, 8), Kglu
          Next Index
        Close #mm
    
        mm = FreeFile
        Open "\\VBOXSVR\Compartit_Windows\glucose_SUB_" & contador & ".txt" For Output As mm
        'For Index = 1 To last - 1 Step 100
        For Index = 0 To last - 1 Step 1000 'UBound(glupt) Step 1
            'Print #mm, Index * dt, glupt(Index), inspt(Index), rate_ins
            Print #mm, Index * dt, ys(Index, 5), insulina(Index)
          Next Index
        Close #mm
End Sub


Function canvivars_seg(ab, cd, ef, gh, ij, contador)
'Private Sub Command1_Click()
       '***************************
        '   DM1
        '***************************
        '----- Implant ----
        Dim K1_mA As Double, K1_pI As Double, K1_ins_in As Double, K1_ins_ex As Double, K1_N As Double
        Dim New_mA As Double, New_pI As Double, New_ins_in As Double, New_ins_ex As Double, New_N As Double
        
        Const IncT As Double = 0.005
        Const gamma As Single = 1 * 10 ^ -8 / 0.1
        Const w1 As Single = 1 * 10 ^ -6.5
        Const exp As Double = 3
        Const timescale As Single = 0.1
        Const d_insi As Single = 0 ';0.03/timescale'; % degradacio ins interna
        Const d_inse As Single = 0 '0.004 / timescale '; % degradacio ins externa
        Const r As Single = 0.1 / timescale '
        Const Vcel As Single = 1 / timescale '; % volum celula
        Const Vext As Single = 100 / timescale '; % volum medi
        Const K As Single = 1 '100 / timescale '; % volum medi
        
        Dim G1 As Double '
        G1 = ab ' 10
        Dim d_mA As Single '
        d_mA = cd ' 0.00009 / timescale '; % missatger pro ins
        Dim KpI As Single
        KpI = ef '0.1 / timescale '; % mA->pI
        Dim Kins As Single
        Kins = gh '1 / timescale '; % pins->in
        Dim Kex As Single
        Kex = ij ' 0.1 / timescale ';% % secrecio ins a fora
         

        Const tamany As Double = 100 '0
        Dim mA As Double '= 0
        Dim pI As Double '= 0
        Dim ins_in As Double '= 0
        Dim ins_ex As Double '= 0
        Dim N As Double '= 0.1
        N = 0.1
        Dim rate_ins As Double  ' 0


        Const status As Double = 1
        Const openclosed As Double = 0
        Const Tdose1 As Double = (24 + 6) * 60 'min
        Const Tdose2 As Double = (24 + 12) * 60 'min
        Const Tdose3 As Double = (24 + 20) * 60  'min

        Const Dose1 As Double = 45000 'mg
        Const Dose2 As Double = 70000 'mg
        Const Dose3 As Double = 70000 'mg

        Const BW As Double = 78 'kg
        Const ib As Single = 41.32 'pmol/l
        Const EGPb As Single = 2.4 'mg/kg/min
        Const gb As Double = 120 'mg/dl
        ' tempi
        Const dt As Single = 0.005 ' time step min
        Const tsim As Double = 150 * 60 'min ¿? son 150 hores
        'Const Tdose4=(Tdose3+tsim)/2 'min
        
        Dim TDose(4) As Double
        TDose(0) = 0
        TDose(1) = Tdose1
        TDose(2) = Tdose2
        TDose(3) = Tdose3
        TDose(4) = tsim

        ' Dosi
        Dim Dose(5) As Double
        Dose(0) = 0
        Dose(1) = Dose1
        Dose(2) = Dose2
        Dose(3) = Dose3
        Dose(4) = 0

        'Dosi insulina
        Const Ipumpb As Double = 1
        Const dose_ins1 As Double = 3
        Const dose_ins2 As Double = 5
        Const dose_ins3 As Double = 5
        Dim dose_i(4) As Double
        dose_i(0) = 0
        dose_i(1) = 0 ' dose_ins1 * 6000
        dose_i(2) = 0 ' dose_ins2 * 6000
        dose_i(3) = 0 ' dose_ins3 * 6000
        dose_i(4) = 0
        
        '  Plasma meal rate of appearance parameters
        Const kabs As Single = 0.0568 'min^-1
        Const kmax As Single = 0.0558 'min^-1
        Const kmin As Single = 0.008 'kmin min^-1
        Const f As Single = 0.9 'dimensionless
        Const b As Single = 0.82 'dimensionless
        Const d As Single = 0.01 'dimensionless

        Const Vg As Single = 1.88 'Glucose Volume in dl/Kg
        Const Vi As Single = 0.05 'Insulin Volume in L/kg

        Const Gpb As Single = gb * Vg 'amount of glucose in the plasma compartment in the ss
        Const ipb As Single = ib * Vi 'pmol/kg amount of plasma insulin
        ' Excretion
        Const ke1 As Single = 0.0005 'min^-1
        Const ke2 As Double = 339 'mg/kg
        'Utilization Parameters
        Const Vmx As Single = 0.047 'mg/kg/min per pmol/L
        Const Vmx_normal As Single = 0.047
        Const km0 As Single = 225.59 'mg/kg
        Const K2 As Single = 0.079 'min^-1
        Const K1 As Single = 0.065 'min^-1
        Const Fsnc As Single = 1 'mg/kg/min

        Const Gtb As Single = (Fsnc - EGPb + K1 * Gpb) / K2 'mg/kg
        Const Vm0 As Single = (EGPb - Fsnc) * (km0 + Gtb) / Gtb 'mg/kg/min
        Const Rdb As Single = EGPb 'mg/kg/min
        Const PCRb As Single = Rdb / gb 'dl/kg/min
        'Insulin Action Parameter
        Const p2U = 0.0331 'min^-1
        ' Plasma Insulin & Insulin Secretion
        Const m1 As Single = 0.19 'min^-1
        Const m2 As Single = 0.484 'min^-1
        Const m5 As Single = 0.0304 'min*kg/pmol
        Const HEb As Single = 0.6 'dimensionless
        Const m4 As Single = 2 / 3 * m2 * HEb 'min^-1
        Const m3 As Double = HEb * m1 / (1 - HEb)
        Const ilb As Double = ipb * m2 / (m1 + m3) 'pmol/kg
        Const SRb As Double = ipb * m4 + ilb * m3 'pmol/kg/min
        ' Production Parameters
        Const ki As Single = 0.0079 'min^1
        Const kp2 As Single = 0.0021 'min^-1
        Const kp3 As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp3_normal As Single = 0.009 'mg/kg/min/(pmol/l)
        Const kp1 As Single = EGPb + kp2 * Gpb + kp3 * ib 'mg/kg/min
        ' PID
        Const Kp As Single = 2.5 'pmol/min per mg/dl (da Steil 2006 Kp=0.025 U/h per mg/dl)
        Const TI As Double = 450 'min da Steil uomo 2006
        Const TD As Double = 66 'min da Steil uomo 2006
        ' insulin infusion
        Const ksc As Single = 0.01 '0.03; %min^-1 =1/tau2 di Steil 2006: tau2=33.5 min in media -> k=0.03 min^-1
        Const kd As Single = 0.0164
        Const ka1 As Single = 0.0018
        Const ka2 As Single = 0.0182
        Const isc1_ss As Double = 0
        Const isc2_ss As Double = 0
        
        ' Cell implant
        Dim gamma1 As Double: gamma1 = 0.0033
        Dim alpha1 As Double: alpha1 = 12.2
        Dim gamma2 As Double: gamma2 = 0.00054
        Dim n1 As Double: n1 = 2
        Dim n2 As Double: n2 = 1.5
        Dim omega1 As Double: omega1 = 10
        Dim omega2 As Double: omega2 = 50
        Dim kappa1 As Double: kappa1 = 0.00385
        Dim kappa2 As Double: kappa2 = 0.0002
        Const FacS As Double = 15

        ' Attenzione:
        '  ys(0)=stomaco1 (q11 in SAAM)
        '  ys(1)=stomaco2 (q12  in SAAM)
        '  ys(2)=intestino (q13 in SAAM)
        '  ys(3)=glucosio plasmatico+organi insulino-indipendenti (q1 in SAAM)
        '  ys(4)=glucosio+organi insulino dipendenti (q2 in SAAM)

        'NUOVO       ys(5)=glucosio nel sottocute in mg/kg
        'NUOVO       ys(6)=insulina NON-MONOMERICA nel sottocute Isc1 in pmol/kg

        '  ys(7)=insulina nel fegato in pmol/kg (q14 in SAAM)
        '  ys(8)=insulina plasmatica in pmol/kg (q7 in SAAM)
        '  ys(9)=azione insulinica in pmol/L (q4 in SAAM)
        '  ys(10)=insulina ritardata 1 in pmol/L (q9 in SAAM)
        '  ys(11)=insulina ritardata 2 in pmol/L(q10 in SAAM)

        ' NUOVISSIMO  ys(12)=insulina MONOMERICA nel sottocute Isc2 in pmol/kg

        ' RENAL EXCRETION;
        
        Temps = 0

        Dim glupt(tsim / 10), gluit(tsim / 10), inspt(tsim / 10), insrt(tsim / 10), actiont(tsim / 10) As Double
        Dim escrez(tsim / dt), prodt(tsim / 10), utilt(tsim / 10), asst(tsim / 10), secrt(tsim / 10), escrt(tsim / 10), absit(tsim / 10) As Double
        Dim insulina(tsim / dt) As Double '(tsim / 10)
        Dim t0 As Double: t0 = 0
        Dim Tspan() As Double, w As Double
        
        Dim ys(tsim / dt, 16) As Double, last As Double: last = 0
        Dim dery(16) As Double, ind As Double
        Dim dosekempt As Double: dosekempt = 90000 'mg
        Dim escr, aa, cc, qsto, kgut, HE As Double
        For j = 0 To 3 Step 1
            w = 0
            Erase Tspan
            'ReDim Tspan((Tdose(j + 1) - Tdose(j)) / dt - 1)
            For num = TDose(j) To TDose(j + 1) Step dt
                ReDim Preserve Tspan(w)
                Tspan(w) = num
                w = w + 1
            Next
            If j = 0 Then

                ' Initialize state variables
                ys(0, 0) = Dose(0)
                ys(0, 1) = 0
                ys(0, 2) = 0
                ys(0, 3) = Gpb
                ys(0, 4) = Gtb
                ys(0, 5) = Gpb
                ys(0, 6) = isc1_ss
                ys(0, 7) = ilb
                ys(0, 8) = ipb
                ys(0, 9) = 0
                ys(0, 10) = ib
                ys(0, 11) = ib
                ys(0, 12) = isc2_ss
                ys(0, 13) = 0
                ys(0, 14) = 0
                ys(0, 15) = 0
                ys(0, 16) = 0
            Else
                ys(last, 0) = ys(last, 0) + Dose(j) 'diabetico j>1
                ys(last, 6) = ys(last, 6) + dose_i(j) / BW
            End If
            If Dose(j) > 0 Then
                dosekempt = ys(last, 0) + ys(last, 1)
            End If
            For ind = 0 To UBound(Tspan) - 1 '.Length - 1
                ' RENAL EXCRETION
                If ys(last, 3) > ke2 Then
                    escr = ke1 * (ys(last, 3) - ke2)
                Else
                    escr = 0
                End If
                ' PLASMA GLUCOSE
                dery(3) = Max((kp1 - kp2 * ys(last, 3) - kp3 * ys(last, 11)), 0) + Max(kabs * ys(last, 2) * f / BW, 0) - Fsnc - K1 * ys(last, 3) + K2 * ys(last, 4) - escr
                dery(4) = K1 * ys(last, 3) - K2 * ys(last, 4) - (Vm0 + Vmx * ys(last, 9)) / (km0 + ys(last, 4)) * ys(last, 4)
                ' ASSORBIMENTO
                dery(0) = -kmax * ys(last, 0)
                qsto = ys(last, 0) + ys(last, 1)
                If (Dose(0) = 0 And Dose(1) = 0 And Dose(2) = 0 And Dose(3) = 0 And Dose(4) = 0 And Dose(5) = 0) Or j = 0 Then
                    dery(1) = 0
                    dery(2) = 0
                Else
                    aa = 5 / 2 / (1 - b) / dosekempt
                    cc = 5 / 2 / d / dosekempt
                    kgut = kmin + (kmax - kmin) / 2 * (Tanh(aa * (qsto - b * dosekempt)) - Tanh(cc * (qsto - d * dosekempt)) + 2)
                    dery(1) = kmax * ys(last, 0) - ys(last, 1) * kgut
                    dery(2) = kgut * ys(last, 1) - kabs * ys(last, 2)
                End If
                ' Ginterstizio (ritardo=10 min rispetto al plasma)
                dery(5) = -0.1 * ys(last, 5) + 0.1 * ys(last, 3)
                'INSULIN) in SC pmol/Kg

                '********************************************
                '* El modelo trabaja en mg/Kg y el codigo del
                '* device lo hace en mM. La conversión es
                '* 1mM=0.018 gr/dL * Vg dL/Kg
                '********************************************
'omega1 = 20
'omega1 = 13 '15
                dery(13) = gamma1 * (0.018 * Vg * ys(last, 5)) ^ n1 / (1 + alpha1 * (0.018 * Vg * ys(last, 5) / omega1) ^ n1) - gamma2 * ys(last, 13) ^ n2 / (1 + (ys(last, 13) / omega2) ^ n2) ' glucosa sottocute
                dery(14) = gamma2 * ys(last, 13) ^ n1 / (1 + (ys(last, 13) / omega2) ^ n2) - kappa1 * ys(last, 14)
                'dery(15) = kappa1 * ys(last, 14) - kappa2 * ys(last, 15) ' consumption ADDED
                'dery(6) = kappa2 * ys(last, 15) - (kd + ka1) * ys(last, 6) 'insulin from implant comes here

'INTRODUIR UNA FUNCION DE L IMPLANT input=ys 5 ouput=Ipumpb
                '  ---
                   Ka = (gamma * (ys(last, 5) / 2) ^ exp) / (1 + w1 * (ys(last, 5) / 2) ^ exp)
                    '--- Eq's ---
                    K1_mA = (Ka * G1 - d_mA * mA) * IncT
                    K1_pI = (KpI * mA - Kins * pI) * IncT
                    K1_ins_in = (Kins * pI - d_insi * ins_in - Kex * ins_in) * IncT
                    K1_ins_ex = (Kex * N * ins_in * (Vcel / Vext) - d_inse * ins_ex) * IncT
                    K1_N = (r * N * (1 - N / K)) * IncT
                    
                    New_mA = mA + K1_mA
                    New_pI = pI + K1_pI
                    New_ins_in = ins_in + K1_ins_in
                    New_ins_ex = K1_ins_ex
                    New_N = N + K1_N
                            
                   
                    ' ----
                'Call implant(ys(last, 5), mA, pI, ins_in, ins_ex, N)
                 rate_ins = tamany * ins_ex
                insulina(last) = rate_ins
                dery(6) = -(kd + ka1) * ys(last, 6) + tamany * ins_ex  'Version original
'dery(6) = FacS * kappa1 * ys(last, 14) - (kd + ka1) * ys(last, 6)  'Cambio mio
'dery(15) = FacS * kappa1 * ys(last, 14)
                dery(12) = kd * ys(last, 6) - ka2 * ys(last, 12)

                ' PLASMA(INSULIN)
                dery(7) = -(m1 + m3) * ys(last, 7) + m2 * ys(last, 8)
                dery(8) = m1 * ys(last, 7) - (m2 + m4) * ys(last, 8) + ka1 * ys(last, 6) + ka2 * ys(last, 12)
                ' INSULIN ACTION in pmol/L
                dery(9) = -p2U * ys(last, 9) + p2U * (ys(last, 8) / Vi - ib)
                ' PRODUCTION Delayed Insulin for PRODUCTION in pmol/L
                dery(10) = ki * (ys(last, 8) / Vi - ys(last, 10))
                dery(11) = ki * (ys(last, 10) - ys(last, 11))
                ' EULER
                For ind2 = 0 To 15
                    ys(last + 1, ind2) = ys(last, ind2) + dt * dery(ind2)
                Next
                    
                    mA = New_mA
                    pI = New_pI
                    ins_in = New_ins_in
                    ins_ex = New_ins_ex
                    N = New_N
    
                
'cv = cv + 1
'Chart1.Column = 1
'Chart1.RowCount = cv
'Chart1.Row = cv
'Chart1.RowLabel = ind * 10
'Chart1.Data = ys(last, 6)
                
                last = last + 1
                DoEvents
            Next
            DoEvents

        Next

        ' Plotting variables, average over 10 Minutes
        For ind = 0 To (tsim / 10) - 1
            DoEvents
            glupt(ind) = ys(ind * 10 / dt, 3) * dt / 10 / Vg
            gluit(ind) = ys(ind * 10 / dt, 4) * dt / 10
            inspt(ind) = ys(ind * 10 / dt, 8) * dt / 10 / Vi
            insrt(ind) = ys(ind * 10 / dt, 11) * dt / 10
            actiont(ind) = ys(ind * 10 / dt, 9) * dt / 10
            absit(ind) = (ka1 * ys(ind * 10 / dt, 6) + ka2 * ys(ind * 10 / dt, 12)) * dt / 10
            ' assorbimento in mg/(Kg*min)
            asst(ind) = kabs * ys(ind * 10 / dt, 2) * f / BW * dt / 10
            For ind2 = 1 To 10 / dt - 1
                ' glup=plasmatic glucose in mg/dl
                glupt(ind) = glupt(ind) + (ys(ind * 10 / dt + ind2, 3)) * dt / 10 / Vg
                ' organi insulino dipendenti in mg/kg
                gluit(ind) = gluit(ind) + (ys(ind * 10 / dt + ind2, 4)) * dt / 10
                ' insp=plasmatic insulin and insr=remote insulin in pmol/L
                inspt(ind) = inspt(ind) + (ys(ind * 10 / dt + ind2, 8)) * dt / 10 / Vi
                insrt(ind) = insrt(ind) + (ys(ind * 10 / dt + ind2, 11)) * dt / 10
                actiont(ind) = actiont(ind) + ys(ind * 10 / dt + ind2, 9) * dt / 10
                ' assorbimento in mg/(Kg*min)
                asst(ind) = asst(ind) + kabs * ys(ind * 10 / dt + ind2, 2) * f / BW * dt / 10
                ' assorbimento s.c. di insulina in pmol/(kg*min)
                absit(ind) = absit(ind) + (ka1 * ys(ind * 10 / dt + ind2, 6) + ka2 * ys(ind * 10 / dt + ind2, 12)) * dt / 10
            Next

            '***************************************************
            '  SALIDA POR PANTALLA
            '***************************************************
            'If ind Mod 2 = 0 Then
             '   cv = cv + 1
              '  Chart1.Column = 1
               ' Chart1.RowCount = cv
                'Chart1.Row = cv
            '    Chart1.RowLabel = ind * 10
             '   Chart1.Data = glupt(ind)
                
              '  Chart1.Column = 2
               ' Chart1.Data = inspt(ind)
            'End If
            
            
            'Chart1.Series("Series1").Points.AddXY(ind * 10, glupt(ind))
            'Debug.Print (glupt(ind))
            ' prod in mg/(Kg*min)
            prodt(ind) = Max(kp1 - kp2 * glupt(ind) * Vg - kp3 * insrt(ind), 0)                       ' util in mg/(Kg*min)
            utilt(ind) = Fsnc + (Vm0 + Vmx * actiont(ind)) / (km0 + gluit(ind)) * gluit(ind)
        Next
        'Chart1.Titles.Add ("D. Type I open loop [Glucose] plasma")
        ' escrezione renale
        For ind = 0 To last - 1
            If ys(ind, 3) > ke2 Then
                escrez(ind) = ke1 * (ys(ind, 3) * Vg - ke2)
            Else
                escrez(ind) = 0
            End If
        Next
        
        
        '----- guardar en document glupt(ind)
        mm = FreeFile
        Open "E:\glucose_" & contador & ".txt" For Output As mm
        For Index = 1 To last - 1 Step 100
           
            Print #mm, Index, ys(Index, 3), insulina(Index)
          Next Index
        Close #mm
    
        
End Function
