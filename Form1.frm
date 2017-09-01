VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Analisi dades i calcul minim"
      Height          =   975
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Optimitzacio per criteris diabetis 1000 cels"
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Dm1 +implant rnd punt MINIM"
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dm1 +implant rnd punt MAXIM"
      Height          =   675
      Left            =   4560
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pacient Sa"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DM1 Implant"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Generador 
      Caption         =   "generador parametres"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "DM1 + I rnd parametres"
      Height          =   495
      Left            =   3120
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double


Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double

nn = FreeFile
myfile = "E:\parametresatzar.txt"
Open myfile For Input As #nn

Dim Sa() As Variant
Dim Diab() As Variant

Sa = pacient_sa()

'Print UBound(Sa)
'For a = 1 To UBound(Sa) Step 100
 '   Print Sa(a)
'Next a

For linia = 1 To 1 ' EOF(nn) canvier per while not eof
    'fitness(1) = 0
    sum_fit1 = 1E+21
    sum_fit2 = 1E+21
    Line Input #nn, textline
    
    splitvalues = Split(textline, ",")
    un = Val(splitvalues(0))
    dos = Val(splitvalues(1))
    tres = Val(splitvalues(2))
    quatre = Val(splitvalues(3))
   '---- buscar dos poblacions
    
    poblacio1 = 1
    poblacio2 = 100000000
    
   Do While sum_fit1 > 100000  '& sum_fit2 < 1000
    sum_fit1 = 0 '1E+21
    sum_fit2 = 0 '1E+21
   ' Print sum_fit1
   ' Print sum_fit2
        Diab1 = diabetic_random(un, dos, tres, quatre, poblacio1, linia, poblacio1)
        diab2 = diabetic_random(un, dos, tres, quatre, poblacio2, linia, poblacio2)
        For j = 1 To UBound(Sa)
            fitness1 = Diab1(j) - Sa(j)
            sum_fit1 = sum_fit1 + fitness1
            
            fitness2 = diab2(j) - Sa(j)
            sum_fit2 = sum_fit2 + fitness2
           
        Next j
       
        Print sum_fit1
        Print sum_fit2
        
        If sum_fit1 > 0 Then
            poblacio1 = poblacio1
            Print poblacio1
        Else
            poblacio1 = (poblacio1 + poblacio2) / 2
            Print poblacio1
        End If
        
        If sum_fit2 > 0 Then
            poblacio2 = poblacio2
           Print poblacio2
       Else
            poblacio2 = (poblacio1 + poblacio2) / 2
           Print poblacio2
        End If
    
   Loop
    
    
        
        
    
    
    '- arreglar la poblacio
  
    'MsgBox (linia)
Next linia


' a
MsgBox ("Fi")
End Sub



Private Sub Command2_Click()

    Call diabetic(0) '104960) '1000000 / 2) '5000)
    MsgBox "Fi"
End Sub


Private Sub Command3_Click()
Dim N() As Variant
    N = pacient_sa()
    MsgBox "Fi Pacient sa"
End Sub

Private Sub Command4_Click()


Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double
Dim maxim_sa As Double, maxima_diab1 As Double, maxim_diab2 As Double, maxim_diab_mig As Double
Dim K As Double, kk As Double
Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double
Dim Sa() As Variant
Dim Diab_1() As Variant
Dim Diab_2() As Variant
Dim diab_mig() As Variant

nn = FreeFile
myfile = "\\VBOXSVR\Compartit_Windows\diab\parametresatzar.txt"
Open myfile For Input As #nn

Sa = pacient_sa()
maxim_sa = 0
minim_sa = 1000


For K = (30 * 60) To UBound(Sa) - 1 Step 1

    'Maxim_sa = Max(Sa) -> bucle que recorri el vector i trii el maxim.
    If Sa(K) > maxim_sa Then
        maxim_sa = Sa(K)
    End If
    ' minim_sa =min(sa)
    If Sa(K) < minim_sa Then
        minim_sa = Sa(K)
    End If
Next K
linia = 0
Do While Not EOF(nn)
    linia = linia + 1
'For linia = 5 To EOF(nn)  'canvier per while not eof
    Line Input #nn, textline 'llegeix la linia
    Print linia
    'splitvalues = Split(textline, " ") 'separa la linia per comes i separa en els 4 parametres
    pp = InStr(1, textline, ",")
    dos = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    tres = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    quatre = Val(Trim(Left(textline, pp - 1)))
    
    pp = InStr(1, textline, ",")
    cinc = Val(Trim(Left(textline, pp - 1)))
    
    

    'un = (splitvalues(0))
    'dos = Val(splitvalues(1))
    'tres = Val(splitvalues(2))
    'quatre = Val(splitvalues(3))
    'cinc = Val(splitvalues(4))
   '---- buscar dos poblacions
    
    poblacio1 = 1
    poblacio2 = 1000000 ' 10000
    cont = 0
    dif_1 = 100
    dif_2 = 100
    poblacio2 = poblacio1
    Do
        poblacio2 = poblacio2 * 2
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        maxim_diab2 = -1E+99
        For kk = (24 * 60) To UBound(Diab_2) - 1 Step 1
            If Diab_2(kk) > maxim_diab2 Then
                maxim_diab2 = Diab_2(kk)
            End If
        Next kk
        If maxim_diab2 < maxim_sa Then Exit Do
    Loop
    
    poblacio1 = Int(poblacio2 / 2)
    
Do While (Abs(dif_1) > 1)  ' Or (cont < 50) ' condicio de sortida que depen de la diferencia entre pic diab i pic sa

    cont = cont + 1
    Form1.Caption = "celula: " & linia & " iteracio: " & cont
   ' Print cont
        poblacio_mig = Int((poblacio1 + poblacio2) / 2) ' calcula poblacio entre mig de la 1 i la 2
 'Print poblacio1
 
     
        Diab_1 = diabetic_random(dos, tres, quatre, cinc, poblacio1, linia, poblacio1)
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        diab_mig = diabetic_random(dos, tres, quatre, cinc, poblacio_mig, linia, poblacio_mig)

'inicialitzar variables de trobar el maxim
        maxim_diab1 = -1E+99 ' maxim_sa '0
        maxim_diab2 = -1E+99 ' maxim_sa '0 '-100000000
        maxim_diab_mig = -1E+99 ' maxim_sa '0 ' -100000000
        
' trobar el maxim de cada un i fer per una banda resta amb el maxim sa
' per laltre comprobar els signes de la resta per saber les noves N1 i N2
    
    For kk = (24 * 60) To UBound(Diab_1) - 1 Step 1  ' he de començar quan comença a menjar
        'Maxim_sa = Max(Sa) -> bucle que recorri el vector i trii el maxim.
        If Diab_1(kk) > maxim_diab1 Then
            maxim_diab1 = Diab_1(kk)
        End If
        
        If Diab_2(kk) > maxim_diab2 Then
            maxim_diab2 = Diab_2(kk)
        End If
        
        If diab_mig(kk) > maxim_diab_mig Then
            maxim_diab_mig = diab_mig(kk)
        End If
        
        
    Next kk

        dif_1 = Abs(maxim_diab1 - maxim_diab2)
        'dif_2 = maxim_diab2 - maxim_sa
        'dif_mig = maxim_diab_mig - maxim_sa
        
'        Print maxim_1
'    Print maxim_2
'    Print maxim_mig
        
        If maxim_diab_mig > maxim_sa Then
            poblacio1 = poblacio_mig
        Else
            poblacio2 = poblacio_mig
        End If


        
       ' If maxim_diab1 > maxim_sa Then
       '     If maxim_diab_mig > maxim_sa Then
       '         poblacio1 = poblacio_mig
       '         dif_1 = dif_mig
       '     Else
       '         poblacio1 = poblacio1
       '         dif_1 = dif_1
       '     End If
       ' End If
   '
   '     If maxim_diab2 < maxim_sa Then
   '         If maxim_diab_mig < maxim_sa Then
   '             poblacio2 = poblacio_mig
   '             dif_2 = dif_mig
   '         Else
   '             poblacio2 = poblacio2
   '             dif_2 = dif_2
   '         End If
   '     End If
        
        
       ' Print dif_1
        'Print Abs(dif_2)

        
       ' Print maxim_diab1
       ' Print dif_1
        
       ' Print maxim_diab2
       ' Print dif_2
        
        Print maxim_diab_mig
       ' Print dif_mig
   Loop
            mm = FreeFile
        
        'If Abs(dif_1) < Abs(dif_2) Then
         
            'Open "E:\diabmaximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
            
            Open "\\VBOXSVR\Compartit_Windows\diab\maximcelula_" & linia & "glucose_" & poblacio1 & ".txt" For Output As mm
            For Index = 0 To UBound(Diab_1) Step 1
                Print #mm, Index, Diab_1(Index)
            Next Index
        'Else
        
           ' Open "\\VBOXSVR\Compartit_Windows\diab\maximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
          '
            'Open "E:\diabmaximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
           ' For Index = 0 To UBound(Diab_2) Step 1
           '     Print #mm, Index, Diab_2(Index)
           ' Next Index
       ' End If
          
        Close #mm
 Loop
 'linia
    
    
    '- arreglar la poblacio
  
    'MsgBox (linia)
'Next linia


' a
'MsgBox ("Fi")
'id = Shell("E:\Codigo_Javier\Codigo_Eva\codi_minim")
id = Shell("\\VBOXSVR\Compartit_Windows\Codigo_Javier\Codigo_Eva\codi_minim")
End
End Sub






Private Sub Command5_Click()
'canviar maxim x minim

mm = FreeFile
Open "\\VBOXSVR\Compartit_Windows\diab\minimcelula3.txt " For Output As mm
Close #mm

hmenjar = (24 + 6) * 60
Print hmenjar
deju = 126
hdeju = hmenjar - 1
Print hdeju

hyper = 200
hhyper = hmenjar + (2 * 60)
Print hhyper


Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double
Dim minim_sa As Double, minim_diab1 As Double, minim_diab2 As Double, minim_diab_mig As Double
Dim K As Double, kk As Double
Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double
Dim Sa() As Variant
Dim Diab_1() As Variant
Dim Diab_2() As Variant
Dim diab_mig() As Variant

nn = FreeFile
myfile = "\\VBOXSVR\Compartit_Windows\diab\parametresatzar_3.txt"
Open myfile For Input As #nn


minim_sa = 70
Print minim_sa
    linia = 0
Do While Not EOF(nn)
    linia = linia + 1
    Print linia
'For linia = 1 To 1 ' EOF(nn) canvier per while not eof
    Line Input #nn, textline 'llegeix la linia
    
    Print linia
    'splitvalues = Split(textline, " ") 'separa la linia per comes i separa en els 4 parametres
    pp = InStr(1, textline, ",")
    dos = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    tres = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    quatre = Val(Trim(Left(textline, pp - 1)))
    
    pp = InStr(1, textline, ",")
    cinc = Val(Trim(Left(textline, pp - 1)))
    
   '---- buscar dos poblacions
    
      poblacio1 = 1
    poblacio2 = 1000000 ' 10000
    cont = 0
    dif_1 = 100
    dif_2 = 100
    poblacio2 = poblacio1
    Do
        poblacio2 = poblacio2 * 2
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        minim_diab2 = 1E+99
        For kk = (24 * 60) To UBound(Diab_2) - 2 Step 1
            If Diab_2(kk) < minim_diab2 Then
                minim_diab2 = Diab_2(kk)
            End If
        Next kk
        If minim_diab2 < minim_sa Then Exit Do
    Loop
    
    poblacio1 = Int(poblacio2 / 2)
    'Print poblacio1
    'Print poblacio2
    
    'Print minim_diab2
    Print minim_sa
    
    Do While (Abs(dif_1) > 1) 'Or Abs(dif_2) > 1 Or cont > 50) ' condicio de sortida que depen de la diferencia entre pic diab i pic sa
        cont = cont + 1
       ' Print poblacio1
        'Print poblacio2
        'Print poblacio_mig
        
        Form1.Caption = "celula: " & linia & " iteracio: " & cont
        'Print cont
            poblacio_mig = (poblacio1 + poblacio2) / 2 ' calcula poblacio entre mig de la 1 i la 2
     '  Print poblacio1
           Diab_1 = diabetic_random(dos, tres, quatre, cinc, poblacio1, linia, poblacio1)
            Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
            diab_mig = diabetic_random(dos, tres, quatre, cinc, poblacio_mig, linia, poblacio_mig)
           
    'inicialitzar variables de trobar el maxim
            minim_diab1 = 1E+99
            minim_diab2 = 1E+99
            minim_diabmig = 1E+99
            
    ' trobar el maxim de cada un i fer per una banda resta amb el maxim sa
    ' per laltre comprobar els signes de la resta per saber les noves N1 i N2
        
        For kk = (24 * 60) To UBound(Diab_1) - 1 Step 1 ' he de començar quan comença a menjar
            'Maxim_sa = Max(Sa) -> bucle que recorri el vector i trii el maxim.
        
            
            If Diab_1(kk) < minim_diab1 Then
                minim_diab1 = Diab_1(kk)
            End If
            
            If Diab_2(kk) < minim_diab2 Then
                minim_diab2 = Diab_2(kk)
            End If
            
            If diab_mig(kk) < minim_diab_mig Then
                minim_diab_mig = diab_mig(kk)
            End If
            
            
        Next kk

       
        dif_1 = Abs(minim_diab1 - minim_diab2)
       
        
        If minim_diab_mig < minim_sa Then
            poblacio2 = poblacio_mig
        Else
            poblacio1 = poblacio_mig
        End If

        
    
   Loop
   NOpt = (poblacio1 + poblacio2) / 2
   Diab = diabetic_random(dos, tres, quatre, cinc, NOpt, linia, NOpt)
    minim_diab = 1E+99
       For kk = (24 * 60) To UBound(Diab_1) - 1 Step 1 ' he de començar quan comença a menjar
        'Maxim_sa = Max(Sa) -> bucle que recorri el vector i trii el maxim.
    
        
        If Diab(kk) < minim_diab Then
            minim_diab = Diab(kk)
        End If
        
        
    Next kk

   
    
            mm = FreeFile
            Open "\\VBOXSVR\Compartit_Windows\diab\minimcelula3.txt" For Append As mm
            'Open "E:\diabmaximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
            
          Print #mm, dos, tres, quatre, cinc, NOpt, minim_diab, Diab(hdeju), Diab(hhyper)
        Close #mm
       
 Loop
 '


' a
MsgBox ("Fi")
End Sub






Private Sub Command6_Click()
'Private Sub Command4_Click()


Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double
Dim maxim_sa As Double, maxima_diab1 As Double, maxim_diab2 As Double, maxim_diab_mig As Double
Dim K As Double, kk As Double
Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double
Dim Sa() As Variant
Dim Diab_1() As Variant
Dim Diab_2() As Variant
Dim diab_mig() As Variant



Sa = pacient_sa()
maxim_sa = 0
minim_sa = 1000


For K = (30 * 60) To UBound(Sa) - 1 Step 1

    'Maxim_sa = Max(Sa) -> bucle que recorri el vector i trii el maxim.
    If Sa(K) > maxim_sa Then
        maxim_sa = Sa(K)
    End If
    ' minim_sa =min(sa)
    If Sa(K) < minim_sa Then
        minim_sa = Sa(K)
    End If
Next K
linia = 0
'Do While Not EOF(nn)
    linia = linia + 1
'For linia = 5 To EOF(nn)  'canvier per while not eof
  
    

    'un = (splitvalues(0))
    'dos = Val(splitvalues(1))
    'tres = Val(splitvalues(2))
    'quatre = Val(splitvalues(3))
    'cinc = Val(splitvalues(4))
   '---- buscar dos poblacions
   
   'Dim d_mA As Single '
        dos = 9.88 * 10 ^ -2 '0.009 / timescale  '; % missatger pro ins
       ' Dim KpI As Single
        tres = 0.75 '  0.1 / timescale '; % mA->pI
      '  Dim Kins As Single
        quatre = 0.46 ' 1 / timescale '; % pins->in
       ' Dim Kex As Single
        cinc = 0.72 '0.1 / timescale  ';% % secrecio ins a fora
    
    poblacio1 = 1
    poblacio2 = 1000000 ' 10000
    cont = 0
    dif_1 = 100
    dif_2 = 100
    
    poblacio2 = poblacio1
    Do
        poblacio2 = poblacio2 * 2
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        maxim_diab2 = -1E+99
        For kk = (24 * 60) To UBound(Diab_2) - 1 Step 1
            If Diab_2(kk) > maxim_diab2 Then
                maxim_diab2 = Diab_2(kk)
            End If
        Next kk
        If maxim_diab2 < maxim_sa Then Exit Do
    Loop
    
    poblacio1 = Int(poblacio2 / 2)
    
    Do While (Abs(dif_1) > 1)  ' Or (cont < 50) ' condicio de sortida que depen de la diferencia entre pic diab i pic sa

    cont = cont + 1
    Form1.Caption = "celula: " & linia & " iteracio: " & cont
   ' Print cont
        poblacio_mig = Int((poblacio1 + poblacio2) / 2) ' calcula poblacio entre mig de la 1 i la 2
 'Print poblacio1
 
     
        Diab_1 = diabetic_random(dos, tres, quatre, cinc, poblacio1, linia, poblacio1)
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        diab_mig = diabetic_random(dos, tres, quatre, cinc, poblacio_mig, linia, poblacio_mig)

'inicialitzar variables de trobar el maxim
        maxim_diab1 = -1E+99 ' maxim_sa '0
        maxim_diab2 = -1E+99 ' maxim_sa '0 '-100000000
        maxim_diab_mig = -1E+99 ' maxim_sa '0 ' -100000000
        
' trobar el maxim de cada un i fer per una banda resta amb el maxim sa
' per laltre comprobar els signes de la resta per saber les noves N1 i N2
    
    For kk = (24 * 60) To UBound(Diab_1) - 1 Step 1  ' he de començar quan comença a menjar
        'Maxim_sa = Max(Sa) -> bucle que recorri el vector i trii el maxim.
        If Diab_1(kk) > maxim_diab1 Then
            maxim_diab1 = Diab_1(kk)
        End If
        
        If Diab_2(kk) > maxim_diab2 Then
            maxim_diab2 = Diab_2(kk)
        End If
        
        If diab_mig(kk) > maxim_diab_mig Then
            maxim_diab_mig = diab_mig(kk)
        End If
        
        
    Next kk

        dif_1 = Abs(maxim_diab1 - maxim_diab2)
        'dif_2 = maxim_diab2 - maxim_sa
        'dif_mig = maxim_diab_mig - maxim_sa
        
'        Print maxim_1
'    Print maxim_2
'    Print maxim_mig
        
        If maxim_diab_mig > maxim_sa Then
            poblacio1 = poblacio_mig
        Else
            poblacio2 = poblacio_mig
        End If


        
       ' If maxim_diab1 > maxim_sa Then
       '     If maxim_diab_mig > maxim_sa Then
       '         poblacio1 = poblacio_mig
       '         dif_1 = dif_mig
       '     Else
       '         poblacio1 = poblacio1
       '         dif_1 = dif_1
       '     End If
       ' End If
   '
   '     If maxim_diab2 < maxim_sa Then
   '         If maxim_diab_mig < maxim_sa Then
   '             poblacio2 = poblacio_mig
   '             dif_2 = dif_mig
   '         Else
   '             poblacio2 = poblacio2
   '             dif_2 = dif_2
   '         End If
   '     End If
        
        
       ' Print dif_1
        'Print Abs(dif_2)

        
       ' Print maxim_diab1
       ' Print dif_1
        
       ' Print maxim_diab2
       ' Print dif_2
        
       ' Print maxim_diab_mig
       ' Print dif_mig
   Loop
    
            mm = FreeFile
        
        If Abs(dif_1) < Abs(dif_2) Then
         
            Open "E:\diabmaximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
            
            'Open "\\VBOXSVR\Compartit_Windows\diab\maximcelula_" & linia & "glucose_" & poblacio1 & ".txt" For Output As mm
            For Index = 0 To UBound(Diab_1) Step 1
                Print #mm, Index, Diab_1(Index)
            Next Index
        Else
        
            'Open "\\VBOXSVR\Compartit_Windows\diab\maximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
            
            Open "E:\diabmaximcelula_" & linia & "glucose_" & poblacio2 & ".txt" For Output As mm
            For Index = 0 To UBound(Diab_2) Step 1
                Print #mm, Index, Diab_2(Index)
            Next Index
        End If
          
        Close #mm
 'Loop
 'linia
    
    
    '- arreglar la poblacio
  
    'MsgBox (linia)
'Next linia


' a
'MsgBox ("Fi")
'id = Shell("E:\Codigo_Javier\Codigo_Eva\codi_minim")
'id = Shell("\\VBOXSVR\Compartit_Windows\Codigo_Javier\Codigo_Eva\codi_minim")
End
End Sub







Private Sub Command7_Click()
'Private Sub Command4_Click()


Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double
Dim maxim_sa As Double, maxima_diab1 As Double, maxim_diab2 As Double, maxim_diab_mig As Double
Dim K As Double, kk As Double
Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double
Dim Sa() As Variant
Dim Diab_1() As Variant
Dim Diab_2() As Variant
Dim diab_mig() As Variant

mm = FreeFile
Open "\\VBOXSVR\Compartit_Windows\dades_3.txt " For Output As mm
Close #mm

nn = FreeFile
myfile = "\\VBOXSVR\Compartit_Windows\diab\parametresatzar_3.txt"
Open myfile For Input As #nn

Sa = pacient_sa()
maxim_sa = 0
minim_sa = 1000

hmenjar = (24 + 6) * 60
Print hmenjar
deju = 126
hdeju = hmenjar - 1
Print hdeju

hyper = 200
hhyper = hmenjar + (2 * 60)
Print hhyper


linia = 0
Do While Not EOF(nn)
    linia = linia + 1

    Line Input #nn, textline 'llegeix la linia
    Print linia
   
    pp = InStr(1, textline, ",")
    dos = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    tres = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    quatre = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    cinc = Val(textline)
    
    
    poblacio1 = 1
    cont = 0
    dif_1 = 100
    dif_2 = 100
    poblacio2 = poblacio1
    Do
        poblacio2 = poblacio2 * 2
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        'maxim_diab2 = -1E+99
       
        maxim_diab2 = Diab_2(hhyper)
        'Print maxim_diab2
        deju_diab2 = Diab_2(hdeju)
        'Print deju_diab2
        If (maxim_diab2 < hyper) Then Exit Do
        DoEvents
        Form1.Caption = poblacio2 & " " & maxim_diab2
    Loop
    
    poblacio1 = Int(poblacio2 / 2)


Do While (Abs(dif_1) > 1)   ' condicio de sortida que depen de la diferencia entre pic diab i pic sa

    cont = cont + 1
    Form1.Caption = "celula: " & linia & " iteracio: " & cont
   ' Print cont
        poblacio_mig = Int((poblacio1 + poblacio2) / 2) ' calcula poblacio entre mig de la 1 i la 2
 'Print poblacio1
 
     
        Diab_1 = diabetic_random(dos, tres, quatre, cinc, poblacio1, linia, poblacio1)
        Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
        diab_mig = diabetic_random(dos, tres, quatre, cinc, poblacio_mig, linia, poblacio_mig)

'inicialitzar variables de trobar el maxim
        maxim_diab1 = Diab_1(hhyper)
        maxim_diab2 = Diab_2(hhyper)
        maxim_diab_mig = diab_mig(hhyper)
        
        deju1 = Diab_1(hdeju)
        deju2 = Diab_2(hdeju)
        dejumig = diab_mig(hdeju)
        
        
        
' trobar el maxim de cada un i fer per una banda resta amb el maxim sa
' per laltre comprobar els signes de la resta per saber les noves N1 i N2
    
    
        dif_1 = Abs(maxim_diab1 - maxim_diab2)
        dif_2 = Abs(deju1 - deju2)
       
        If (maxim_diab_mig > hyper) Then
            poblacio1 = poblacio_mig
        Else
            poblacio2 = poblacio_mig
        End If
   Loop
            
'Un cop optimitzada la N hiper- mirem si també s'ha optmitzat la de dejú
NOpt = Int(poblacio1 + poblacio2) / 2
Diab_1 = diabetic_random(dos, tres, quatre, cinc, NOpt, linia, NOpt)

If Diab_1(hdeju) < deju Then
        'Grabar parametres i glucoses
        minim = Min(Diab_1)
        mm = FreeFile
        Open "\\VBOXSVR\Compartit_Windows\dades_3.txt" For Append As mm
            Print #mm, dos, tres, quatre, cinc, NOpt, minim, Diab_1(hdeju), Diab_1(hhyper)
        Close #mm
Else
    'Cal seguir optimitzant el dejú
    poblacio2 = Int(NOpt * 2)
    poblacio1 = NOpt
    dif_1 = 1E+99
    Do While (Abs(dif_1) > 1)  ' Or (cont < 50) ' condicio de sortida que depen de la diferencia entre pic diab i pic sa
       ' Print cont
            poblacio_mig = Int((poblacio1 + poblacio2) / 2) ' calcula poblacio entre mig de la 1 i la 2
     'Print poblacio1
     
         
            Diab_1 = diabetic_random(dos, tres, quatre, cinc, poblacio1, linia, poblacio1)
            Diab_2 = diabetic_random(dos, tres, quatre, cinc, poblacio2, linia, poblacio2)
            diab_mig = diabetic_random(dos, tres, quatre, cinc, poblacio_mig, linia, poblacio_mig)
    
    'inicialitzar variables de trobar el maxim
            maxim_diab1 = Diab_1(hhyper)
            maxim_diab2 = Diab_2(hhyper)
            maxim_diab_mig = diab_mig(hhyper)
            
            deju1 = Diab_1(hdeju)
            deju2 = Diab_2(hdeju)
            dejumig = diab_mig(hdeju)
            
            
            
    ' trobar el maxim de cada un i fer per una banda resta amb el maxim sa
    ' per laltre comprobar els signes de la resta per saber les noves N1 i N2
        
        
            dif_1 = Abs(deju1 - deju2)
           
            If (dejumig > deju) Then
                poblacio1 = poblacio_mig
            Else
                poblacio2 = poblacio_mig
            End If
       Loop
       
       'Aquí tot ha d'estar optimitzat tot
       'Grabar parametres i glucoses
       NOpt = Int(poblacio1 + poblacio2) / 2
       Diab_1 = diabetic_random(dos, tres, quatre, cinc, NOpt, linia, NOpt)
       
       
        minim = 1E+99
        For x = hmenjar To UBound(Diab_1) - 1
            If Diab_1(x) < minim Then
                minim = Diab_1(x)
            End If
        Next x
        
        mm = FreeFile
        Open "\\VBOXSVR\Compartit_Windows\dades_3.txt " For Append As mm
            Print #mm, dos, tres, quatre, cinc, NOpt, minim, Diab_1(hdeju), Diab_1(hhyper)
        Close #mm
End If
            
    
 Loop

End Sub







Private Sub Command8_Click()



Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double
Dim maxim_sa As Double, maxima_diab1 As Double, maxim_diab2 As Double, maxim_diab_mig As Double
Dim K As Double, kk As Double
Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double
Dim Sa() As Variant
Dim Diab_1() As Variant
Dim Diab_2() As Variant
Dim diab_mig() As Variant

mm = FreeFile
Open "\\VBOXSVR\Compartit_Windows\diab\dades_minim_1.txt " For Output As mm
Close #mm


nn = FreeFile
myfile = "\\VBOXSVR\Compartit_Windows\diab\dades_1.txt"
Open myfile For Input As #nn



hmenjar = (24 + 6) * 60
'Print hmenjar
deju = 126
hdeju = hmenjar - 1
'Print hdeju

hyper = 200
hhyper = hmenjar + (2 * 60)
'Print hhyper


linia = 0
Do While Not EOF(nn)
    linia = linia + 1
    Print linia
'--- he dafegir llegir un ultim valor
   Line Input #nn, textline 'llegeix la linia
   
   
    pp = InStr(1, textline, ",")
    dos = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    tres = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    quatre = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    cinc = Val(Trim(Left(textline, pp - 1)))
    
    textline = Trim(Right(textline, Len(textline) - pp))
    poblacio = Val(Trim(Left(textline, pp - 1)))
    
    
    
    Diab = diabetic_random(dos, tres, quatre, cinc, poblacio, linia, poblacio)
        
' trobar valor minim
    minim = 1E+99
    For x = hmenjar To UBound(Diab) - 1
        If Diab(x) < minim Then
            minim = Diab(x)
        End If
    Next x
    
    mm = FreeFile
        Open "\\VBOXSVR\Compartit_Windows\dades_minim_1.txt" For Append As mm
            Print #mm, linia, poblacio, minim
        Close #mm
 Loop
            
    


End Sub


Private Sub Command9_Click()



Dim N As Double, poblacio As Double, sum_fit1 As Double, sum_fit2 As Double
Dim fitness1 As Double
Dim fitness2 As Double
Dim maxim_sa As Double, maxima_diab1 As Double, maxim_diab2 As Double, maxim_diab_mig As Double
Dim K As Double, kk As Double
Dim un As Double, dos As Double, tres As Double, quatre As Double, cinc As Double
Dim Sa() As Variant
Dim Diab_1() As Variant
Dim Diab_2() As Variant
Dim diab_mig() As Variant




nn = FreeFile
myfile = "\\VBOXSVR\Compartit_Windows\minimcelula3_analisi.txt"
Open myfile For Input As #nn



hmenjar = (24 + 6) * 60
'Print hmenjar
deju = 126
hdeju = hmenjar - 1
'Print hdeju

hyper = 200
hhyper = hmenjar + (2 * 60)
'Print hhyper


linia = 0
Do While Not EOF(nn)
    linia = linia + 1
    Print linia
'--- he dafegir llegir un ultim valor
   Line Input #nn, textline 'llegeix la linia
   
   
    pp = InStr(1, textline, ",")
    dos = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    tres = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    quatre = Val(Trim(Left(textline, pp - 1)))
    textline = Trim(Right(textline, Len(textline) - pp))
    
    pp = InStr(1, textline, ",")
    cinc = Val(Trim(Left(textline, pp - 1)))
    
    textline = Trim(Right(textline, Len(textline) - pp))
    poblacio = Val(Trim(Left(textline, pp - 1)))
    
    
    
    Diab = diabetic_random(dos, tres, quatre, cinc, poblacio, linia, poblacio)
        
' trobar valor minim
    minim = 1E+99
     mm = FreeFile
     Open "\\VBOXSVR\Compartit_Windows\celulaoptimitzada3" & linia & ".txt" For Output As mm
           
    For x = 1 To UBound(Diab) - 1
         Print #mm, x, Diab(x)
    Next x
    
        Close #mm
 Loop


'End Sub

End Sub

Private Sub Generador_Click()
' generar parametres al atzar
Randomize Timer
Dim variables(5, 1000) As Double
mm = FreeFile
Open "\\VBOXSVR\Compartit_Windows\diab\parametresatzar_2.txt" For Output As mm
'Open "E:\diab\parametresatzar_1000.txt" For Output As mm
For cont = 1 To 333
    variables(0, cont) = Rnd * 100
    variables(1, cont) = Rnd * 100
    variables(2, cont) = Rnd * 100
    variables(3, cont) = Rnd * 100
   ' variables(4, cont) = Rnd

    Print #mm, Str(variables(0, cont)), ",", Str(variables(1, cont)), ", ", Str(variables(2, cont)), " ,", Str(variables(3, cont)) ', " ,", Str(variables(4, cont))
          
        
Next cont
MsgBox ("Fi")
Close #mm
End Sub
