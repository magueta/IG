VERSION 5.00
Begin VB.Form Fver 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre-Procesador"
   ClientHeight    =   6180
   ClientLeft      =   1035
   ClientTop       =   435
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6180
   ScaleWidth      =   7860
   Begin VB.CommandButton Command9 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Resolver"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Parametros relacionados"
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   " Propiedades        del Medio"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "  Condiciones       de fronteras"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Mallado"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "  Editar Sólidos"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar Sólidos"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   5985
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   60
      Width           =   6015
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   1095
      Left            =   6360
      TabIndex        =   8
      Top             =   3300
      Width           =   1215
   End
End
Attribute VB_Name = "Fver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub CEscala()
    EscalaXY = Xl
    If Yl > Xl Then
        EscalaXY = Yl
    End If
    Mediaescala = (EscalaXY * 1.1) / 2
    With Picture1
        .ScaleHeight = -2 * Mediaescala
        .ScaleWidth = 2 * Mediaescala
        .ScaleLeft = -Mediaescala + Xl / 2
        .ScaleTop = Yl / 2 + Mediaescala
    End With
End Sub

Sub DiDominio()
    Picture1.Cls
    Picture1.Line (0, 0)-(Xl, Yl), , B
End Sub
Sub DiMalla()
    For I = 1 To L1
         For J = 1 To M1
             Picture1.PSet (X(I), Y(J)), QBColor(12)
         Next J
     Next I
     For I = 2 To L2
         For J = 2 To M2
             Picture1.Line (XU(I), YV(J))-(XU(I), YV(J + 1)), QBColor(8)
             Picture1.Line (XU(I), YV(J))-(XU(I + 1), YV(J)), QBColor(8)
         Next J
     Next I
    Picture1.Line (XU(L1), YV(M1))-(XU(L1), YV(2)), QBColor(8)
    Picture1.Line (XU(2), YV(M1))-(XU(L1), YV(M1)), QBColor(8)
End Sub

Sub DiSolido()
    If Ns > 0 Then
        For K = 1 To Ns
            Picture1.Line (S(0, K), S(1, K))-(S(2, K), S(3, K)), (vbMagenta), B
            Picture1.CurrentX = S(0, K) / 2 + S(2, K) / 2
            Picture1.CurrentY = S(1, K) / 2 + S(3, K) / 2
            Picture1.Print K
        Next K
    End If
End Sub

Sub Ocgxykr()
'Optener k
For J = 1 To M1
    For I = 1 To L1
        RHO(I, J) = Pro(0)
        rK(I, J) = Pro(1)
        Be(I, J) = Pro(2)
        CP(I, J) = Pro(3)
        GBX(I, J) = Gx * Pro(4)
        GBY(I, J) = Gy * Pro(4)
        For K = 1 To Ns
            If (S(0, K) <= X(I) And S(2, K) >= X(I)) And (S(1, K) <= Y(J) And S(3, K) >= Y(J)) _
            Then
                RHO(I, J) = S(5, K)
                rK(I, J) = S(4, K)
                Be(I, J) = 1E+38
                CP(I, J) = S(6, K)
            End If
        Next K
    Next I
Next J




End Sub

Sub OTFE()
ReDim TSp(L1, M1), TSc(L1, M1)
    For K = 1 To Ndx * Ndy
        Ji = (K - 1) \ Ndx + 1
        II = K - (Ji - 1) * Ndx
    For J = MM(Ji - 1) To MM(Ji) - 1
    For I = LL(II - 1) To LL(II) - 1
        TSp(I, J) = TFE(2, K)
        TSc(I, J) = TFE(1, K)
    Next I
    Next J
    Next K
End Sub

Sub OTUV()
For J = 2 To M2
    For I = 2 To L2
        T(I, J) = VI(0)
        If FU.Check1.Value = vbChecked Then
            U(I, J) = VI(1)
        Else
            U(I, J) = 0
        End If
        
        If FU.Check1.Value = vbChecked Then
            V(I, J) = VI(2)
        Else
            V(I, J) = 0
        End If
        
        For K = 1 To Ns
            If (S(0, K) <= X(I) And S(2, K) >= X(I)) And (S(1, K) <= Y(J) And S(3, K) >= Y(J)) _
                Then
                V(I + 1, J) = 0
                U(I, J + 1) = 0
                U(I, J) = 0
                V(I, J) = 0
            End If
        Next K
    Next I
Next J
For K = 1 To Ndxi
    
    If Txi(1, K) = 1 Then
        For I = Xi(1, K) + 1 To Xi(2, K) + 1
            T(I, 1) = Txi(2, K)
        Next I
    End If
    If Txi(1, K) = 4 Then
        For I = Xi(1, K) + 1 To Xi(2, K) + 1
            T(I, 1) = DL(Txi(3, K), Txi(2, K), X(Xi(1, K) + 1), X(Xi(2, K) + 1), X(I))
        Next I
    End If

    If Vxi(1, K) = 1 Then
        For I = Xi(1, K) + 1 To Xi(2, K) + 1
            V(I, 2) = DP(Vxi(2, K), Vxi(3, K), X(Xi(1, K) + 1), X(Xi(2, K) + 1), X(I))
        Next I
    End If
    If Uxi(1, K) = 2 Then
        For I = Xi(1, K) + 1 To Xi(2, K) + 1
            U(I, 1) = Uxi(2, K)
        Next I
    End If
    If Vxi(1, K) = 2 Then
        For I = Xi(1, K) + 1 To Xi(2, K) + 1
            V(I, 2) = Vxi(2, K)
        Next I
    End If
Next K


For K = 1 To Ndxs
    If Txs(1, K) = 1 Then
        For I = Xs(1, K) + 1 To Xs(2, K) + 1
            T(I, M1) = Txs(2, K)
        Next I
    End If
    If Txs(1, K) = 4 Then
        For I = Xs(1, K) + 1 To Xs(2, K) + 1
            T(I, M1) = DL(Txs(3, K), Txs(2, K), X(Xs(1, K) + 1), X(Xs(2, K) + 1), X(I))
        Next I
    End If
    
    If Uxs(1, K) = 2 Then
        For I = Xs(1, K) + 1 To Xs(2, K) + 1
            U(I, M1) = Uxs(2, K)
        Next I
    End If
        If Vxs(1, K) = 1 Then
        For I = Xs(1, K) + 1 To Xs(2, K) + 1
            V(I, M1) = DP(Vxs(2, K), Vxs(3, K), X(Xs(1, K) + 1), X(Xs(2, K) + 1), X(I))
        Next I
    End If
    
    If Vxs(1, K) = 2 Then
        For I = Xs(1, K) + 1 To Xs(2, K) + 1
            V(I, M1) = Vxs(2, K)
        Next I
    End If
Next K

For K = 1 To Ndyi
    If Tyi(1, K) = 1 Then
        For I = Yi(1, K) + 1 To Yi(2, K) + 1
            T(1, I) = Tyi(2, K)
        Next I
    End If
    If Tyi(1, K) = 4 Then
        For I = Yi(1, K) + 1 To Yi(2, K) + 1
            T(1, I) = DL(Tyi(3, K), Tyi(2, K), Y(Yi(1, K) + 1), Y(Yi(2, K) + 1), Y(I))
        Next I
    End If
    If Uyi(1, K) = 1 Then
        For I = Yi(1, K) + 1 To Yi(2, K) + 1
            U(2, I) = DP(Uyi(2, K), Uyi(3, K), Y(Yi(1, K) + 1), Y(Yi(2, K) + 1), Y(I))
        Next I
    End If

    If Uyi(1, K) = 2 Then
        For I = Yi(1, K) + 1 To Yi(2, K) + 1
            U(2, I) = Uyi(2, K)
        Next I
    End If
    If Vyi(1, K) = 2 Then
        For I = Yi(1, K) + 1 To Yi(2, K) + 1
            V(1, I) = Vyi(2, K)
        Next I
    End If
Next K
For K = 1 To Ndyd
    If Tyd(1, K) = 1 Then
        For I = Yd(1, K) + 1 To Yd(2, K) + 1
            T(L1, I) = Tyd(2, K)
        Next I
    End If
    If Tyd(1, K) = 4 Then
        For I = Yd(1, K) + 1 To Yd(2, K) + 1
            T(L1, I) = DL(Tyd(3, K), Tyd(2, K), Y(Yd(1, K) + 1), Y(Yd(2, K) + 1), Y(I))
        Next I
    End If
    
    If Uyd(1, K) = 1 Then
        For I = Yd(1, K) + 1 To Yd(2, K) + 1
            U(L1, I) = DP(Uyd(2, K), Uyd(3, K), Y(Yd(1, K) + 1), Y(Yd(2, K) + 1), Y(I))
        Next I
    End If
    If Uyd(1, K) = 2 Then
        For I = Yd(1, K) + 1 To Yd(2, K) + 1
            U(L1, I) = Uyd(2, K)
        Next I
    End If
    If Vyd(1, K) = 2 Then
        For I = Yd(1, K) + 1 To Yd(2, K) + 1
            V(L1, I) = Vyd(2, K)
        Next I
    End If
Next K
End Sub

Private Sub Command1_Click()
    Fsolido.Show (1)
    DiDominio
    DiSolido
End Sub

Private Sub Command2_Click()

FBorrar.Show (1)
DiDominio
If M1 > 2 And L1 > 2 Then
    DiMalla
End If
DiSolido


End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Command4_Click()
    ObtenerS
    Dim K As Integer
    FMalla.Show (1)
    DiDominio
    DiMalla
    DiSolido
End Sub

Private Sub Command5_Click()
FCF.Show (1)

End Sub

Private Sub Command6_Click()
FPR.Show (1)

End Sub

Private Sub Command7_Click()
Fpro.Show (1)
End Sub

Private Sub Command8_Click()
ReDim T(L1, M1), U(L1, M1), V(L1, M1), RHO(L1, M1)
ReDim rK(L1, M1), Be(L1, M1), CP(L1, M1), GBX(L1, M1), GBY(L1, M1)
Static Solve As String
Call OTUV
Call Ocgxykr
Call OTFE
Open "solido.el" For Output As #1
Print #1, Ns
For I = 1 To Ns
    Print #1, S(o, I), S(1, I), S(2, I), S(3, I), S(4, I), S(5, I), S(6, I)
Next I
Print #1, Pro(1)

Close #1
Open "geom.el" For Output As #1
Print #1, Mode
Print #1, Xl, Yl, R1
Print #1, L1, M1
Solve = ""
If FU.Check1 = vbChecked Then Solve = "T " Else Solve = "F "
If FV.Check1 = vbChecked Then Solve = Solve & "T " _
                                        Else Solve = Solve & "F "
If FE.Check1 = vbChecked Then Solve = Solve & "T" _
                                        Else Solve = Solve & "F"
Print #1, Solve

For I = 2 To L1
    Print #1, XU(I)
Next I
For J = 2 To M1
    Print #1, YV(J)
Next J
Print #1, Simetria
Close #1

Open "param.el" For Output As #1
Print #1, "F"
Print #1, NTime(0)
Print #1, NTime(2), NTime(3), NTime(1), NTime(4)
Print #1, Relax(1), Relax(2), Relax(0), Relax(3)
Print #1, 100
Close #1


Open "t.el" For Output As #1
For J = 1 To M1
    For I = 1 To L1
        Print #1, T(I, J)
    Next I
Next J

Close #1
Open "v.el" For Output As #1
For J = 2 To M1
    For I = 1 To L1
        Print #1, V(I, J)
    Next I
Next J

Close #1
Open "u.el" For Output As #1
For J = 1 To M1
    For I = 2 To L1
        Print #1, U(I, J)
    Next I
Next J
Close #1
Open "p.el" For Output As #1
For J = 1 To M1
    For I = 1 To L1
        Print #1, 0
    Next I
Next J
Close #1
Open "gama.el" For Output As #1
For J = 1 To M1
    For I = 1 To L1
        Print #1, rK(I, J), CP(I, J)
    Next I
Next J
For J = 1 To M1
    For I = 1 To L1
        Print #1, Be(I, J)
    Next I
Next J
For J = 1 To M1
    For I = 1 To L1
        Print #1, GBX(I, J)
    Next I
Next J
For J = 1 To M1
    For I = 1 To L1
        Print #1, GBY(I, J)
    Next I
Next J

Close #1
Open "rho.el" For Output As #1
For J = 1 To M1
    For I = 1 To L1
        Print #1, RHO(I, J)
    Next I
Next J
Close #1
Open "termi.el" For Output As #1
For J = 1 To M1
    For I = 1 To L1
        Print #1, TSc(I, J), TSp(I, J)
    Next I
Next J

Close #1
Open "contor.el" For Output As #1
    Print #1, Ndxi, Ndyi, Ndxs, Ndyd
    For J = 1 To Ndxi
        For I = 1 To 2
            Print #1, Xi(I, J)
        Next I
        Select Case Txi(1, J)
            Case 0
                Print #1, 1
                Print #1, 0
                Print #1, 0
            Case 1
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 2
                Print #1, 1
                Print #1, Txi(2, J)
                Print #1, 0
            Case 4
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 3
                Print #1, 4
                Print #1, Txi(2, J)
                Print #1, Txi(3, J)
        End Select
        If Vxi(1, J) = 0 Then Print #1, 1 Else Print #1, 0
    Next J
    For J = 1 To Ndyi
        For I = 1 To 2
            Print #1, Yi(I, J)
        Next I
        Select Case Tyi(1, J)
            Case 0
                Print #1, 1
                Print #1, 0
                Print #1, 0
            Case 1
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 2
                Print #1, 1
                Print #1, Tyi(2, J)
                Print #1, 0
            Case 4
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 3
                Print #1, 4
                Print #1, Tyi(2, J)
                Print #1, Tyi(3, J)
        End Select
        If Uyi(1, J) = 0 Then Print #1, 1 Else Print #1, 0
    Next J
    For J = 1 To Ndxs
        For I = 1 To 2
            Print #1, Xs(I, J)
        Next I
        Select Case Txs(1, J)
            Case 0
                Print #1, 1
                Print #1, 0
                Print #1, 0
            Case 1
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 2
                Print #1, 1
                Print #1, Txs(2, J)
                Print #1, 0
            Case 4
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 3
                Print #1, 4
                Print #1, Txs(2, J)
                Print #1, Txs(3, J)
        End Select
        If Vxs(1, J) = 0 Then Print #1, 1 Else Print #1, 0
    Next J
    For J = 1 To Ndyd
        For I = 1 To 2
            Print #1, Yd(I, J)
        Next I
        Select Case Tyd(1, J)
            Case 0
                Print #1, 1
                Print #1, 0
                Print #1, 0
            Case 1
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 2
                Print #1, 1
                Print #1, Tyd(2, J)
                Print #1, 0
            Case 4
                Print #1, 0
                Print #1, 0
                Print #1, 0
            Case 3
                Print #1, 4
                Print #1, Tyd(2, J)
                Print #1, Tyd(3, J)
        End Select
        If Uyd(1, J) = 0 Then Print #1, 1 Else Print #1, 0
    Next J
Close #1

id = Shell("Igs.exe", vbNormalFocus)


End Sub
Private Sub Command9_Click()
q = Shell("pre", vbNormalFocus)
End
End Sub


Private Sub Form_Activate()
Label1 = "   Xl = " & Xl & "    " & Chr$(13)
Label1 = Label1 & "    Yl = " & Yl & "    " & Chr$(13)
If Mode = 2 Then Label1 = Label1 & "    R(1) = " & R1 & "    " & Chr$(13)
Label1 = Label1 & "    l1 = " & L1 & "    " & Chr$(13)
Label1 = Label1 & "    M1 = " & M1 & "    "

End Sub

Private Sub Form_Load()
    For kk = 0 To 3
        Pro(kk) = 0.1
    Next kk
    
    CEscala
    DiDominio
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


