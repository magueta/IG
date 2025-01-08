Attribute VB_Name = "Rutinas"


Sub ObtenerE()
        ReDim Preserve Xi(2, Ndxi), Yi(2, Ndyi), _
                                Xs(2, Ndxs), Yd(2, Ndyd)
        ReDim Preserve Txi(3, Ndxi), Txs(3, Ndxs), _
                                Tyi(3, Ndyi), Tyd(3, Ndyd)
        ReDim Preserve Vxi(3, Ndxi), Vxs(3, Ndxs), _
                                Uyi(3, Ndyi), Uyd(3, Ndyd)
        ReDim Preserve Uxi(3, Ndxi), Uxs(3, Ndxs), _
                                Vyi(3, Ndyi), Vyd(3, Ndyd)
        For I = 1 To Ndxi
            Txi(1, I) = 0: Txi(2, I) = 0: Txi(3, I) = 0
            Vxi(1, I) = 2: Vxi(2, I) = 0: Vxi(3, I) = 0
            Uxi(1, I) = 2: Uxi(2, I) = 0: Uxi(3, I) = 0
        Next I
        For I = 1 To Ndxs
            Txs(1, I) = 0: Txs(2, I) = 0: Txs(3, I) = 0
            Vxs(1, I) = 2: Vxs(2, I) = 0: Vxs(3, I) = 0
            Uxs(1, I) = 2: Uxs(2, I) = 0: Uxs(3, I) = 0
        Next I
        For I = 1 To Ndyi
            Tyi(1, I) = 0: Tyi(2, I) = 0: Tyi(3, I) = 0
            Uyi(1, I) = 2: Uyi(2, I) = 0: Uyi(3, I) = 0
            Vyi(1, I) = 2: Vyi(2, I) = 0: Vyi(3, I) = 0
        Next I
        For I = 1 To Ndyd
            Tyd(1, I) = 0: Tyd(2, I) = 0: Tyd(3, I) = 0
            Uyd(1, I) = 2: Uyd(2, I) = 0: Uyd(3, I) = 0
            Vyd(1, I) = 2: Vyd(2, I) = 0: Vyd(3, I) = 0
        Next I
        
        Xi(1, 1) = 1
        Xi(2, 1) = L3 \ Ndxi
        For I = 2 To Ndxi
            Xi(1, I) = Xi(2, I - 1) + 1
            Xi(2, I) = Xi(1, I) + L3 \ Ndxi - 1
        Next I
        Xi(2, Ndxi) = L3
        
        Xs(1, 1) = 1
        Xs(2, 1) = L3 \ Ndxs
        For I = 2 To Ndxs
            Xs(1, I) = Xs(2, I - 1) + 1
            Xs(2, I) = Xs(1, I) + L3 \ Ndxs - 1
        Next I
        Xs(2, Ndxs) = L3
        
        Yi(1, 1) = 1
        Yi(2, 1) = M3 \ Ndyi
        For I = 2 To Ndyi
            Yi(1, I) = Yi(2, I - 1) + 1
            Yi(2, I) = Yi(1, I) + M3 \ Ndyi - 1
        Next I
        Yi(2, Ndyi) = M3
        
        Yd(1, 1) = 1
        Yd(2, 1) = M3 \ Ndyd
        For I = 2 To Ndyd
            Yd(1, I) = Yd(2, I - 1) + 1
            Yd(2, I) = Yd(1, I) + M3 \ Ndyd - 1
        Next I
        Yd(2, Ndyd) = M3
        
        For J = 0 To 3
            If FCF.Check2(J) = vbChecked Then
                Select Case J
                    Case 0
                        Xi(1, 1) = 1
                        Xi(2, 1) = LL(1) - 2
                        For I = 2 To Ndxi
                            Xi(1, I) = Xi(2, I - 1) + 1
                            Xi(2, I) = LL(I) - 2
                        Next I
                        Xi(2, Ndxi) = L3
                    Case 1
                        Yi(1, 1) = 1
                        Yi(2, 1) = MM(1) - 2
                        For I = 2 To Ndyi
                            Yi(1, I) = Yi(2, I - 1) + 1
                            Yi(2, I) = MM(I) - 2
                        Next I
                        Yi(2, Ndyi) = M3
                    Case 2
                        Xs(1, 1) = 1
                        Xs(2, 1) = LL(1) - 2
                        For I = 2 To Ndxs
                            Xs(1, I) = Xs(2, I - 1) + 1
                            Xs(2, I) = LL(I) - 2
                        Next I
                        Xs(2, Ndxs) = L3
                    Case 3
                        Yd(1, 1) = 1
                        Yd(2, 1) = MM(1) - 2
                        For I = 2 To Ndyd
                            Yd(1, I) = Yd(2, I - 1) + 1
                            Yd(2, I) = MM(I) - 2
                        Next I
                        Yd(2, Ndyd) = M3
                End Select
            End If
        Next J
End Sub

Sub Cadicio()
    L2 = L1 - 1
    M2 = M1 - 1
    L3 = L2 - 1
    M3 = M2 - 1
    X(L1) = Xl
    Y(M1) = Yl
    For I = 2 To L2
        X(I) = (XU(I + 1) + XU(I)) / 2
        XCV(I) = -XU(I) + XU(I + 1)
    Next I
    For J = 2 To M2
        Y(J) = (YV(J + 1) + YV(J)) / 2
        YCV(J) = -YV(J) + YV(J + 1)
    Next J
    For I = 2 To L1
        XDif(I) = X(I) - X(I - 1)
    Next I
    For J = 2 To M1
        YDif(J) = Y(J) - Y(J - 1)
    Next J
    
End Sub

Sub Expone(TCoor As Boolean, Bloq As Long, Pow As Single, TT As Long)
    Static Ll1 As Single, Mm1 As Single, Xl2 As Single, Yl2 As Single, _
           Ik As Long, Jk As Long, Frac As Single
    Select Case TCoor
        
        Case True
            Select Case TT
                Case 1
                    Pow = 1
                    XU(LL(Bloq - 1)) = Dx(0, Bloq)
                    XU(LL(Bloq)) = Dx(0, Bloq) + Dx(1, Bloq)
                    Ll1 = L(Bloq) / 2
                    Xl2 = Dx(1, Bloq) / 2
                    For Ik = 1 To Ll1
                        Frac = Xl2 * (Ik / Ll1) ^ Pow
                        XU(Ik + LL(Bloq - 1)) = XU(LL(Bloq - 1)) + Frac
                        XU(LL(Bloq) - Ik) = XU(LL(Bloq)) - Frac
                    Next Ik
                
                Case 2
                    XU(LL(Bloq - 1)) = Dx(0, Bloq)
                    XU(LL(Bloq)) = Dx(0, Bloq) + Dx(1, Bloq)
                    Ll1 = L(Bloq) / 2
                    Xl2 = Dx(1, Bloq) / 2
                    For Ik = 1 To Ll1
                        Frac = Xl2 * (Ik / Ll1) ^ Pow
                        XU(Ik + LL(Bloq - 1)) = XU(LL(Bloq - 1)) + Frac
                        XU(LL(Bloq) - Ik) = XU(LL(Bloq)) - Frac
                    Next Ik
                
                Case 3
                    XU(LL(Bloq - 1)) = Dx(0, Bloq)
                    XU(LL(Bloq)) = Dx(0, Bloq) + Dx(1, Bloq)
                    Ll1 = L(Bloq) / 2
                    Xl2 = Dx(1, Bloq) / 2
                    For Ik = 1 To Ll1
                        Frac = Xl2 * (Ik / Ll1) ^ (1 / Pow)
                        XU(Ik + LL(Bloq - 1)) = XU(LL(Bloq - 1)) + Frac
                        XU(LL(Bloq) - Ik) = XU(LL(Bloq)) - Frac
                    Next Ik
                
                Case 4
                    XU(LL(Bloq - 1)) = Dx(0, Bloq)
                    XU(LL(Bloq)) = Dx(0, Bloq) + Dx(1, Bloq)
                    Ll1 = L(Bloq)
                    Xl2 = Dx(1, Bloq)
                    For Ik = 1 To Ll1
                        Frac = Xl2 * (Ik / Ll1) ^ Pow
                        XU(Ik + LL(Bloq - 1)) = XU(LL(Bloq - 1)) + Frac
                    Next Ik
                
                Case 5
                    XU(LL(Bloq - 1)) = Dx(0, Bloq)
                    XU(LL(Bloq)) = Dx(0, Bloq) + Dx(1, Bloq)
                    Ll1 = L(Bloq)
                    Xl2 = Dx(1, Bloq)
                    For Ik = 1 To Ll1
                        Frac = Xl2 * (Ik / Ll1) ^ Pow
                        XU(LL(Bloq) - Ik) = XU(LL(Bloq)) - Frac
                    Next Ik
            
            End Select
        Case False
            Select Case TT
                
                Case 1
                    Pow = 1
                    YV(MM(Bloq - 1)) = Dy(0, Bloq)
                    YV(MM(Bloq)) = Dy(0, Bloq) + Dy(1, Bloq)
                    Mm1 = M(Bloq) / 2
                    Yl2 = Dy(1, Bloq) / 2
                    For Jk = 1 To Mm1
                        Frac = Yl2 * (Jk / Mm1) ^ Pow
                        YV(Jk + MM(Bloq - 1)) = YV(MM(Bloq - 1)) + Frac
                        YV(MM(Bloq) - Jk) = YV(MM(Bloq)) - Frac
                    Next Jk
                
                Case 2
                    YV(MM(Bloq - 1)) = Dy(0, Bloq)
                    YV(MM(Bloq)) = Dy(0, Bloq) + Dy(1, Bloq)
                    Mm1 = M(Bloq) / 2
                    Yl2 = Dy(1, Bloq) / 2
                    For Jk = 1 To Mm1
                        Frac = Yl2 * (Jk / Mm1) ^ Pow
                        YV(Jk + MM(Bloq - 1)) = YV(MM(Bloq - 1)) + Frac
                        YV(MM(Bloq) - Jk) = YV(MM(Bloq)) - Frac
                    Next Jk
                
                Case 3
                    YV(MM(Bloq - 1)) = Dy(0, Bloq)
                    YV(MM(Bloq)) = Dy(0, Bloq) + Dy(1, Bloq)
                    Mm1 = M(Bloq) / 2
                    Yl2 = Dy(1, Bloq) / 2
                    For Jk = 1 To Mm1
                        Frac = Yl2 * (Jk / Mm1) ^ (1 / Pow)
                        YV(Jk + MM(Bloq - 1)) = YV(MM(Bloq - 1)) + Frac
                        YV(MM(Bloq) - Jk) = YV(MM(Bloq)) - Frac
                    Next Jk
                
                Case 4
                    YV(MM(Bloq - 1)) = Dy(0, Bloq)
                    YV(MM(Bloq)) = Dy(0, Bloq) + Dy(1, Bloq)
                    Mm1 = M(Bloq)
                    Yl2 = Dy(1, Bloq)
                    For Jk = 1 To Mm1
                        Frac = Yl2 * (Jk / Mm1) ^ Pow
                        YV(Jk + MM(Bloq - 1)) = YV(MM(Bloq - 1)) + Frac
                    Next Jk
                
                Case 5
                    YV(MM(Bloq - 1)) = Dy(0, Bloq)
                    YV(MM(Bloq)) = Dy(0, Bloq) + Dy(1, Bloq)
                    Mm1 = M(Bloq)
                    Yl2 = Dy(1, Bloq)
                    For Jk = 1 To Mm1
                        Frac = Yl2 * (Jk / Mm1) ^ Pow
                        YV(MM(Bloq) - Jk) = YV(MM(Bloq)) - Frac
                    Next Jk
            
            End Select
    
    End Select
End Sub

Sub ObtenerS()
    Nss = 2 * Ns + 2
    ReDim Sx(Nss), Sy(Nss)
    Sx(1) = 0
    Sy(1) = 0
    Sx(Nss) = Xl
    Sy(Nss) = Yl
    For K = 2 To Nss - 2 Step 2
        Sx(K) = S(0, K \ 2)
        Sx(K + 1) = S(2, K \ 2)
        Sy(K) = S(1, K \ 2)
        Sy(K + 1) = S(3, K \ 2)
        
    Next K
    OrdenaS Sx
    OrdenaS Sy
    Ndx = 0
    Ndy = 0
    For K = 1 To Nss - 1
        If Sx(K + 1) <> Sx(K) Then
            Ndx = Ndx + 1
            ReDim Preserve Dx(1, Ndx)
            Dx(1, Ndx) = Sx(K + 1) - Sx(K)
            Dx(0, Ndx) = Sx(K)
        End If
    Next K
    For K = 1 To Nss - 1
        If Sy(K + 1) <> Sy(K) Then
            Ndy = Ndy + 1
            ReDim Preserve Dy(1, Ndy)
            Dy(1, Ndy) = Sy(K + 1) - Sy(K)
            Dy(0, Ndy) = Sy(K)
        End If
    Next K

End Sub

'**********************************************
'Ordenacion por el metodo Shell.
'Hecha por Juan C. Magueta B.
'
'Ass() es la "lista" a ordenar.
'
'
'Metodo Shell de ordenación fue descubierto por Donald Shell
'hace mas de treinta años, nadie sabe como ordena tan rapido.
'
'Este metodo se debe escojer cuando una lista sea medianamente grande.
'
'*********************************************************************
Sub OrdenaS(Ass)
    Dim NumeroDeEntradas As Long, Increm As Long, JJ As Long, Temp As Single
    NumeroDeEntradas = UBound(Ass)
    Increm = NumeroDeEntradas \ 2
    Do Until Increm < 1
        For II = Increm + 1 To NumeroDeEntradas                 '
            Temp = Ass(II)                                      '
            For JJ = II - Increm To 1 Step -Increm              '
                If Temp >= Ass(JJ) Then Exit For                '
                Ass(JJ + Increm) = Ass(JJ)                      '
            Next JJ                                             '
            Ass(JJ + Increm) = Temp                             '
        Next II                                                 '
        Increm = Increm \ 2                                     '
    Loop                                                        ' Fin del Bucle de ordenación
    
End Sub

Sub VCUniforme(TCoor As Boolean, Bloq)
    Select Case TCoor
        
        Case True
            XU(LL(Bloq - 1)) = Dx(0, Bloq)
            XU(LL(Bloq)) = Dx(0, Bloq) + Dx(1, Bloq)
            For I = LL(Bloq - 1) + 1 To LL(Bloq) - 1
                XU(I) = (I - LL(Bloq - 1)) / (L(Bloq)) * Dx(1, Bloq) + XU(LL(Bloq - 1))
            Next I
        
        Case False
            YV(MM(Bloq - 1)) = Dy(0, Bloq)
            YV(MM(Bloq)) = Dy(0, Bloq) + Dy(1, Bloq)
            For I = MM(Bloq - 1) + 1 To MM(Bloq) - 1
                YV(I) = (I - MM(Bloq - 1)) / (M(Bloq)) * Dy(1, Bloq) + YV(MM(Bloq - 1))
            Next I
    End Select

End Sub


