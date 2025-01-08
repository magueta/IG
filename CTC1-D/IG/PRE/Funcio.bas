Attribute VB_Name = "Funciones"
Function DL(Vs, VI, Ex1, Ex2, XX)
    Static M, N
    M = (Vs - VI) / (Ex2 - Ex1)
    N = VI - Ex1 * M
    DL = M * XX + N
End Function

Function DP(VExt As Single, VCent As Single, Ex1 _
                   As Single, Ex2 As Single, XX As Single) As Single
      Static AA As Single, BB As Single, GG As Single
      AA = 4 * (VExt - VCent) / ((Ex2 - Ex1) * (Ex2 - Ex1))
      BB = (VExt - VCent) * 2 / (Ex2 - Ex1) _
               - (3 * Ex2 + Ex1) / 2 * AA
      GG = VCent - (((Ex1 + Ex2)) * ((Ex1 + Ex2)) * AA / 4 _
              + ((Ex2 + Ex1) / 2 * BB))
      DP = AA * XX * XX + BB * XX + GG
End Function
Function Maximo(A, B)
    If A > B Then Maximo = A Else Maximo = B
End Function
