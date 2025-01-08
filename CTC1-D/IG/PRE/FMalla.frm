VERSION 5.00
Begin VB.Form FMalla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mallado"
   ClientHeight    =   4365
   ClientLeft      =   1335
   ClientTop       =   1395
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Configuración de bloques"
      Height          =   4155
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3135
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de malla"
         Height          =   2145
         Left            =   60
         TabIndex        =   8
         Top             =   1920
         Width           =   3015
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1980
            TabIndex        =   14
            Top             =   1680
            Width           =   795
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Agrupados a la derecha"
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   1200
            Width           =   2715
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Agrupados a la izquierda"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   915
            Width           =   2715
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Agrupados al centro"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   690
            Width           =   2715
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Agrupados a los extremos"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   465
            Width           =   2715
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Volumenes de control uniforme"
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   2715
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Densidad de V.C. en la dirección escogida"
            Height          =   435
            Left            =   120
            TabIndex        =   16
            Top             =   1620
            Width           =   1755
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   1635
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   1860
         TabIndex        =   5
         Top             =   900
         Value           =   1
         Width           =   195
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bloques en la dirreción y"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   540
         Width           =   2355
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bloques en la dirreción x"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Numeros de V.C. del bloque"
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2835
      Left            =   3360
      ScaleHeight     =   2775
      ScaleWidth      =   2775
      TabIndex        =   1
      Top             =   120
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seguir"
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   3600
      Width           =   795
   End
End
Attribute VB_Name = "FMalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CEscala()
    Mediaescala = (EscalaXY * 1.1) / 2
    With Picture1
        .ScaleHeight = -2 * Mediaescala
        .ScaleWidth = 2 * Mediaescala
        .ScaleLeft = -Mediaescala + Xl / 2
        .ScaleTop = Yl / 2 + Mediaescala
    End With
End Sub

Private Sub Command1_Click()
    M3 = 0
    L3 = 0
    LL(1) = L(1)
    MM(1) = M(1)
    LL(0) = 2
    MM(0) = 2
    For K = 1 To Ndx
        L3 = L3 + L(K)
    Next K
    
    For K = 1 To Ndy
        M3 = M3 + M(K)
    Next K
    
    For K = 1 To Ndx
        LL(K) = LL(K - 1) + L(K)
    Next K
    
    For K = 1 To Ndy
        MM(K) = MM(K - 1) + M(K)
    Next K
    
    L2 = L3 + 1
    M2 = M3 + 1
    L1 = L2 + 1
    M1 = M2 + 1
    
    ReDim Preserve XU(L1), YV(M1), X(L1), Y(M1), XDif(L1), YDif(M1), XCV(L1), YCV(M1)
    
    For I = 1 To Ndx
        Call Expone(True, I, PowX(I), Tmx(I))
    Next I
    
    For I = 1 To Ndy
        Call Expone(False, I, PowY(I), Tmy(I))
    Next I
    
    Call Cadicio
    Me.Hide
End Sub

Private Sub Form_Activate()
    CEscala
    ReDim Preserve Tmx(Ndx), Tmy(Ndy), L(Ndx), M(Ndy), PowX(Ndx), _
    PowY(Ndy), LL(Ndx), MM(Ndy)
    ReDim BloqX(Ndx), BloqY(Ndy)
    For K = 1 To Ndx
        BloqX(K) = "Bloque X " & Str(K)
    Next K
        For K = 1 To Ndy
        BloqY(K) = "Bloque Y " & Str(K)
    Next K
    Picture1.Cls
    VScroll1.Max = 1
    VScroll1.Min = Ndx
    VScroll1 = 1
    Old = Picture1.FillStyle
    OldC = Picture1.FillColor
    Picture1.FillStyle = 3
    Picture1.FillColor = QBColor(10)
    Text1 = BloqX(VScroll1)
    Picture1.Line (Dx(1, VScroll1) + Dx(0, VScroll1), 0)-(Dx(0, VScroll1), Yl), , B
    Picture1.FillStyle = Old
    Picture1.FillColor = OldC
    Picture1.Line (0, 0)-(Xl, Yl), , B
    Option1 = True
    If Ns > 0 Then
        For K = 1 To Ns
            Picture1.Line (S(0, K), S(1, K))-(S(2, K), S(3, K)), , B
            Picture1.CurrentX = S(0, K) / 2 + S(2, K) / 2
            Picture1.CurrentY = S(1, K) / 2 + S(3, K) / 2
        Next K
    End If
  
    
    
End Sub

Private Sub Option1_Click()
Picture1.Cls
VScroll1.Max = 1
VScroll1.Min = Ndx
VScroll1 = 1
Option3.Caption = "Volumenes de control uniforme"
Option4.Caption = "Agrupados a los extremos"
Option5.Caption = "Agrupados al centro"
Option6.Caption = "Agrupados a la izquierda"
Option7.Caption = "Agrupados a la derecha"
Old = Picture1.FillStyle
OldC = Picture1.FillColor
Picture1.FillStyle = 3
Picture1.FillColor = QBColor(10)
Text1 = BloqX(VScroll1)
Text2 = L(VScroll1)
Text3 = PowX(VScroll1)
Select Case Tmx(VScroll1)
    Case 1
        Option3 = True
    Case 2
        Option4 = True
    Case 3
        Option5 = True
    Case 4
        Option6 = True
    Case 5
        Option7 = True
    Case Else
        Tmx(VScroll1) = 1
        Option3 = True
        PowX(VScroll1) = 1
End Select
Picture1.Line (Dx(1, VScroll1) + Dx(0, VScroll1), 0)-(Dx(0, VScroll1), Yl), , B
Picture1.FillStyle = Old
Picture1.FillColor = OldC
Picture1.Line (0, 0)-(Xl, Yl), , B
     If Ns > 0 Then
        For K = 1 To Ns
            Picture1.Line (S(0, K), S(1, K))-(S(2, K), S(3, K)), , B
            Picture1.CurrentX = S(0, K) / 2 + S(2, K) / 2
            Picture1.CurrentY = S(1, K) / 2 + S(3, K) / 2
        Next K
    End If
End Sub


Private Sub Option2_Click()
Picture1.Cls
VScroll1.Max = 1
VScroll1.Min = Ndy
VScroll1 = 1
Option3.Caption = "Volumenes de control uniforme"
Option4.Caption = "Agrupados a los extremos"
Option5.Caption = "Agrupados al centro"
Option6.Caption = "Agrupados abajo"
Option7.Caption = "Agrupados ariba"
Old = Picture1.FillStyle
OldC = Picture1.FillColor
Picture1.FillStyle = 2
Picture1.FillColor = QBColor(10)
Text1 = BloqY(VScroll1)
Text2 = M(VScroll1)
Text3 = PowY(VScroll1)
Picture1.Line (0, Dy(1, VScroll1) + Dy(0, VScroll1))-(Xl, Dy(0, VScroll1)), , B
Picture1.FillStyle = Old
Picture1.FillColor = OldC
Picture1.Line (0, 0)-(Xl, Yl), , B
Select Case Tmy(VScroll1)
    Case 1
        Option3 = True
    Case 2
        Option4 = True
    Case 3
        Option5 = True
    Case 4
        Option6 = True
    Case 5
        Option7 = True
    Case Else
        Tmy(VScroll1) = 1
        Option3 = True
        PowY(VScroll1) = 1
End Select

If Ns > 0 Then
    For K = 1 To Ns
        Picture1.Line (S(0, K), S(1, K))-(S(2, K), S(3, K)), , B
        Picture1.CurrentX = S(0, K) / 2 + S(2, K) / 2
        Picture1.CurrentY = S(1, K) / 2 + S(3, K) / 2
    Next K
End If

End Sub


Private Sub Option3_Click()
If Option1 Then
    Tmx(VScroll1) = 1
Else
    Tmy(VScroll1) = 1
End If
End Sub

Private Sub Option4_Click()
If Option1 Then
    Tmx(VScroll1) = 2
Else
    Tmy(VScroll1) = 2
End If

End Sub


Private Sub Option5_Click()
If Option1 Then
    Tmx(VScroll1) = 3
Else
    Tmy(VScroll1) = 3
End If

End Sub


Private Sub Option6_Click()
If Option1 Then
    Tmx(VScroll1) = 4
Else
    Tmy(VScroll1) = 4
End If

End Sub


Private Sub Option7_Click()
If Option1 Then
    Tmx(VScroll1) = 5
Else
    Tmy(VScroll1) = 5
End If

End Sub


Private Sub Text2_Change()
If Option1 Then
    L(VScroll1) = Val(Text2)
Else
    M(VScroll1) = Val(Text2)
End If

End Sub

Private Sub Text3_Change()
If Option1 Then
    PowX(VScroll1) = Val(Text3)
Else
    PowY(VScroll1) = Val(Text3)
End If
End Sub


Private Sub VScroll1_Change()
Picture1.Cls
Old = Picture1.FillStyle
OldC = Picture1.FillColor
Picture1.FillColor = QBColor(10)
Select Case Option1.Value
    Case True
        Text1 = BloqX(VScroll1)
        Text2 = L(VScroll1)
        Select Case Tmx(VScroll1)
            Case 1
                Option3 = True
            Case 2
                Option4 = True
            Case 3
                Option5 = True
            Case 4
                Option6 = True
            Case 5
                Option7 = True
            Case Else
                Tmx(VScroll1) = 1
                Option3 = True
                PowX(VScroll1) = 1
        End Select
        Text3 = PowX(VScroll1)
        Picture1.FillStyle = 3
        Picture1.Line (Dx(1, VScroll1) + Dx(0, VScroll1), 0)-(Dx(0, VScroll1), Yl), , B
        Picture1.FillStyle = Old
        Picture1.FillColor = OldC
    Case False
        Text1 = BloqY(VScroll1)
        Picture1.FillStyle = 2
        Picture1.Line (0, Dy(1, VScroll1) + Dy(0, VScroll1))-(Xl, Dy(0, VScroll1)), , B
        Picture1.FillStyle = Old
        Picture1.FillColor = OldC
        Text2 = M(VScroll1)
        Select Case Tmy(VScroll1)
            Case 1
                Option3 = True
            Case 2
                Option4 = True
            Case 3
                Option5 = True
            Case 4
                Option6 = True
            Case 5
                Option7 = True
            Case Else
                Tmy(VScroll1) = 1
                Option3 = True
                PowY(VScroll1) = 1
        End Select
        Text3 = PowY(VScroll1)
End Select
 Picture1.Line (0, 0)-(Xl, Yl), , B
 If Ns > 0 Then
    For K = 1 To Ns
        
        Picture1.Line (S(0, K), S(1, K))-(S(2, K), S(3, K)), , B
        Picture1.CurrentX = S(0, K) / 2 + S(2, K) / 2
        Picture1.CurrentY = S(1, K) / 2 + S(3, K) / 2
    Next K
End If
    
End Sub


