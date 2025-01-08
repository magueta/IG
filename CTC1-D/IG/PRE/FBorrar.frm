VERSION 5.00
Begin VB.Form FBorrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Sólidos"
   ClientHeight    =   4935
   ClientLeft      =   540
   ClientTop       =   1050
   ClientWidth     =   8340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4463
      TabIndex        =   20
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   2663
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   960
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   4260
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   4260
      TabIndex        =   10
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   5820
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   5820
      TabIndex        =   8
      Top             =   780
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      Height          =   1695
      Left            =   3600
      TabIndex        =   1
      Top             =   1800
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   1800
         TabIndex        =   4
         Text            =   "1"
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1800
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1800
         TabIndex        =   2
         Text            =   "1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Conductividad"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Densidad"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Cp"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1260
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   480
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "X1"
      Height          =   195
      Left            =   3840
      TabIndex        =   16
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Y1"
      Height          =   195
      Left            =   3840
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "X2"
      Height          =   195
      Left            =   5400
      TabIndex        =   14
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Y2"
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "FBorrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Me.Hide
    Bl = False
End Sub

Private Sub Command2_Click()
If VScroll1 = 0 Then
    Mensa$ = "No hay  Solido seleccionado!!"
    re = MsgBox(Mensa$, vbOKOnly + vbApplicationModal + vbCritical, "Advertencia")
Else
    Mensa$ = "Estas seguro de Borrar el Solido Nº " & VScroll1 & " ?"
    re = MsgBox(Mensa$, vbYesNo + vbApplicationModal + vbQuestion, "Comfirmación")
    If re = vbYes Then
            For II = 1 To Ns - 1
                If II >= VScroll1 Then
                    For Ji = 0 To 6
                        S(Ji, II) = S(Ji, II + 1)
                    Next Ji
                End If
            Next II
            Ns = Ns - 1
            ReDim Preserve S(6, Ns)
            DiDominio
            DiSolido
            VScroll1.Max = Ns
            VScroll1.Min = 0
            VScroll1.Value = 0
        End If
    End If
End Sub

Private Sub Form_Activate()
    Command1.Caption = "Seguir"
    Label7 = "Xmax= " & Xl & Chr$(13) & " Ymax= " & Yl
    CEscala
    DiDominio
    DiSolido
    VScroll1.Max = Ns
    VScroll1.Min = 0
    VScroll1.Value = 0
    Text2 = "No hay Solido seleccionado"
    
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

Private Sub Text1_Change(Index As Integer)
If VScroll1 <> 0 Then
            S(Index, VScroll1) = Val(Text1(Index))
End If
End Sub

Private Sub VScroll1_Change()
If VScroll1 = 0 Then
    Text2 = "No hay Solido seleccionado"
    DiDominio
    DiSolido
    For K = 0 To 6
         Text1(K) = ""
    Next K

Else
    Text2 = "Solido Nº " & VScroll1
    DiDominio
    DiSolido

    For K = 0 To 6
         Text1(K) = S(K, VScroll1)
    Next K
    Picture1.Line (S(0, VScroll1), S(1, VScroll1))-(S(2, VScroll1), S(3, VScroll1)), vbYellow, BF
End If
End Sub

