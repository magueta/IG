VERSION 5.00
Begin VB.Form FTfE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   1380
   ClientTop       =   1470
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Bloques"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   1020
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Sc"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Sp  "
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1020
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3195
      Left            =   2400
      ScaleHeight     =   3135
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seguir"
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   2580
      Width           =   1215
   End
End
Attribute VB_Name = "FTfE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    For I = 0 To 1
        Text2(I) = TFE(I + 1, Combo1.ListIndex + 1)
    Next I
    DiMalla
    J = (Combo1.ListIndex) \ Ndx + 1
    I = Combo1.ListIndex + 1 - (J - 1) * Ndx
    Picture1.Line (XU(LL(I - 1)), YV(MM(J - 1)))-(XU(LL(I)), YV(MM(J))), vbWhite, B

End Sub

Private Sub Command1_Click()
Me.Hide
End Sub


Private Sub Form_Activate()
Combo1.Locked = False
CEscala
DiMalla
 
    Combo1.Clear
    Cont = 0
    For J = 1 To Ndy
    For I = 1 To Ndx
        Cont = Cont + 1
        Combo1.AddItem ("Bloque Nº " & Cont)
    Next I
    Next J
    Combo1.ListIndex = 0
End Sub

Sub CEscala()
    Mediaescala = (EscalaXY * 1.1) / 2
    With Picture1
        .ScaleHeight = -2 * Mediaescala
        .ScaleWidth = 2 * Mediaescala
        .ScaleLeft = -Mediaescala + Xl / 2
        .ScaleTop = Yl / 2 + Mediaescala
    End With
End Sub
Sub DiMalla()
    Picture1.Cls
    'For I = 1 To L1
    '     For J = 1 To M1
    '         Picture1.PSet (X(I), Y(J)), QBColor(12)
    '     Next J
    ' Next I
     For I = 2 To L2
         For J = 2 To M2
             Picture1.Line (XU(I), YV(J))-(XU(I), YV(J + 1)), QBColor(8)
             Picture1.Line (XU(I), YV(J))-(XU(I + 1), YV(J)), QBColor(8)
         Next J
     Next I
    Picture1.Line (XU(L1), YV(M1))-(XU(L1), YV(2)), QBColor(8)
    Picture1.Line (XU(2), YV(M1))-(XU(L1), YV(M1)), QBColor(8)
    If Ns > 0 Then
        For K = 1 To Ns
            Picture1.Line (S(0, K), S(1, K))-(S(2, K), S(3, K)), vbMagenta, B
        Next K
    End If
End Sub


Private Sub Text2_Change(Index As Integer)
 TFE(Index + 1, Combo1.ListIndex + 1) = Val(Text2(Index))
End Sub


