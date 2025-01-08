VERSION 5.00
Begin VB.Form FIni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dominio"
   ClientHeight    =   2925
   ClientLeft      =   3090
   ClientTop       =   2025
   ClientWidth     =   2385
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   2385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   1020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Dominio.frx":0000
      Left            =   540
      List            =   "Dominio.frx":000A
      TabIndex        =   6
      Text            =   "Cartesianas"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seguir"
      Height          =   495
      Left            =   540
      TabIndex        =   4
      Top             =   2340
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "R(1)"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Radio de referencia al eje axial"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tipo de coordenadas"
      Height          =   195
      Left            =   390
      TabIndex        =   5
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Yl"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Longitud en la dirección y"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Xl"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Longitud en la dirección x"
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "FIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    Text3.Visible = False
    Label4.Visible = False
Else
    Text3.Visible = True
    Label4.Visible = True
End If

End Sub


Private Sub Command1_Click()
Xl = Val(Text1)
Yl = Val(Text2)

Mode = 1
If Combo1.ListIndex = 1 Then
    Mode = 2
    R1 = Val(Text3)
End If
If Xl <= 0 Or Yl <= 0 Then
     Mensa$ = "Xl o Yl no puede ser igual o menor que 0"
    re = MsgBox(Mensa$, vbOKOnly + vbApplicationModal + vbInformation, "Error")
Else
    EscalaXY = Xl / Yl
    Fver.Show
    Unload FIni
End If


End Sub


