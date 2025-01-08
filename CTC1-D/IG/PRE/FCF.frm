VERSION 5.00
Begin VB.Form FCF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Condiciones de Fronteras"
   ClientHeight    =   4785
   ClientLeft      =   945
   ClientTop       =   495
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4620
   Begin VB.CommandButton Command5 
      Caption         =   "Seguir"
      Height          =   435
      Left            =   1110
      TabIndex        =   23
      Top             =   4260
      Width           =   2355
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ecuación de Momento en Y"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1110
      TabIndex        =   22
      Top             =   3660
      Width           =   2355
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ecuación de Momento en X"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1110
      TabIndex        =   21
      Top             =   3120
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ecuación de Energía"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1110
      TabIndex        =   20
      Top             =   2580
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Caption         =   "Regiones de Fronteras"
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   1980
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   18
         Top             =   1620
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Height          =   195
         Index           =   3
         Left            =   2220
         TabIndex        =   17
         Top             =   1620
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   16
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   14
         Top             =   1260
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Height          =   195
         Index           =   2
         Left            =   2220
         TabIndex        =   13
         Top             =   1260
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Text            =   "1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   900
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Height          =   195
         Index           =   1
         Left            =   2220
         TabIndex        =   9
         Top             =   900
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   5
         Top             =   540
         Width           =   195
      End
      Begin VB.CheckBox Check2 
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   4
         Top             =   540
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Eje de Simetría"
         Height          =   195
         Left            =   2970
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Derecha"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Superior"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Izquierda"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Por Bloques"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Inferior"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   855
      End
   End
End
Attribute VB_Name = "FCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



 Sub Check1_Click()
    If Check1 = vbChecked Then
        For I = 0 To 3
            Option1(I).Enabled = True
            Option1(0) = Not Check1
        Next I
    Else
        For I = 0 To 3
            Option1(I).Enabled = False
            Option1(I) = False
            Check2(I).Enabled = True
            Text1(I).Enabled = Not (vbChecked = Check2(I))
        Next I
        Simetria = 0
    End If
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
   
End Sub

 Sub Check2_Click(Index As Integer)
    Select Case Index
        Case 0
            Text1(Index).Enabled = Not (vbChecked = Check2(Index))
            Text1(Index) = Ndx
        Case 1
            Text1(Index).Enabled = Not (vbChecked = Check2(Index))
            Text1(Index) = Ndy
        Case 2
            Text1(Index).Enabled = Not (vbChecked = Check2(Index))
            Text1(Index) = Ndx
        Case 3
            Text1(Index).Enabled = Not (vbChecked = Check2(Index))
            Text1(Index) = Ndy
    End Select
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    
End Sub

Private Sub Command1_Click()
    Call ObtenerE
    For I = 0 To 3
        If Not (Check2(I).Value = vbChecked) _
        And Val(Text1(I)) > 1 Then
            FRegion.Show (vbModal)
            Exit For
        End If
    Next I
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    
    
End Sub

 Sub Command2_Click()
    FE.Show (1)
End Sub

 Sub Command3_Click()
    FU.Show (1)
End Sub


 Sub Command4_Click()
    FV.Show (1)
End Sub

Sub Command5_Click()
    Me.Hide
End Sub


 Sub Form_Activate()
 
 ReDim Preserve TFE(2, Ndx * Ndy) As Single

 End Sub

 Sub Form_Load()
Ndxi = 1
Ndxs = 1
Ndyi = 1
Ndyd = 1
End Sub

 Sub Label1_Click(Index As Integer)

End Sub


 Sub Option1_Click(Index As Integer)
    For J = 0 To 3
        Text1(J).Enabled = True
        Check2(J).Enabled = True
        
    Next J
    Check2(Index) = 0
    Text1(Index).Enabled = Not Option1(Index)
    Check2(Index).Enabled = Not Option1(Index)
    Text1(Index) = 1
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    
    If Option1(Index) Then
        Simetria = Index + 1
    Else
        Simetria = 0
    End If
    
End Sub

 Sub Text1_Change(Index As Integer)
    Select Case Index
        Case 0
            Ndxi = Val(Text1(Index))
        Case 1
            Ndyi = Val(Text1(Index))
        Case 2
            Ndxs = Val(Text1(Index))
        Case 3
            Ndyd = Val(Text1(Index))
    End Select
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
  
End Sub


