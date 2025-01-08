VERSION 5.00
Begin VB.Form FRegion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regiones"
   ClientHeight    =   2865
   ClientLeft      =   5115
   ClientTop       =   2985
   ClientWidth     =   2640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   2640
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   2100
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seguir"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frontera"
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   180
      Width           =   1755
      Begin VB.OptionButton Option1 
         Caption         =   "Inferior"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Izquierda"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Superior"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Derecha"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2100
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Región Nº"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2100
      Width           =   555
   End
End
Attribute VB_Name = "FRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
' caca = True
'     Select Case caca
'        Case Option1(0)
'            Label2 = "xmin.= " & XU(Xi(1, Combo1.ListIndex + 1) + 1)
'            VScroll2.Min = Xi(1, Combo1.ListIndex + 1) + 2
'            VScroll2.Max = L1 + (Combo1.ListIndex + 1) - Ndxi
'            VScroll2 = Xi(2, Combo1.ListIndex + 1) + 2
'            Label1 = "xmax.= "
'        Case Option1(1)
'            Label2 = "ymin.= " & YV(Yi(1, Combo1.ListIndex + 1) + 1)
'            VScroll2.Min = Yi(1, Combo1.ListIndex + 1) + 2
'            VScroll2.Max = M1 + (Combo1.ListIndex + 1) - Ndyi
'            VScroll2 = Yi(2, Combo1.ListIndex + 1) + 2
'            Label1 = "ymax.= "
'        Case Option1(2)
'            Label2 = "xmin.= " & XU(Xs(1, Combo1.ListIndex + 1) + 1)
'            VScroll2.Min = Xs(1, Combo1.ListIndex + 1) + 2
'            VScroll2.Max = L1 + (Combo1.ListIndex + 1) - Ndxs
'            VScroll2 = Xs(2, Combo1.ListIndex + 1) + 2
'            Label1 = "xmax.= "
'        Case Option1(3)
'            Label2 = "ymin.= " & YV(Yd(1, Combo1.ListIndex + 1) + 1)
'            VScroll2.Min = Yd(1, Combo1.ListIndex + 1) + 2
'            VScroll2.Max = M1 + (Combo1.ListIndex + 1) - Ndyd
'            VScroll2 = Yd(2, Combo1.ListIndex + 1) + 2
'            Label1 = "ymax.= "
'    End Select
End Sub

Private Sub Combo1_Click()
  caca = True
     Select Case caca
        Case Option1(0)
            Label2 = "xmin.= " & XU(Xi(1, Combo1.ListIndex + 1) + 1)
            tep = Xi(2, Combo1.ListIndex + 1) + 2
            VScroll2.Min = Xi(1, Combo1.ListIndex + 1) + 2
            VScroll2.Max = L1 + (Combo1.ListIndex + 1) - Ndxi
            If tep < VScroll2.Min Then tep = VScroll2.Min
            VScroll2 = tep
            Label1 = "xmax.= "
        Case Option1(1)
            Label2 = "ymin.= " & YV(Yi(1, Combo1.ListIndex + 1) + 1)
            tep = Yi(2, Combo1.ListIndex + 1) + 2
            VScroll2.Min = Yi(1, Combo1.ListIndex + 1) + 2
            VScroll2.Max = M1 + (Combo1.ListIndex + 1) - Ndyi
            If tep < VScroll2.Min Then tep = VScroll2.Min
            VScroll2 = tep
            Label1 = "ymax.= "
        Case Option1(2)
            Label2 = "xmin.= " & XU(Xs(1, Combo1.ListIndex + 1) + 1)
            tep = Xs(2, Combo1.ListIndex + 1) + 2
            VScroll2.Min = Xs(1, Combo1.ListIndex + 1) + 2
            VScroll2.Max = L1 + (Combo1.ListIndex + 1) - Ndxs
            If tep < VScroll2.Min Then tep = VScroll2.Min
            VScroll2 = tep
            Label1 = "xmax.= "
        Case Option1(3)
            Label2 = "ymin.= " & YV(Yd(1, Combo1.ListIndex + 1) + 1)
            tep = Yd(2, Combo1.ListIndex + 1) + 2
            VScroll2.Min = Yd(1, Combo1.ListIndex + 1) + 2
            VScroll2.Max = M1 + (Combo1.ListIndex + 1) - Ndyd
            If tep < VScroll2.Min Then tep = VScroll2.Min
            VScroll2 = tep
            Label1 = "ymax.= "
    End Select
  

End Sub


Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
   For kk = 3 To 0 Step -1
        Option1(kk).Enabled = Not (FCF.Check2(kk).Value = vbChecked) And Val(FCF.Text1(kk)) > 1
        Option1(kk) = Not (FCF.Check2(kk).Value = vbChecked) And Val(FCF.Text1(kk)) > 1
        
    Next kk
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub

Private Sub Option1_Click(Index As Integer)
     Combo1.Clear
    For J = 0 To 3
        If Option1(J) Then
            Select Case (J)
                Case 0
                    For I = 1 To Ndxi - 1
                        Combo1.AddItem (I)
                        
                        
                    Next I
                    Combo1.ListIndex = 0
                Case 1
                    For I = 1 To Ndyi - 1
                        Combo1.AddItem (I)
                        
                        
                    Next I
                    Combo1.ListIndex = 0
                Case 2
                    For I = 1 To Ndxs - 1
                       Combo1.AddItem (I)
                      
                       
                    Next I
                    Combo1.ListIndex = 0
                Case 3
                    For I = 1 To Ndyd - 1
                        Combo1.AddItem (I)
                        
                        
                    Next I
                    Combo1.ListIndex = 0
            End Select
    End If
Next J
  
End Sub

Private Sub VScroll2_Change()
    caca = True
    Select Case caca
        Case Option1(0)
            Label2 = "xmin.= " & XU(Xi(1, Combo1.ListIndex + 1) + 1)
            Text2 = XU(VScroll2)
            Xi(2, Combo1.ListIndex + 1) = VScroll2 - 2
            Xi(1, Combo1.ListIndex + 2) = VScroll2 - 1

        Case Option1(1)
            Label2 = "ymin.= " & YV(Yi(1, Combo1.ListIndex + 1) + 1)
            Text2 = YV(VScroll2)
            Yi(2, Combo1.ListIndex + 1) = VScroll2 - 2
            Yi(1, Combo1.ListIndex + 2) = VScroll2 - 1

        Case Option1(2)
            Label2 = "xmin.= " & XU(Xs(1, Combo1.ListIndex + 1) + 1)
            Text2 = XU(VScroll2)
            Xs(2, Combo1.ListIndex + 1) = VScroll2 - 2
            Xs(1, Combo1.ListIndex + 2) = VScroll2 - 1

        Case Option1(3)
            Label2 = "ymin.= " & YV(Yd(1, Combo1.ListIndex + 1) + 1)
            Label1 = "ymax.= "
            Text2 = YV(VScroll2)
            Yd(2, Combo1.ListIndex + 1) = VScroll2 - 2
            Yd(1, Combo1.ListIndex + 2) = VScroll2 - 1

    End Select

End Sub

