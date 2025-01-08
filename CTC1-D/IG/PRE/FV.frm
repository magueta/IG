VERSION 5.00
Begin VB.Form FV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ecuación de momento en Y"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   375
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frontera"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   360
      TabIndex        =   19
      Top             =   300
      Width           =   1755
      Begin VB.OptionButton Option1 
         Caption         =   "Derecha"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Superior"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Izquierda"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Inferior"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ubicación"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   1755
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Región Nº"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "V.C.S."
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   18
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "V.C.I."
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   660
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Frontera"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   360
      TabIndex        =   4
      Top             =   2940
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "Flujo t. Desarr."
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Perfil Parabólico"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Constante"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Valor"
         Height          =   915
         Left            =   60
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   7
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   6
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "h"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "v"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Activar"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Térnino Fuente"
      Enabled         =   0   'False
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seguir"
      Height          =   255
      Left            =   420
      TabIndex        =   1
      Top             =   5700
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   2700
      ScaleHeight     =   6075
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   60
      Width           =   6135
   End
End
Attribute VB_Name = "FV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub LlenarC()
    Combo1.Clear
    For J = 0 To 3
        If Option1(J) Then
            Select Case (J)
                Case 0
                    For I = 1 To Ndxi
                        Combo1.AddItem (I)
                    Next I
                Case 1
                    For I = 1 To Ndyi
                        Combo1.AddItem (I)
                    Next I
                Case 2
                    For I = 1 To Ndxs
                        Combo1.AddItem (I)
                    Next I
                Case 3
                    For I = 1 To Ndyd
                        Combo1.AddItem (I)
                    Next I
            End Select
        End If
    Next J
    Combo1.ListIndex = 0
    
End Sub

 Sub Check1_Click()
    Frame1.Enabled = (Check1 = vbChecked)
    Frame2.Enabled = (Check1 = vbChecked)
    Frame3.Enabled = (Check1 = vbChecked)
    Command1.Enabled = (Check1 = vbChecked)
End Sub


 Sub Combo1_Change()
Combo1.Locked = False

    
End Sub

 Sub Combo1_Click()
 DiMalla
 caca = True
 Select Case caca
    Case Option1(0)
        Text1(0) = Xi(1, Combo1.ListIndex + 1)
        Text1(1) = Xi(2, Combo1.ListIndex + 1)
        Picture1.Line (XU(Xi(1, Combo1.ListIndex + 1) + 1), YV(2))-(XU(Xi(2, Combo1.ListIndex + 1) + 2), YV(2)), vbBlue
        
    Case Option1(1)
        Text1(0) = Yi(1, Combo1.ListIndex + 1)
        Text1(1) = Yi(2, Combo1.ListIndex + 1)
        
        Picture1.Line (XU(2), YV(Yi(1, Combo1.ListIndex + 1) + 1))-(XU(2), YV(Yi(2, Combo1.ListIndex + 1) + 2)), vbBlue
    Case Option1(2)
        Text1(0) = Xs(1, Combo1.ListIndex + 1)
        Text1(1) = Xs(2, Combo1.ListIndex + 1)
        Picture1.Line (XU(Xs(1, Combo1.ListIndex + 1) + 1), YV(M1))-(XU(Xs(2, Combo1.ListIndex + 1) + 2), YV(M1)), vbBlue
    Case Option1(3)
        Text1(0) = Yd(1, Combo1.ListIndex + 1)
        Text1(1) = Yd(2, Combo1.ListIndex + 1)
        Picture1.Line (XU(L1), YV(Yd(1, Combo1.ListIndex + 1) + 1))-(XU(L1), YV(Yd(2, Combo1.ListIndex + 1) + 2)), vbBlue
End Select
Select Case True
    Case Option1(0)
        Option2(Vxi(1, Combo1.ListIndex + 1)) = True
        For I = 0 To 1
            Text2(I) = Vxi(I + 2, Combo1.ListIndex + 1)
        Next I
    Case Option1(1)
        Option2(Vyi(1, Combo1.ListIndex + 1)) = True
        For I = 0 To 1
            Text2(I) = Vyi(I + 2, Combo1.ListIndex + 1)
        Next I
    Case Option1(2)
        Option2(Vxs(1, Combo1.ListIndex + 1)) = True
        For I = 0 To 1
            Text2(I) = Vxs(I + 2, Combo1.ListIndex + 1)
        Next I
    Case Option1(3)
        Option2(Vyd(1, Combo1.ListIndex + 1)) = True
        For I = 0 To 1
            Text2(I) = Vyd(I + 2, Combo1.ListIndex + 1)
        Next I
End Select

End Sub

 Sub Command1_Click()
    FTfV.Show (1)
    For I = 0 To 3
        Option1(I) = False
    Next I
    For I = 0 To 2
        Option2(I) = False
    Next I
    Frame4.Visible = False
End Sub

Sub Command2_Click()
  Me.Hide
End Sub


 Sub Form_Activate()
    CEscala
    Picture1.AutoRedraw = True
    DiMalla
    For I = 0 To 3
        Option1(I).Enabled = True
    Next I
    Select Case Simetria
        Case 0
            For I = 0 To 3
                Option1(I).Enabled = True
            Next I
        Case Else
            Option1(Simetria - 1).Enabled = False
            Option1(Simetria - 1) = False
    End Select
    Combo1 = ""
    Text1(0) = ""
    Text1(1) = ""
    Text2(0) = ""
    Text2(1) = ""
    For I = 0 To 3
        Option1(I) = False
    Next I
    For I = 0 To 2
        Option2(I) = False
    Next I

End Sub

Sub CEscala()
    Mediaescala = (EscalaXY * 1.1) / 2
    Picture1.ScaleHeight = -2 * Mediaescala
    Picture1.ScaleWidth = 2 * Mediaescala
    Picture1.ScaleLeft = -Mediaescala + Xl / 2
    Picture1.ScaleTop = Yl / 2 + Mediaescala
End Sub

Sub DiMalla()
    Picture1.Cls

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
    Select Case Simetria
        Case 1
            Picture1.Line (XU(2), YV(2))-(XU(L1), YV(2)), vbCyan
        Case 2
            Picture1.Line (XU(2), YV(2))-(XU(2), YV(M1)), vbCyan
        Case 3
            Picture1.Line (XU(L1), YV(M1))-(XU(2), YV(M1)), vbCyan
        Case 4
            Picture1.Line (XU(L1), YV(M1))-(XU(L1), YV(2)), vbCyan
    End Select
End Sub

 Sub Form_Load()
Combo1.Clear
For I = 1 To Ndxi
    Vxi(1, I) = 2
    Vxi(2, I) = 0
    Vxi(3, I) = 0
    
Next I
For I = 1 To Ndyi
    Vyi(1, I) = 2
    Vyi(2, I) = 0
    Vyi(3, I) = 0
Next I
For I = 1 To Ndxs
    Vxs(1, I) = 2
    Vxs(2, I) = 0
    Vxs(3, I) = 0
Next I
For I = 1 To Ndyd
    Vyd(1, I) = 2
    Vyd(2, I) = 0
    Vyd(3, I) = 0
Next I

End Sub

Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


 Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


 Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


 Sub Label1_Click()

End Sub


Sub Label2_Click(Index As Integer)

End Sub


Sub Label3_Click(Index As Integer)

End Sub


Sub Option1_Click(Index As Integer)
Call LlenarC

'For I = 0 To 4
'    Option2(I) = False
'Next I
Select Case True
    Case Option1(0)
        Option2(Vxi(1, Combo1.ListIndex + 1)) = True
        Option2(1).Enabled = True
        Option2(0).Enabled = True
    Case Option1(1)
        Option2(Vyi(1, Combo1.ListIndex + 1)) = True
        Option2(1).Enabled = False
        Option2(0).Enabled = False
    Case Option1(2)
        Option2(Vxs(1, Combo1.ListIndex + 1)) = True
        Option2(1).Enabled = True
        Option2(0).Enabled = True
    Case Option1(3)
        Option2(Vyd(1, Combo1.ListIndex + 1)) = True
        Option2(1).Enabled = False
        Option2(0).Enabled = False
End Select

End Sub


Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0
            Frame4.Visible = False
            For I = 0 To 1
                Text2(I).Visible = False
            Next I
            
            Select Case True
                    Case Option1(0)
                        Vxi(1, Combo1.ListIndex + 1) = Index
                    Case Option1(1)
                        Vyi(1, Combo1.ListIndex + 1) = Index
                    Case Option1(2)
                        Vxs(1, Combo1.ListIndex + 1) = Index
                    Case Option1(3)
                        Vyd(1, Combo1.ListIndex + 1) = Index
                End Select
            

        Case 1
            Frame4.Visible = True
            Text2(1).Visible = True
            Text2(0).Visible = True
            Label3(0) = "v ext."
            Label3(1) = "v cen."
            Label3(1).Visible = True
                Select Case True
                    Case Option1(0)
                        For I = 0 To 1
                            Text2(I) = Vxi(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vxi(1, Combo1.ListIndex + 1) = Index
                    Case Option1(1)
                        For I = 0 To 1
                            Text2(I) = Vyi(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vyi(1, Combo1.ListIndex + 1) = Index
                    Case Option1(2)
                        For I = 0 To 1
                            Text2(I) = Vxs(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vxs(1, Combo1.ListIndex + 1) = Index
                    Case Option1(3)
                        For I = 0 To 1
                            Text2(I) = Vyd(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vyd(1, Combo1.ListIndex + 1) = Index
                    End Select
            
            
        Case 2
            Frame4.Visible = True
            Text2(1).Visible = False
            Text2(0).Visible = True
            Label3(0) = "v"
            Label3(1).Visible = False
            
                Select Case True
                    Case Option1(0)
                        For I = 0 To 1
                            Text2(I) = Vxi(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vxi(1, Combo1.ListIndex + 1) = Index
                    Case Option1(1)
                        For I = 0 To 1
                            Text2(I) = Vyi(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vyi(1, Combo1.ListIndex + 1) = Index
                    Case Option1(2)
                        For I = 0 To 1
                            Text2(I) = Vxs(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vxs(1, Combo1.ListIndex + 1) = Index
                    Case Option1(3)
                        For I = 0 To 1
                            Text2(I) = Vyd(I + 2, Combo1.ListIndex + 1)
                        Next I
                        Vyd(1, Combo1.ListIndex + 1) = Index
                End Select
    End Select
End Sub


Sub Picture1_Click()

End Sub


Sub Text1_Change(Index As Integer)

End Sub


Sub Text2_Change(Index As Integer)
Select Case True
    Case Option1(0)
        If Index = 0 Then
             Vxi(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        Else
            Vxi(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        End If
        Case Option1(1)
        
        If Index = 0 Then
             Vyi(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        Else
            Vyi(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        End If
        
        Case Option1(2)
        
        If Index = 0 Then
             Vxs(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        Else
            Vxs(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        End If
        
        Case Option1(3)
        
        If Index = 0 Then
             Vyd(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        Else
            Vyd(Index + 2, Combo1.ListIndex + 1) = Val(Text2(Index))
        End If
        
End Select
End Sub


