VERSION 5.00
Begin VB.Form FPR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   3330
   ClientLeft      =   975
   ClientTop       =   1515
   ClientWidth     =   3360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Iteraciones"
      Height          =   1875
      Left            =   120
      TabIndex        =   17
      Top             =   1380
      Width           =   1635
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   660
         TabIndex        =   27
         Text            =   "1"
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   660
         TabIndex        =   25
         Text            =   "1"
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   660
         TabIndex        =   23
         Text            =   "1"
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   660
         TabIndex        =   21
         Text            =   "1"
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   660
         TabIndex        =   19
         Text            =   "2"
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "P"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "v"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "u"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "T"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Max."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Relax"
      Height          =   1575
      Left            =   1800
      TabIndex        =   8
      Top             =   60
      Width           =   1455
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Text            =   "1"
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Text            =   "1"
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Text            =   "1"
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Text            =   "1"
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "P"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "v"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "u"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "T"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores Iniciales"
      Height          =   1275
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1635
      Begin VB.TextBox Text1 
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Text            =   "0"
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Text            =   "0"
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Text            =   "0"
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "v"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "u"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "T"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seguir"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "FPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub


Private Sub Form_Activate()
    Label1(0).Visible = (FE.Check1 = vbChecked)
    Text1(0).Visible = (FE.Check1 = vbChecked)
    Label2(0).Visible = (FE.Check1 = vbChecked)
    Text2(0).Visible = (FE.Check1 = vbChecked)
    Label3(1).Visible = (FE.Check1 = vbChecked)
    Text3(1).Visible = (FE.Check1 = vbChecked)
       
    Label1(1).Visible = (FU.Check1 = vbChecked)
    Text1(1).Visible = (FU.Check1 = vbChecked)
    Label2(1).Visible = (FU.Check1 = vbChecked)
    Text2(1).Visible = (FU.Check1 = vbChecked)
    Label3(2).Visible = (FU.Check1 = vbChecked)
    Text3(2).Visible = (FU.Check1 = vbChecked)
    
    Label1(2).Visible = (FV.Check1 = vbChecked)
    Text1(2).Visible = (FV.Check1 = vbChecked)
    Label2(2).Visible = (FV.Check1 = vbChecked)
    Text2(2).Visible = (FV.Check1 = vbChecked)
    Label3(3).Visible = (FV.Check1 = vbChecked)
    Text3(3).Visible = (FV.Check1 = vbChecked)
    
    Label2(3).Visible = (FU.Check1 = vbChecked) Or (FV.Check1 = vbChecked)
    Text2(3).Visible = (FU.Check1 = vbChecked) Or (FV.Check1 = vbChecked)
    Label3(4).Visible = (FU.Check1 = vbChecked) Or (FV.Check1 = vbChecked)
    Text3(4).Visible = (FU.Check1 = vbChecked) Or (FV.Check1 = vbChecked)
    
End Sub

Private Sub Form_Load()
NTime(0) = 2
NTime(1) = 1
NTime(2) = 1
NTime(3) = 1
NTime(4) = 1
VI(0) = 0
VI(1) = 0
VI(2) = 0
Relax(0) = 1
Relax(1) = 1
Relax(2) = 1
Relax(3) = 1
End Sub


Private Sub Text1_Change(Index As Integer)
VI(Index) = Val(Text1(Index))
End Sub


Private Sub Text2_Change(Index As Integer)
Relax(Index) = Val(Text2(Index))
End Sub


Private Sub Text3_Change(Index As Integer)
NTime(Index) = Val(Text3(Index))
End Sub


