VERSION 5.00
Begin VB.Form FTfU 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1365
   ClientLeft      =   4905
   ClientTop       =   2280
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Seguir"
      Height          =   255
      Left            =   780
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fureza externa"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2475
      Begin VB.TextBox Text1 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Aceleración x"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
   End
End
Attribute VB_Name = "FTfU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub


Private Sub Text1_Change()
Gx = Val(Text1)
End Sub

