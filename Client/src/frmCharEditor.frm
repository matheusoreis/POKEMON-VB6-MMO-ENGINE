VERSION 5.00
Begin VB.Form frmCharEditor 
   Caption         =   "Editor de Personagem - Painel Administrativo"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmBlank 
      Caption         =   "Informações do Jogador"
      Height          =   2775
      Index           =   1
      Left            =   2640
      TabIndex        =   16
      Top             =   1080
      Width           =   3015
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   20
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   17
         Top             =   495
         Width           =   615
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Organização:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmBlank 
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame frmBlank 
      Caption         =   "Visual"
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   12
         Top             =   2295
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   1695
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   495
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Calça:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Camisa:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Cabelo:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCharEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSet_Click(Index As Integer)

End Sub
