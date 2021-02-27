VERSION 5.00
Begin VB.Form frmLeilao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leilão"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   2280
   End
   Begin VB.Timer Tmr1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2280
   End
   Begin VB.Frame fraComprar 
      Caption         =   "Leilão"
      Height          =   5295
      Left            =   7200
      TabIndex        =   14
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdComprar 
         Caption         =   "Comprar"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   4920
         Width           =   2055
      End
      Begin VB.ListBox lstLeilao 
         Height          =   4155
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblTime 
         Caption         =   "Tempo: 0"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblVendedor 
         Caption         =   "Vendedor: None"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Width           =   1695
      End
   End
   Begin VB.Frame fraMyLeiloes 
      Caption         =   "Meus Itens no mercado"
      Height          =   3135
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdRetirar 
         Caption         =   "Retirar item do Mercado"
         Height          =   255
         Left            =   530
         TabIndex        =   13
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ListBox lstMyLeiloes 
         Height          =   2400
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraLeiloar 
      Caption         =   "Seu Mercado"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdLeiloar 
         Caption         =   "Leiloar"
         Height          =   255
         Left            =   670
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton optTH2 
         Caption         =   "2 Horas"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optTH1 
         Caption         =   "1 Hora"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "P.Credit"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dollar"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.TextBox txtPreço 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "1"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cmbBolsa 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Tempo:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Preço:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmLeilao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Price As Long
Private Tempo As Byte
Private Tipo As Byte

Private Sub Form_Load()
    Tipo = 1
    Tempo = 1
    Check1.value = 1
    optTH1.value = 1
    Timer1.Enabled = False
End Sub

Private Sub Check1_Click()
    If Check1.value > 0 Then
        Tipo = 1
        Check2.value = 0
    Else
        Tipo = 2
        Check2.value = 1
    End If
End Sub

Private Sub Check2_Click()
    If Check2.value > 0 Then
        Tipo = 2
        Check1.value = 0
    Else
        Tipo = 1
        Check1.value = 1
    End If
End Sub

Private Sub cmdComprar_Click()
    If lstLeilao.ListIndex < 0 Then
        MsgBox "Escolha um item para comprar!", vbCritical
    Else
        If Leilao(lstLeilao.ListIndex + 1).Vendedor <> vbNullString Then
            If MsgBox("Deseja comprar este item ?", vbYesNo) = vbYes Then
                SendComprar lstLeilao.ListIndex + 1
            Else
                MsgBox "Operação cancelada!", vbInformation
            End If
        Else
            MsgBox "Não tem item algum neste slot!", vbCritical
        End If
    End If
End Sub

Private Sub cmdLeiloar_Click()
    If cmbBolsa.ListIndex > 0 Then
        If GetPlayerInvItemNum(Index, cmbBolsa.ListIndex) = 0 Then
            MsgBox "Você tem que escolher um item para leiloar!", vbCritical
            Exit Sub
        Else
            If GetPlayerInvItemNum(Index, cmbBolsa.ListIndex) = 1 Or GetPlayerInvItemNum(Index, cmbBolsa.ListIndex) = 2 Then
                MsgBox "Você não pode leiloar Dinheiro!", vbCritical
                Exit Sub
            Else
                If txtPreço.text = 0 Then
                    MsgBox "Você tem que definir um preço!", vbCritical
                Else
                    If MsgBox("Você tem certeza que deseja efetuar o Leilão ?", vbYesNo) = vbYes Then
                        SendLeiloar cmbBolsa.ListIndex, txtPreço.text, Tempo, Tipo
                        Timer1.Enabled = True
                    Else
                        MsgBox "Operação cancelada!", vbInformation
                    End If
                End If
            End If
        End If
    Else
        MsgBox "Você deve escolher um item da sua Bolsa para leiloar!", vbCritical
    End If
End Sub

Private Sub cmdRetirar_Click()

    If lstMyLeiloes.ListIndex > 0 Then
        If Leilao(Player(MyIndex).MyLeiloes(lstMyLeiloes.ListIndex)).Vendedor <> GetPlayerName(MyIndex) Then
            MsgBox "Você não pode retirar um item do leilão que não o pertence, desculpe!", vbCritical
        Else
            If MsgBox("Tem certeza que quer retirar este item do leilão ?", vbYesNo) = vbYes Then

                SendRetirar Player(MyIndex).MyLeiloes(lstMyLeiloes.ListIndex)

                Timer1.Enabled = True
            Else
                MsgBox "Operação cancelada!", vbInformation
            End If
        End If
    Else
        MsgBox "Escolha um leilão seu para cancelar...", vbCritical
    End If
End Sub

Private Sub optTH1_Click()
    Tempo = 1
End Sub

Private Sub optTH2_Click()
    Tempo = 2
End Sub

Private Sub Timer1_Timer()
    SendALeilao
    Timer1.Enabled = False
End Sub

Private Sub txtPrice_Change()
    If Not IsNumeric(txtPrice) Then
        MsgBox "Você não pode usar letras aqui!", vbCritical
        txtPrice.text = Price
        Exit Sub
    End If

    If txtPrice.text >= MAX_LONG Then
        MsgBox "Você não pode usar um valor igual ou maior que " & MAX_LONG, vbCritical
        txtPrice.text = Price
        Exit Sub
    End If

    Price = txtPrice.text
End Sub

Private Sub Tmr1_Timer()
    Dim i As Long

    For i = 1 To MAX_LEILAO
        If Leilao(i).Tempo > 1 Then
            Leilao(i).Tempo = Leilao(i).Tempo - 1
        End If
    Next
    If lstLeilao.ListIndex >= 0 Then
        lblTime.Caption = "Tempo: " & Leilao(lstLeilao.ListIndex + 1).Tempo
    End If
End Sub
