VERSION 5.00
Begin VB.Form frmPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Painel administrador"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   8880
      Left            =   5880
      TabIndex        =   69
      Top             =   0
      Width           =   3735
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   0
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   39
         Top             =   480
         Value           =   1
         Width           =   1575
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   1
         Left            =   240
         Max             =   6
         Min             =   1
         TabIndex        =   41
         Top             =   1200
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   2
         Left            =   1920
         Max             =   100
         Min             =   1
         TabIndex        =   40
         Top             =   480
         Value           =   1
         Width           =   1575
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   3
         Left            =   240
         Max             =   15000
         Min             =   1
         TabIndex        =   42
         Top             =   2040
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   4
         Left            =   2040
         Max             =   15000
         Min             =   1
         TabIndex        =   43
         Top             =   2040
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   5
         Left            =   240
         Max             =   15000
         Min             =   1
         TabIndex        =   44
         Top             =   2760
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   6
         Left            =   2040
         Max             =   15000
         Min             =   1
         TabIndex        =   45
         Top             =   2760
         Value           =   1
         Width           =   1455
      End
      Begin VB.CommandButton CmdPokeSpawn 
         Caption         =   "Spawn Pokeball"
         Height          =   375
         Left            =   1800
         TabIndex        =   58
         Top             =   8040
         Width           =   1695
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   7
         Left            =   240
         Max             =   255
         TabIndex        =   46
         Top             =   3480
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   8
         Left            =   240
         Max             =   255
         TabIndex        =   48
         Top             =   4200
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   9
         Left            =   240
         Max             =   255
         TabIndex        =   50
         Top             =   4920
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   10
         Left            =   2040
         Max             =   255
         TabIndex        =   47
         Top             =   3480
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   11
         Left            =   2040
         Max             =   255
         TabIndex        =   49
         Top             =   4200
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   12
         Left            =   240
         Max             =   255
         TabIndex        =   51
         Top             =   5760
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   13
         Left            =   240
         Max             =   255
         TabIndex        =   53
         Top             =   6480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   14
         Left            =   2040
         Max             =   255
         TabIndex        =   52
         Top             =   5760
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   15
         Left            =   2040
         Max             =   255
         TabIndex        =   54
         Top             =   6480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   16
         Left            =   240
         Max             =   500
         TabIndex        =   55
         Top             =   7320
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   17
         Left            =   2040
         Max             =   1
         TabIndex        =   56
         Top             =   7320
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeInfo 
         Height          =   375
         Index           =   18
         Left            =   240
         Max             =   1
         TabIndex        =   57
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Level: 1"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   88
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "#01 Bulbasaur"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   87
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pokéball: 1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   86
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "VitaHp: 1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   85
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "VitalMp: 1"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   84
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "MVitalHp: 1"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   83
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "MVitalMp: 1"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   82
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Line LineBlank 
         Index           =   5
         X1              =   240
         X2              =   3480
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Atq: 0"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   81
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "SpAtk: 0"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   80
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Defesa: 0"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   79
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade: 0"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   78
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Line LineBlank 
         Index           =   7
         X1              =   240
         X2              =   3480
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Spell: None"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   77
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Spell: None"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   76
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Spell: None"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   75
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Spell: None"
         Height          =   255
         Index           =   15
         Left            =   2040
         TabIndex        =   74
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "SpDef: 0"
         Height          =   255
         Index           =   11
         Left            =   2040
         TabIndex        =   73
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Humor: 0"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   72
         Top             =   7080
         Width           =   1695
      End
      Begin VB.Line LineBlank 
         Index           =   2
         X1              =   240
         X2              =   3480
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex: Male"
         Height          =   255
         Index           =   17
         Left            =   2040
         TabIndex        =   71
         Top             =   7080
         Width           =   1335
      End
      Begin VB.Label lblPokeInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Shiny: Não"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   70
         Top             =   7800
         Width           =   1335
      End
   End
   Begin VB.Frame frmBlank 
      Height          =   4680
      Index           =   0
      Left            =   3960
      TabIndex        =   64
      Top             =   15
      Width           =   1815
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   1935
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   3015
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   4095
         Width           =   1335
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Cabelo:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   67
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Camisa:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   66
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Calça:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   65
         Top             =   3480
         Width           =   1095
      End
   End
   Begin VB.Frame frmBlank 
      Height          =   2535
      Index           =   1
      Left            =   3960
      TabIndex        =   61
      Top             =   4680
      Width           =   1815
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Setar"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   1935
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Organização:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   62
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8880
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   3735
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level Up"
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   8280
         Width           =   1575
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   8280
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   375
         Left            =   240
         Min             =   1
         TabIndex        =   24
         Top             =   7800
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   375
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   23
         Top             =   7080
         Value           =   1
         Width           =   3255
      End
      Begin VB.CommandButton cmdSSMap 
         Caption         =   "Screenshot Map"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   3255
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Desovar"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Apagar Bans"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Mapas criados"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Localização"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdAQuest 
         Caption         =   "Quest"
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton cmdTreinador 
         Caption         =   "Treinador"
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   6120
         Width           =   1575
      End
      Begin VB.CommandButton cmdPokedex 
         Caption         =   "Pokémons"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animações"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Habilidades"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   6120
         Width           =   1575
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Recursos"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Mapa"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Ir para o mapa"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "Ir até"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Trazer"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Banir"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Retirar"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Nome do jogador...."
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   7560
         Width           =   3255
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   6840
         Width           =   3255
      End
      Begin VB.Line LineBlank 
         Index           =   3
         X1              =   240
         X2              =   3480
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line LineBlank 
         Index           =   1
         X1              =   240
         X2              =   3480
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line LineBlank 
         Index           =   0
         X1              =   240
         X2              =   3480
         Y1              =   1680
         Y2              =   1680
      End
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Editor de animação
Private Sub cmdAAnim_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    SendRequestEditAnimation

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Banir
Private Sub cmdABan_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Apagar banimentos
Private Sub cmdADestroy_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then

        Exit Sub
    End If

    SendBanDestroy

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de item
Private Sub cmdAItem_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then

        Exit Sub
    End If

    SendRequestEditItem

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Retirar
Private Sub cmdAKick_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub

    If Len(Trim$(txtAName.text)) < 1 Then Exit Sub

    SendKick Trim$(txtAName.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAKick_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Mostrar localização
Private Sub cmdALoc_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

        Exit Sub
    End If

    BLoc = Not BLoc

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Edito de mapa
Private Sub cmdAMap_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If

    SendRequestEditMap

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Mostrar mapas criados
Private Sub cmdAMapReport_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

        Exit Sub
    End If

    SendMapReport

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de Npc
Private Sub cmdANpc_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then

        Exit Sub
    End If

    SendRequestEditNpc

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de Quest
Private Sub cmdAQuest_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub

    SendRequestEditQuest

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAQuest_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de recursos
Private Sub cmdAResource_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then

        Exit Sub
    End If

    SendRequestEditResource

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Desovar mapa
Private Sub cmdARespawn_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

        Exit Sub
    End If

    SendMapRespawn

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de lojas
Private Sub cmdAShop_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then

        Exit Sub
    End If

    SendRequestEditShop

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Spawn Item
Private Sub cmdASpawn_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then

        Exit Sub
    End If

    SendSpawnItem scrlAItem.value, scrlAAmount.value, 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de habilidades
Private Sub cmdASpell_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then

        Exit Sub
    End If

    SendRequestEditSpell

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Ir para o mapa
Private Sub cmdAWarp_Click()
    Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Número do mapa é invalido.", Red)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Trazer para
Private Sub cmdAWarp2Me_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Ir para
Private Sub cmdAWarpMe2_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Level Up
Private Sub cmdLevel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    SendRequestLevelUp

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Editor de pokémons
Private Sub cmdPokedex_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        Exit Sub
    End If

    SendRequestEditPokemon

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdPokedex_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


'ScreenShot Map
Private Sub cmdSSMap_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render the map temp
    ScreenshotMap

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdTreinador_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmCharEditor.Show

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Scrool Quantidade
Private Sub scrlAAmount_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'Scrool Item
Private Sub scrlAItem_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & scrlAItem.value & " " & Trim$(Item(scrlAItem.value).Name)
    If Item(scrlAItem.value).Type = ITEM_TYPE_CURRENCY Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Modificação do personagem
Private Sub cmdSet_Click(Index As Integer)
Select Case Index
    Case 0
        If Len(Trim$(txtAName.text)) < 2 Then Exit Sub
        If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtInfo(0).text)) Then Exit Sub
        
        'Modificar Sprite
        Call SendSetSprite(Trim$(txtAName.text), CLng(Trim$(txtInfo(0).text)))
    Case 5
        If Len(Trim$(txtAName.text)) < 2 Then Exit Sub
        If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtInfo(0).text)) Then Exit Sub
        
        Call SendSetOrganization(Trim$(txtAName.text), CLng(Trim$(txtInfo(5).text)))
End Select
End Sub

' Editor de pokémon
Private Sub CmdPokeSpawn_Click()
    SendSpawnItem 0, 0, 1, scrlPokeInfo(0).value, scrlPokeInfo(1).value, _
                  scrlPokeInfo(2).value, 0, scrlPokeInfo(3).value, scrlPokeInfo(4).value, _
                  scrlPokeInfo(5).value, scrlPokeInfo(6).value, scrlPokeInfo(7).value, _
                  scrlPokeInfo(10).value, scrlPokeInfo(9).value, scrlPokeInfo(8).value, _
                  scrlPokeInfo(11).value, scrlPokeInfo(12).value, scrlPokeInfo(13).value, _
                  scrlPokeInfo(14).value, scrlPokeInfo(15).value, scrlPokeInfo(16).value, scrlPokeInfo(17).value, scrlPokeInfo(18).value
End Sub

Private Sub scrlPokeInfo_Change(Index As Integer)
    Dim Conteudo As String, i As Long
    Select Case Index
    Case 0

        If scrlPokeInfo(Index).value = 0 Then
            lblPokeInfo(Index).Caption = "#00 None"

            For i = 3 To 6
                scrlPokeInfo(i).value = 1
            Next

            scrlPokeInfo(7).value = 0
            scrlPokeInfo(8).value = 0
            scrlPokeInfo(9).value = 0
            scrlPokeInfo(10).value = 0
            scrlPokeInfo(11).value = 0

        Else
            If scrlPokeInfo(Index).value < 10 Then
                lblPokeInfo(Index).Caption = "#0" & scrlPokeInfo(Index).value & " " & Trim$(Pokemon(scrlPokeInfo(Index).value).Name)

                For i = 3 To 6
                    If Pokemon(scrlPokeInfo(0).value).Vital(1) > 0 Then
                        scrlPokeInfo(i).value = Pokemon(scrlPokeInfo(0).value).Vital(1)
                    End If
                Next

                scrlPokeInfo(7).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(1)
                scrlPokeInfo(8).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(3)
                scrlPokeInfo(9).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(2)
                scrlPokeInfo(10).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(4)
                scrlPokeInfo(11).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(5)

            Else
                lblPokeInfo(Index).Caption = "#" & scrlPokeInfo(Index).value & " " & Trim$(Pokemon(scrlPokeInfo(Index).value).Name)

                For i = 3 To 6
                    If Pokemon(scrlPokeInfo(0).value).Vital(1) > 0 Then
                        scrlPokeInfo(i).value = Pokemon(scrlPokeInfo(0).value).Vital(1)
                    End If
                Next

                scrlPokeInfo(7).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(1)
                scrlPokeInfo(8).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(3)
                scrlPokeInfo(9).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(2)
                scrlPokeInfo(10).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(4)
                scrlPokeInfo(11).value = Pokemon(scrlPokeInfo(0).value).Add_Stat(5)

            End If
        End If

        Exit Sub
    Case 1
        Conteudo = "Pokéball: "
    Case 2
        Conteudo = "Level: "
    Case 3
        Conteudo = "VitalHp: "
    Case 4
        Conteudo = "VitalMp: "
    Case 5
        Conteudo = "MVitalHp: "
    Case 6
        Conteudo = "MVitalMp: "
    Case 7
        Conteudo = "Atq: "
    Case 8
        Conteudo = "SpAtk: "
    Case 9
        Conteudo = "Defesa: "
    Case 10
        Conteudo = "Velocidade: "
    Case 11
        Conteudo = "SpDef:"
    Case 12, 13, 14, 15

        If scrlPokeInfo(Index).value = 0 Then
            lblPokeInfo(Index).Caption = "Spell: None"
        Else
            lblPokeInfo(Index).Caption = "Spell: " & Trim$(Spell(scrlPokeInfo(Index).value).Name)
        End If

        Exit Sub
    Case 16
        Conteudo = "Humor:"

    Case 17
        If scrlPokeInfo(Index).value = 0 Then
            lblPokeInfo(Index).Caption = "Sex: Male"
        Else
            lblPokeInfo(Index).Caption = "Sex: Female"
        End If
        Exit Sub

    Case 18
        If scrlPokeInfo(Index).value = 0 Then
            lblPokeInfo(Index).Caption = "Shiny: Não"
        Else
            lblPokeInfo(Index).Caption = "Shiny: Sim"
        End If
        Exit Sub

    End Select

    lblPokeInfo(Index).Caption = Conteudo & scrlPokeInfo(Index).value

End Sub
