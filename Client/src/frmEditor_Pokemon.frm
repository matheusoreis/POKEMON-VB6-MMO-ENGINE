VERSION 5.00
Begin VB.Form frmEditor_Pokemon 
   Caption         =   "Pokémon Editor"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   27
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame FrameBlank 
      Caption         =   "Data Editor"
      Height          =   7335
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame1 
         Caption         =   "Animated Frame"
         Height          =   1815
         Left            =   1800
         TabIndex        =   48
         Top             =   2880
         Width           =   1695
         Begin VB.HScrollBar scrlAnimFrame 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   5
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAnimFrame 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   5
            TabIndex        =   49
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblAnimFrame 
            Caption         =   "Fly/Surf: 0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblAnimFrame 
            Caption         =   "Normal: 0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame FrameBlank 
         Caption         =   "Male/Female %"
         Height          =   1815
         Index           =   8
         Left            =   120
         TabIndex        =   43
         Top             =   2880
         Width           =   1575
         Begin VB.HScrollBar scrlControlSex 
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   100
            TabIndex        =   45
            Top             =   1200
            Value           =   50
            Width           =   1215
         End
         Begin VB.HScrollBar scrlControlSex 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   100
            TabIndex        =   44
            Top             =   600
            Value           =   50
            Width           =   1215
         End
         Begin VB.Label lblControlSex 
            BackStyle       =   0  'Transparent
            Caption         =   "Female: 50%"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblControlSex 
            BackStyle       =   0  'Transparent
            Caption         =   "Male: 50%"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame FrameBlank 
         Caption         =   "Status Base"
         Height          =   2055
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   4680
         Width           =   3375
         Begin VB.HScrollBar scrlBaseExp 
            Height          =   255
            LargeChange     =   10
            Left            =   2040
            Max             =   5
            TabIndex        =   62
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar ScrlVitals 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   1080
            Max             =   255
            TabIndex        =   42
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar ScrlVitals 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   39
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar ScrlStat 
            Height          =   255
            Index           =   5
            LargeChange     =   10
            Left            =   1080
            Max             =   255
            TabIndex        =   37
            Top             =   1680
            Width           =   855
         End
         Begin VB.HScrollBar ScrlStat 
            Height          =   255
            Index           =   4
            LargeChange     =   10
            Left            =   2040
            Max             =   255
            TabIndex        =   35
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar ScrlStat 
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   33
            Top             =   1680
            Width           =   855
         End
         Begin VB.HScrollBar ScrlStat 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   1080
            Max             =   255
            TabIndex        =   31
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar ScrlStat 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   29
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblBaseExp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp: 0"
            Height          =   195
            Left            =   2040
            TabIndex        =   63
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label lblHP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP: 0"
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   41
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   420
         End
         Begin VB.Label lblHP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HP: 0"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   405
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SpDef: 0"
            Height          =   195
            Index           =   5
            Left            =   1080
            TabIndex        =   38
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   630
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed: 0"
            Height          =   195
            Index           =   4
            Left            =   2040
            TabIndex        =   36
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SpAtk: 0"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   615
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Defesa: 0"
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   32
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   690
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ataque: 0"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   690
         End
      End
      Begin VB.Frame FrameBlank 
         Caption         =   "Habilidades - 1"
         Height          =   1215
         Index           =   4
         Left            =   3600
         TabIndex        =   20
         Top             =   4680
         Width           =   3375
         Begin VB.HScrollBar scrlHabilidadeIndex 
            Height          =   255
            Left            =   120
            Max             =   20
            Min             =   1
            TabIndex        =   54
            Top             =   840
            Value           =   1
            Width           =   1335
         End
         Begin VB.HScrollBar scrlHabilidadeLevel 
            Height          =   255
            LargeChange     =   10
            Left            =   1920
            Max             =   255
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar scrlHabilidadeSpell 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblHabLevel 
            AutoSize        =   -1  'True
            Caption         =   "Level: 0"
            Height          =   195
            Left            =   1920
            TabIndex        =   24
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   570
         End
         Begin VB.Label lblHabSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell: None"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame FrameBlank 
         Height          =   2655
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6855
         Begin VB.HScrollBar scrlPokemon 
            Height          =   255
            Left            =   3000
            Max             =   255
            TabIndex        =   10
            Top             =   1350
            Width           =   975
         End
         Begin VB.HScrollBar scrlCRate 
            Height          =   255
            LargeChange     =   10
            Left            =   4320
            Max             =   255
            TabIndex        =   66
            Top             =   1680
            Width           =   855
         End
         Begin VB.HScrollBar scrlEggTime 
            Height          =   255
            LargeChange     =   10
            Left            =   5400
            Max             =   255
            TabIndex        =   64
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar scrlExpType 
            Height          =   255
            LargeChange     =   10
            Left            =   4320
            Max             =   255
            TabIndex        =   60
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar scrlAnimAttack 
            Height          =   255
            LargeChange     =   10
            Left            =   4320
            Max             =   255
            TabIndex        =   58
            Top             =   480
            Width           =   855
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Index           =   2
            ItemData        =   "frmEditor_Pokemon.frx":0000
            Left            =   2880
            List            =   "frmEditor_Pokemon.frx":003D
            TabIndex        =   17
            Text            =   "Nenhum"
            Top             =   2100
            Width           =   1215
         End
         Begin VB.TextBox txtDesc 
            Height          =   1455
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   400
            Width           =   2655
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Index           =   1
            ItemData        =   "frmEditor_Pokemon.frx":00DE
            Left            =   2880
            List            =   "frmEditor_Pokemon.frx":011B
            TabIndex        =   12
            Text            =   "Nenhum"
            Top             =   1720
            Width           =   1215
         End
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   960
            Left            =   3000
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   11
            Top             =   400
            Width           =   960
         End
         Begin VB.Label lblCRate 
            AutoSize        =   -1  'True
            Caption         =   "C.Rate: 0"
            Height          =   195
            Left            =   4320
            TabIndex        =   67
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   675
         End
         Begin VB.Label lblEggTime 
            AutoSize        =   -1  'True
            Caption         =   "Egg Time: 0"
            Height          =   195
            Left            =   5400
            TabIndex        =   65
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   855
         End
         Begin VB.Label lblExpType 
            AutoSize        =   -1  'True
            Caption         =   "Exp: Rápido"
            Height          =   195
            Left            =   4320
            TabIndex        =   61
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   870
         End
         Begin VB.Line Line1 
            X1              =   4200
            X2              =   4200
            Y1              =   2520
            Y2              =   240
         End
         Begin VB.Label lblAnimAttack 
            AutoSize        =   -1  'True
            Caption         =   "AnimAttack: None"
            Height          =   195
            Left            =   4320
            TabIndex        =   59
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label3 
            Caption         =   "PokéDex Description:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   200
            UseMnemonic     =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame FrameBlank 
         Caption         =   "Evolução - 1"
         Height          =   1815
         Index           =   2
         Left            =   3600
         TabIndex        =   4
         Top             =   2880
         Width           =   3375
         Begin VB.HScrollBar scrlFelicidade 
            Height          =   255
            LargeChange     =   10
            Left            =   1920
            Max             =   255
            TabIndex        =   56
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar scrlEvolutionIndex 
            Height          =   255
            Left            =   120
            Max             =   8
            Min             =   1
            TabIndex        =   55
            Top             =   1440
            Value           =   1
            Width           =   1215
         End
         Begin VB.CheckBox chkEvo 
            Caption         =   "N.Evol"
            Height          =   255
            Left            =   1920
            TabIndex        =   53
            Top             =   1440
            Width           =   855
         End
         Begin VB.HScrollBar scrlEvolutionStone 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   18
            Top             =   1080
            Width           =   1215
         End
         Begin VB.HScrollBar scrlEvolutionLevel 
            Height          =   255
            LargeChange     =   10
            Left            =   1920
            Max             =   255
            TabIndex        =   7
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar scrlEvolutionPoke 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   500
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblFelicidade 
            AutoSize        =   -1  'True
            Caption         =   "Felicidade: 0"
            Height          =   195
            Left            =   1920
            TabIndex        =   57
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   900
         End
         Begin VB.Label lblEvoStone 
            AutoSize        =   -1  'True
            Caption         =   "Stone: None"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblEvoLevel 
            AutoSize        =   -1  'True
            Caption         =   "Level: 0"
            Height          =   195
            Left            =   1920
            TabIndex        =   8
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   810
         End
         Begin VB.Label lblEvoPoke 
            AutoSize        =   -1  'True
            Caption         =   "Pokémon: None"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1155
         End
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Frame FrameBlank 
      Caption         =   "Pokémon Editor"
      Height          =   7335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6885
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Pokemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EvolutionIndex As Long
Dim HabilidadeIndex As Long

Private Sub chkEvo_Click()
    Pokemon(EditorIndex).NotEvo = chkEvo.value
End Sub

Private Sub cmbTipo_Click(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Pokemon(EditorIndex).Tipo(Index) = cmbTipo(Index).ListIndex
End Sub

Private Sub cmdCancel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call PokemonEditorCancel

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_POKEMONS Then Exit Sub

    ClearPokemon EditorIndex

    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    PokemonEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call PokemonEditorOk

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Long

    EvolutionIndex = 1
    HabilidadeIndex = 1
    scrlEvolutionIndex.max = 8
    scrlPokemon.max = NumCharacters

End Sub

Private Sub lstIndex_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    PokemonEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimAttack_Change()

    If scrlAnimAttack.value > 0 Then
        lblAnimAttack.Caption = "AnimAttack:" & Trim$(Animation(scrlAnimAttack.value).Name)
    Else
        lblAnimAttack.Caption = "AnimAttack: None"
    End If

    Pokemon(EditorIndex).AnimAttack = scrlAnimAttack.value

End Sub

Private Sub scrlAnimFrame_Change(Index As Integer)

    Select Case Index
    Case 0
        lblAnimFrame(0).Caption = "Normal:" & scrlAnimFrame(0).value
        Pokemon(EditorIndex).AnimFrame(1) = scrlAnimFrame(0).value
    Case 1
        lblAnimFrame(1).Caption = "Fly/Surf:" & scrlAnimFrame(1).value
        Pokemon(EditorIndex).AnimFrame(2) = scrlAnimFrame(1).value
    End Select

End Sub

Private Sub scrlBaseExp_Change()
    lblBaseExp.Caption = "Exp: " & scrlBaseExp.value
    Pokemon(EditorIndex).ExpBase = scrlBaseExp.value
End Sub

Private Sub scrlControlSex_Change(Index As Integer)

    Select Case Index
    Case 0
        scrlControlSex(1).value = 100 - scrlControlSex(0).value
        lblControlSex(0).Caption = "Male: " & scrlControlSex(0).value & "%"
        lblControlSex(1).Caption = "Female: " & scrlControlSex(1).value & "%"
        Pokemon(EditorIndex).ControlSex = scrlControlSex(0).value
    End Select

End Sub

Private Sub scrlCRate_Change()
    lblCRate.Caption = "C.Rate: " & scrlCRate.value
    Pokemon(EditorIndex).CRate = scrlCRate.value
End Sub

Private Sub scrlEggTime_Change()
    lblEggTime.Caption = "Egg Time: " & scrlEggTime.value
    Pokemon(EditorIndex).EggTime = scrlEggTime.value
End Sub

Private Sub scrlEvolutionIndex_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EvolutionIndex = scrlEvolutionIndex.value
    FrameBlank(2).Caption = "Evolução - " & EvolutionIndex
    scrlEvolutionPoke.value = Pokemon(EditorIndex).Evolução(EvolutionIndex).Pokemon
    scrlEvolutionLevel.value = Pokemon(EditorIndex).Evolução(EvolutionIndex).Level
    scrlEvolutionStone.value = Pokemon(EditorIndex).Evolução(EvolutionIndex).Pedra

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEvolutionIndex_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEvolutionLevel_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblEvoLevel.Caption = "Level: " & scrlEvolutionLevel.value
    Pokemon(EditorIndex).Evolução(EvolutionIndex).Level = scrlEvolutionLevel.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEvolutionLevel_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEvolutionPoke_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlEvolutionPoke.value > 0 Then
        lblEvoPoke.Caption = "Pokémon: " & Trim$(Pokemon(scrlEvolutionPoke.value).Name)
    Else
        lblEvoPoke.Caption = "Pokémon: None"
    End If

    Pokemon(EditorIndex).Evolução(EvolutionIndex).Pokemon = scrlEvolutionPoke.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEvolutionPoke_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEvolutionStone_Change()
    Dim Elemento As String

    Select Case scrlEvolutionStone.value
    Case 1
        Elemento = "Fogo"
    Case 2
        Elemento = "Água"
    Case 3
        Elemento = "Grama"
    Case 4
        Elemento = "Elétrico"
    Case 5
        Elemento = "Terrestre"
    Case 6
        Elemento = "Normal"
    Case 7
        Elemento = "Pedra"
    Case 8
        Elemento = "Voador"
    Case 9
        Elemento = "Venenoso"
    Case 10
        Elemento = "Inseto"
    Case 11
        Elemento = "Noturno"
    Case 12
        Elemento = "Fantasma"
    Case 13
        Elemento = "Psíquico"
    Case 14
        Elemento = "Dragão"
    Case 15
        Elemento = "Metálico"
    Case 16
        Elemento = "Gelo"
    Case 17
        Elemento = "Lutador"
    Case 18
        Elemento = "Fada"
    Case Else
        Elemento = "Nenhuma"
    End Select

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlEvolutionStone.value > 0 Then
        lblEvoStone.Caption = "Stone: " & Elemento
    Else
        lblEvoStone.Caption = "Stone: Nenhuma"
    End If

    Pokemon(EditorIndex).Evolução(EvolutionIndex).Pedra = scrlEvolutionStone.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEvolutionStone_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExpType_Change()
    Dim TipoString As String

    Select Case scrlExpType.value
    Case 0
        TipoString = "Rápido"
    Case 1
        TipoString = "Médio Rápido"
    Case 2
        TipoString = "Médio Lento"
    Case 3
        TipoString = "Lento"
    End Select

    lblExpType.Caption = "Exp: " & TipoString
    Pokemon(EditorIndex).ExpType = scrlExpType.value
End Sub

Private Sub scrlFelicidade_Change()
    lblFelicidade.Caption = "Felicidade: " & scrlFelicidade.value
    Pokemon(EditorIndex).HappyBase = scrlFelicidade.value
End Sub

Private Sub scrlHabilidadeIndex_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HabilidadeIndex = scrlHabilidadeIndex.value
    FrameBlank(4).Caption = "Habilidades - " & HabilidadeIndex
    scrlHabilidadeSpell.value = Pokemon(EditorIndex).Habilidades(HabilidadeIndex).Spell
    scrlHabilidadeLevel.value = Pokemon(EditorIndex).Habilidades(HabilidadeIndex).Level

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHabilidadeIndex_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHabilidadeLevel_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblHabLevel.Caption = "Level: " & scrlHabilidadeLevel.value

    Pokemon(EditorIndex).Habilidades(HabilidadeIndex).Level = scrlHabilidadeLevel.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHabilidadeSpell_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHabilidadeSpell_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlHabilidadeSpell.value > 0 Then
        lblHabSpell.Caption = "Spell: " & Trim$(Spell(scrlHabilidadeSpell.value).Name)
    Else
        lblHabSpell.Caption = "Spell: None"
    End If

    Pokemon(EditorIndex).Habilidades(HabilidadeIndex).Spell = scrlHabilidadeSpell.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHabilidadeSpell_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPokemon_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Pokemon(EditorIndex).Sprite = scrlPokemon.value
    Call EditorPokemon_BltSprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPokemon_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlStat_Change(Index As Integer)
    Dim text As String

    Select Case Index
    Case 1
        text = "Ataque: "
    Case 2
        text = "Defesa: "
    Case 3
        text = "SpAtk: "
    Case 4
        text = "Speed: "
    Case 5
        text = "SpDef: "
    End Select

    lblStat(Index).Caption = text & ScrlStat(Index).value
    Pokemon(EditorIndex).Add_Stat(Index) = ScrlStat(Index).value

End Sub

Private Sub ScrlVitals_Change(Index As Integer)
    Dim text As String

    Select Case Index
    Case 1
        text = "Hp: "
    Case 2
        text = "Mp: "
    End Select

    lblHP(Index).Caption = text & ScrlVitals(Index).value
    Pokemon(EditorIndex).Vital(Index) = ScrlVitals(Index).value

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Pokemon(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_POKEMONS Then Exit Sub

    Pokemon(EditorIndex).Desc = txtDesc.text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Pokemon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

