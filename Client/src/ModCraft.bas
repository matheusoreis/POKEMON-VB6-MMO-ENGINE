Attribute VB_Name = "ModPokemon"
'/////////////////////////////////////////////////////////////////////
'/////////// Pokemon System - Developed by Alifer ////////////////////
'/////////////////////////////////////////////////////////////////////

'Constants
Public Const MAX_POKEMONS As Long = 999
Public Const EDITOR_POKEMON As Byte = 7

'Publics
Public Pokemon(1 To MAX_POKEMONS) As PokemonRec

'Modificado?
Public Pokemon_Changed(1 To MAX_POKEMONS) As Boolean

Private Type PokeHabRec
    Spell As Long
    Level As Long
End Type

Private Type PokeEvoRec
    Pokemon As Long
    Level As Long
    Pedra As Byte
End Type

Private Type PokemonRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sprite As Long
    Tipo(1 To 2) As Long
    Evolução(1 To 8) As PokeEvoRec
    Habilidades(1 To 20) As PokeHabRec
    AnimAttack As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    ExpType As Byte
    ControlSex As Byte    '0 = 100% Female, 100% = 100% Male
    AnimFrame(1 To 2) As Byte
    NotEvo As Byte
    HappyBase As Byte
    ExpBase As Integer
    EggTime As Integer
    CRate As Byte
End Type

' ////////////
' // Editor //
' ////////////

Public Sub PokemonEditorInit()
    Dim i As Long

    EditorIndex = frmEditor_Pokemon.lstIndex.ListIndex + 1

    With frmEditor_Pokemon
        .txtName.text = Trim$(Pokemon(EditorIndex).Name)
        .txtDesc.text = Trim$(Pokemon(EditorIndex).Desc)
        .scrlPokemon = Pokemon(EditorIndex).Sprite
        For i = 1 To 2
            .cmbTipo(i).ListIndex = Pokemon(EditorIndex).Tipo(i)
        Next
        .scrlEvolutionIndex = 1
        .scrlHabilidadeIndex = 1
        .scrlEvolutionPoke = Pokemon(EditorIndex).Evolução(1).Pokemon
        .scrlEvolutionLevel = Pokemon(EditorIndex).Evolução(1).Level
        .scrlEvolutionStone = Pokemon(EditorIndex).Evolução(1).Pedra
        .scrlHabilidadeSpell = Pokemon(EditorIndex).Habilidades(1).Spell
        .scrlHabilidadeLevel = Pokemon(EditorIndex).Habilidades(1).Level
        .scrlAnimAttack = Pokemon(EditorIndex).AnimAttack
        .scrlControlSex(0).value = Pokemon(EditorIndex).ControlSex
        .scrlAnimFrame(0).value = Pokemon(EditorIndex).AnimFrame(1)
        .scrlAnimFrame(1).value = Pokemon(EditorIndex).AnimFrame(2)
        .chkEvo.value = Pokemon(EditorIndex).NotEvo
        .scrlFelicidade = Pokemon(EditorIndex).HappyBase
        .scrlBaseExp = Pokemon(EditorIndex).ExpBase
        .scrlEggTime = Pokemon(EditorIndex).EggTime
        .scrlExpType.value = Pokemon(EditorIndex).ExpType

        For i = 1 To Stats.Stat_Count - 1
            .ScrlStat(i).value = Pokemon(EditorIndex).Add_Stat(i)
        Next

        For i = 1 To Vitals.Vital_Count - 1
            .ScrlVitals(i).value = Pokemon(EditorIndex).Vital(i)
        Next

    End With
    Call EditorPokemon_BltSprite

    Pokemon_Changed(EditorIndex) = True
End Sub

Public Sub PokemonEditorOk()
    Dim i As Long

    For i = 1 To MAX_POKEMONS
        If Pokemon_Changed(i) Then
            Call SendSavePokemon(i)
        End If
    Next

    Unload frmEditor_Pokemon
    Editor = 0
    ClearChanged_Pokemon

End Sub

Public Sub PokemonEditorCancel()
    Editor = 0
    Unload frmEditor_Pokemon
    ClearChanged_Pokemon
    SendRequestPokemon
End Sub

Public Sub ClearChanged_Pokemon()
    ZeroMemory Pokemon_Changed(1), MAX_POKEMONS * 2    ' 2 = boolean length
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearPokemon(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Pokemon(Index)), LenB(Pokemon(Index)))
    Pokemon(Index).Name = vbNullString
End Sub

Sub ClearPokemons()
    Dim i As Long
    For i = 1 To MAX_POKEMONS
        Call ClearPokemon(i)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Public Sub SendRequestEditPokemon()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditPokemon
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendSavePokemon(ByVal PokemonNum As Long)
    Dim Buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte

    Set Buffer = New clsBuffer
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(PokemonNum)), PokemonSize
    Buffer.WriteLong CSavePokemon
    Buffer.WriteLong PokemonNum
    Buffer.WriteBytes PokemonData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendRequestPokemon()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPokemon
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendSelectPokeInicial()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSelectPoke
    Buffer.WriteByte SelectPokeInicial
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

'//////////////////////////////////
'/////// Receber Pacotes //////////
'//////////////////////////////////

Public Sub HandlePokemonEditor()
    Dim i As Long

    frmEditor_Pokemon.Visible = True
    With frmEditor_Pokemon
        Editor = EDITOR_POKEMON
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_POKEMONS
            .lstIndex.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        PokemonEditorInit
    End With

End Sub

Public Sub HandleUpdatePokemon(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    ' Update the Pokemon
    PokemonSize = LenB(Pokemon(n))
    ReDim PokemonData(PokemonSize - 1)
    PokemonData = Buffer.ReadBytes(PokemonSize)
    CopyMemory ByVal VarPtr(Pokemon(n)), ByVal VarPtr(PokemonData(0)), PokemonSize
    Set Buffer = Nothing
End Sub

Public Sub HandlePlayerPokedex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    For i = 1 To MAX_POKEMONS
        Player(Index).Pokedex(i) = Buffer.ReadLong
    Next

    frmMain.ListPokes.Clear

    For i = 1 To 251    'MAX_POKEMONS
        If Player(Index).Pokedex(i) = 1 Then
            frmMain.ListPokes.AddItem i & "-" & Trim$(Pokemon(i).Name)

        Else
            frmMain.ListPokes.AddItem i & "-" & "???"
        End If
    Next    '???

    Set Buffer = Nothing

End Sub
