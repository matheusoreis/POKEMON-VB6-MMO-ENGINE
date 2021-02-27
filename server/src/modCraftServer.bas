Attribute VB_Name = "modPokemonServer"
'/////////////////////////////////////////////////////////////////////
'///////////// Pokémon System - Developed by Alifer //////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Constants
Public Const MAX_POKEMONS As Long = 999
Public Const EDITOR_POKEMON As Byte = 7

'Publics
Public Pokemon(1 To MAX_POKEMONS) As PokemonRec

'Variaveis do Editor Pokémon
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
ControlSex As Byte '0 = 100% Female, 100% = 100% Male
AnimFrame(1 To 2) As Byte
NotEvo As Byte
HappyBase As Byte
ExpBase As Integer
EggTime As Integer
CRate As Byte
End Type

' **************
' ** Pokémons **
' **************

Sub SavePokemon(ByVal PokemonNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\pokemons\pokemon" & PokemonNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Pokemon(PokemonNum)
    Close #F
End Sub

Sub SavePokemons()
    Dim i As Long
    Call SetStatus("Saving Pokemons... ")

    For i = 1 To MAX_POKEMONS
        Call SavePokemon(i)
    Next

End Sub

Sub LoadPokemons()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckPokemons

    For i = 1 To MAX_POKEMONS
        filename = App.Path & "\data\pokemons\pokemon" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Pokemon(i)
        Close #F
    Next

End Sub

Sub CheckPokemons()
    Dim i As Long

    For i = 1 To MAX_POKEMONS

        If Not FileExist("\Data\pokemons\pokemon" & i & ".dat") Then
            Call SavePokemon(i)
        End If

    Next

End Sub

Sub ClearPokemon(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Pokemon(Index)), LenB(Pokemon(Index)))
    Pokemon(Index).Name = vbNullString
    Pokemon(Index).Desc = vbNullString
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

Sub SendPokemonS(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_POKEMONS
        If LenB(Trim$(Pokemon(i).Name)) > 0 Then
            Call SendUpdatePokemonTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdatePokemonToAll(ByVal PokemonNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte
    Set Buffer = New clsBuffer
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(PokemonNum)), PokemonSize
    Buffer.WriteLong SUpdatePokemon
    Buffer.WriteLong PokemonNum
    Buffer.WriteBytes PokemonData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdatePokemonTo(ByVal Index As Long, ByVal PokemonNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte
    Set Buffer = New clsBuffer
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(PokemonNum)), PokemonSize
    Buffer.WriteLong SUpdatePokemon
    Buffer.WriteLong PokemonNum
    Buffer.WriteBytes PokemonData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerPokedex(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerPokedex
    For i = 1 To 251
    Buffer.WriteLong Player(Index).Pokedex(i)
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

'//////////////////////////////////
'/////// Receber Pacotes //////////
'//////////////////////////////////

Sub HandleRequestEditPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPokemonEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSavePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_POKEMONS Then
        Exit Sub
    End If
    
    ' Update the Quest
    PokemonSize = LenB(Pokemon(n))
    ReDim PokemonData(PokemonSize - 1)
    PokemonData = Buffer.ReadBytes(PokemonSize)
    CopyMemory ByVal VarPtr(Pokemon(n)), ByVal VarPtr(PokemonData(0)), PokemonSize
    Set Buffer = Nothing

    ' Save it
    Call SendUpdatePokemonToAll(n)
    Call SavePokemon(n)
    Call AddLog(GetPlayerName(Index) & " salvou Pokemon #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerPokedex Index
End Sub
