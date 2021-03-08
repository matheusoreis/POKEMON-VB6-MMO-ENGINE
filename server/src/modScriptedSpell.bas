Attribute VB_Name = "modScriptedSpell"
Sub ScriptedSpell(ByVal Index As Long, ByVal Script As Long, ByVal SpellNum As Long)
Dim i, MapNum As Long
Dim Dir As Long

MapNum = GetPlayerMap(Index)
X = GetPlayerX(Index)
Y = GetPlayerY(Index)
Dir = GetPlayerDir(Index)

    Select Case Script
    Case 1 'Iluminar
        If Map(MapNum).Moral = 4 Then
            Player(Index).PokeLight = True
            SendPlayerData Index
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then Call PlayerMsg(Index, Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " usou a habilidade Iluminar e está tudo visivel agora.", BrightGreen)
        End If
    Case 2

    Case Else
        Call PlayerMsg(Index, "nada", Red)
    End Select
End Sub

