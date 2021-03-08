Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
    Dim filename As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
    Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
    Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
    Print #1, ""
    Close #1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, value As String)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call WritePrivateProfileString$(Header, Var, value, File)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
    Dim filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\Data Files\config.ini"

    Call PutVar(filename, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(filename, "Options", "Username", Trim$(Options.Username))
    Call PutVar(filename, "Options", "Password", Trim$(Options.Password))
    Call PutVar(filename, "Options", "SavePass", str(Options.SavePass))
    Call PutVar(filename, "Options", "IP", Options.IP)
    Call PutVar(filename, "Options", "Port", str(Options.Port))
    Call PutVar(filename, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(filename, "Options", "Music", str(Options.Music))
    Call PutVar(filename, "Options", "Sound", str(Options.sound))
    Call PutVar(filename, "Options", "Wasd", str(Options.wasd))
    Call PutVar(filename, "Options", "Debug", str(Options.Debug))
    Call PutVar(filename, "Options", "MiniMap", str(Options.MiniMap))
    Call PutVar(filename, "Options", "Quest", str(Options.Quest))

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
    Dim filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\Data Files\config.ini"

    If Not FileExist(filename, True) Then
        Options.Game_Name = "Pokémon Origins Online"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.sound = 1
        Options.wasd = 0
        Options.Debug = 0
        Options.MiniMap = 1
        Options.Quest = 1
        SaveOptions
    Else
        Options.Game_Name = GetVar(filename, "Options", "Game_Name")
        Options.Username = GetVar(filename, "Options", "Username")
        Options.Password = GetVar(filename, "Options", "Password")
        Options.SavePass = Val(GetVar(filename, "Options", "SavePass"))
        Options.IP = GetVar(filename, "Options", "IP")
        Options.Port = Val(GetVar(filename, "Options", "Port"))
        Options.MenuMusic = GetVar(filename, "Options", "MenuMusic")
        Options.Music = GetVar(filename, "Options", "Music")
        Options.sound = GetVar(filename, "Options", "Sound")
        Options.wasd = GetVar(filename, "Options", "Wasd")
        Options.Debug = GetVar(filename, "Options", "Debug")
        Options.MiniMap = GetVar(filename, "Options", "MiniMap")
        Options.Quest = GetVar(filename, "Options", "Quest")
    End If

    ' show in GUI
    If Options.Music = 0 Then
        frmMain.optMOff.value = True
    Else
        frmMain.optMOn.value = True
    End If

    If Options.sound = 0 Then
        frmMain.optSOff.value = True
    Else
        frmMain.optSOn.value = True
    End If
    
    If Options.wasd = 0 Then
        frmMain.optWOff.value = True
    Else
        frmMain.optWOn.Visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim f As Long
    Dim X As Long
    Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Map.Name
    Put #f, , Map.Music
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY
    Put #f, , Map.Weather
    Put #f, , Map.Intensity

    For X = 1 To 2
        Put #f, , Map.LevelPoke(X)
    Next

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #f, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #f, , Map.Npc(X)
    Next

    Close #f

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
    Dim filename As String
    Dim f As Long
    Dim X As Long
    Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Map.Name
    Get #f, , Map.Music
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
    Get #f, , Map.Weather
    Get #f, , Map.Intensity

    For X = 1 To 2
        Get #f, , Map.LevelPoke(X)
    Next

    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #f, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #f, , Map.Npc(X)
    Next

    Close #f
    ClearTempTile

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "\tilesets\" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend

    If NumTileSets = 0 Then Exit Sub

    ReDim DDS_Tileset(1 To NumTileSets)
    ReDim DDSD_Tileset(1 To NumTileSets)
    ReDim TilesetTimer(1 To NumTileSets)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend

    If NumCharacters = 0 Then Exit Sub

    ReDim DDS_Character(1 To NumCharacters)
    ReDim DDSD_Character(1 To NumCharacters)
    ReDim CharacterTimer(1 To NumCharacters)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend

    If NumPaperdolls = 0 Then Exit Sub

    ReDim DDS_Paperdoll(1 To NumPaperdolls)
    ReDim DDSD_Paperdoll(1 To NumPaperdolls)
    ReDim PaperdollTimer(1 To NumPaperdolls)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend

    If NumAnimations = 0 Then Exit Sub

    ReDim DDS_Animation(1 To NumAnimations)
    ReDim DDSD_Animation(1 To NumAnimations)
    ReDim AnimationTimer(1 To NumAnimations)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)
        numitems = numitems + 1
        i = i + 1
    Wend

    If numitems = 0 Then Exit Sub

    ReDim DDS_Item(1 To numitems)
    ReDim DDSD_Item(1 To numitems)
    ReDim ItemTimer(1 To numitems)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Resources\" & i & GFX_EXT)
        NumResources = NumResources + 1
        i = i + 1
    Wend

    If NumResources = 0 Then Exit Sub

    ReDim DDS_Resource(1 To NumResources)
    ReDim DDSD_Resource(1 To NumResources)
    ReDim ResourceTimer(1 To NumResources)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "SpellIcons\" & i & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend

    If NumSpellIcons = 0 Then Exit Sub

    ReDim DDS_SpellIcon(1 To NumSpellIcons)
    ReDim DDSD_SpellIcon(1 To NumSpellIcons)
    ReDim SpellIconTimer(1 To NumSpellIcons)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        NumFaces = NumFaces + 1
        i = i + 1
    Wend

    If NumFaces = 0 Then Exit Sub

    ReDim DDS_Face(1 To NumFaces)
    ReDim DDSD_Face(1 To NumFaces)
    ReDim FaceTimer(1 To NumFaces)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFacesShiny()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Faces\Shiny\" & i & GFX_EXT)
        NumFacesShiny = NumFacesShiny + 1
        i = i + 1
    Wend

    If NumFaces = 0 Then Exit Sub

    ReDim DDS_FaceShiny(1 To NumFacesShiny)
    ReDim DDSD_FaceShiny(1 To NumFacesShiny)
    ReDim FaceShinyTimer(1 To NumFacesShiny)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFacesShiny", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPokeIcons()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "PokeIcon\" & i & GFX_EXT)
        NumPokeIcons = NumPokeIcons + 1
        i = i + 1
    Wend

    If NumPokeIcons = 0 Then Exit Sub

    ReDim DDS_PokeIcons(1 To NumPokeIcons)
    ReDim DDSD_PokeIcons(1 To NumPokeIcons)
    ReDim PokeIconTimer(1 To NumPokeIcons)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPokeIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPokeIconShiny()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "PokeIcon\Shiny\" & i & GFX_EXT)
        NumPokeIconShiny = NumPokeIconShiny + 1
        i = i + 1
    Wend

    If NumPokeIconShiny = 0 Then Exit Sub

    ReDim DDS_PokeIconShiny(1 To NumPokeIconShiny)
    ReDim DDSD_PokeIconShiny(1 To NumPokeIconShiny)
    ReDim PokeIconShinyTimer(1 To NumPokeIconShiny)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPokeIconshiny", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckHairNum()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "characters\Cabelos\" & i & GFX_EXT)
        HairNum = HairNum + 1
        i = i + 1
    Wend

    If HairNum = 0 Then Exit Sub

    ReDim DDS_Hair(1 To HairNum)
    ReDim DDSD_Hair(1 To HairNum)
    ReDim HairTimer(1 To HairNum)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckHairNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).sound = "None."

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).sound = "None."

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If Index = 0 Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).Exp

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Exp = Exp

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerHonra(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerHonra = Player(Index).Honra

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerHonra", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerHonra(ByVal Index As Long, ByVal Honra As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Honra = Honra

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerHonra", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal value As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function

    GetPlayerMaxVital = Player(Index).MaxVital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If Index = 0 Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal value As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Index = 0 Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    Player(Index).Stat(Stat) = value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).num

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal itemNum As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).num = itemNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVisuais(ByVal Index As Long, ByVal ViSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If ViSlot = 0 Then Exit Function
    GetPlayerVisuais = Player(ViSlot).Visuais(ViSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVisuais", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVisuais(ByVal Index As Long, ByVal ViSlot As Long, ByVal ViNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(ViSlot).Visuais(ViSlot) = ViNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVisuais", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerTeleport(ByVal Index As Long, ByVal TpSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If TpSlot = 0 Then Exit Function
    GetPlayerTeleport = Player(TpSlot).Teleport(TpSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerTeleport", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerTeleport(ByVal Index As Long, ByVal TpSlot As Long, ByVal TpNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(TpSlot).Teleport(TpSlot) = TpNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerTeleport", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).value

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).value = ItemValue

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = InvNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'######################## Pokémon Info ################################

'PokeInfo Pokemon
Function GetPlayerInvItemPokeInfoPokemon(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Or invslot > MAX_INV Then Exit Function

    GetPlayerInvItemPokeInfoPokemon = PlayerInv(invslot).PokeInfo.Pokemon
End Function

Sub SetPlayerInvItemPokeInfoPokemon(ByVal Index As Long, ByVal invslot As Long, ByVal PokeNum As Long)
    PlayerInv(invslot).PokeInfo.Pokemon = PokeNum
End Sub

'PokeInfo Pokéball
Function GetPlayerInvItemPokeInfoPokeball(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoPokeball = PlayerInv(invslot).PokeInfo.Pokeball
End Function

Sub SetPlayerInvItemPokeInfoPokeball(ByVal Index As Long, ByVal invslot As Long, ByVal PokeballNum As Long)
    PlayerInv(invslot).PokeInfo.Pokeball = PokeballNum
End Sub

'PokeInfo Level
Function GetPlayerInvItemPokeInfoLevel(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoLevel = PlayerInv(invslot).PokeInfo.Level
End Function

Sub SetPlayerInvItemPokeInfoLevel(ByVal Index As Long, ByVal invslot As Long, ByVal Level As Long)
    PlayerInv(invslot).PokeInfo.Level = Level
End Sub

'PokeInfo Exp
Function GetPlayerInvItemPokeInfoExp(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoExp = PlayerInv(invslot).PokeInfo.Exp
End Function

Sub SetPlayerInvItemPokeInfoExp(ByVal Index As Long, ByVal invslot As Long, ByVal Exp As Long)
    PlayerInv(invslot).PokeInfo.Exp = Exp
End Sub

'PokeInfo MaxVital
Function GetPlayerInvItemPokeInfoMaxVital(ByVal Index As Long, ByVal invslot As Long, ByVal VitalType As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoMaxVital = PlayerInv(invslot).PokeInfo.MaxVital(VitalType)
End Function

Sub SetPlayerInvItemPokeInfoMaxVital(ByVal Index As Long, ByVal invslot As Long, ByVal MaxVital As Long, ByVal VitalType As Long)
    PlayerInv(invslot).PokeInfo.MaxVital(VitalType) = MaxVital
End Sub

'PokeInfo Vital
Function GetPlayerInvItemPokeInfoVital(ByVal Index As Long, ByVal invslot As Long, ByVal VitalType As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoVital = PlayerInv(invslot).PokeInfo.Vital(VitalType)
End Function

Sub SetPlayerInvItemPokeInfoVital(ByVal Index As Long, ByVal invslot As Long, ByVal Vital As Long, ByVal VitalType As Long)
    PlayerInv(invslot).PokeInfo.Vital(VitalType) = Vital
End Sub

'PokeInfo Spells
Function GetPlayerInvItemPokeInfoSpell(ByVal Index As Long, ByVal invslot As Long, ByVal spellslot As Byte) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoSpell = PlayerInv(invslot).PokeInfo.Spells(spellslot)
End Function

Sub SetPlayerInvItemPokeInfoSpell(ByVal Index As Long, ByVal invslot As Long, ByVal SpellNum As Long, ByVal spellslot As Long)
    PlayerInv(invslot).PokeInfo.Spells(spellslot) = SpellNum
End Sub

'PokeInfo Status
Function GetPlayerInvItemPokeInfoStat(ByVal Index As Long, ByVal invslot As Long, ByVal StatNum As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemPokeInfoStat = PlayerInv(invslot).PokeInfo.Stat(StatNum)
End Function

Sub SetPlayerInvItemPokeInfoStat(ByVal Index As Long, ByVal invslot As Long, ByVal StatNum As Long, ByVal StatValue As Long)
    If StatNum = 0 Or StatNum > Stats.Stat_Count - 1 Then Exit Sub
    PlayerInv(invslot).PokeInfo.Stat(StatNum) = StatValue
End Sub

'PokeInfo Stats negativo
Function GetPlayerInvItemNgt(ByVal Index As Long, ByVal invslot As Long, ByVal NgtNum As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemNgt = PlayerInv(invslot).PokeInfo.Negatives(NgtNum)
End Function

Sub SetPlayerInvItemNgt(ByVal Index As Long, ByVal invslot As Long, ByVal NgtNum As Long, ByVal NgtValue As Long)
    If NgtNum = 0 Or NgtNum > MAX_NEGATIVES Then Exit Sub
    PlayerInv(invslot).PokeInfo.Negatives(NgtNum) = NgtValue
End Sub

'PokeInfo Felicidade
Function GetPlayerInvItemFelicidade(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemFelicidade = PlayerInv(invslot).PokeInfo.Felicidade
End Function

Sub SetPlayerInvItemFelicidade(ByVal Index As Long, ByVal invslot As Long, ByVal Felicidade As Long)
    PlayerInv(invslot).PokeInfo.Felicidade = Felicidade
End Sub

'PokeInfo Sexo
Function GetPlayerInvItemSexo(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemSexo = PlayerInv(invslot).PokeInfo.Sexo
End Function

Sub SetPlayerInvItemSexo(ByVal Index As Long, ByVal invslot As Long, ByVal Sexo As Long)
    PlayerInv(invslot).PokeInfo.Sexo = Sexo
End Sub

'PokeInfo Shiny
Function GetPlayerInvItemShiny(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemShiny = PlayerInv(invslot).PokeInfo.Shiny
End Function

Sub SetPlayerInvItemShiny(ByVal Index As Long, ByVal invslot As Long, ByVal Shiny As Long)
    PlayerInv(invslot).PokeInfo.Shiny = Shiny
End Sub

'PokeInfo Berry
Function GetPlayerInvItemBerry(ByVal Index As Long, ByVal invslot As Long, ByVal BerryStat As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function

    GetPlayerInvItemBerry = PlayerInv(invslot).PokeInfo.Berry(BerryStat)
End Function

Sub SetPlayerInvItemBerry(ByVal Index As Long, ByVal invslot As Long, ByVal BerryStat As Long, ByVal Valor As Long)
'If BerryStat = 0 Or BerryStat > MAX_BERRYS Then Exit Sub
    PlayerInv(invslot).PokeInfo.Berry(BerryStat) = Valor
End Sub

'############################Poke Info Equipment##################################################

Function GetPlayerEquipmentPokeInfoPokemon(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoPokemon = Player(Index).EquipPokeInfo(EquipmentSlot).Pokemon
End Function

Sub SetPlayerEquipmentPokeInfoPokemon(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Pokemon = InvNum
End Sub

'PokeInfo Pokeball
Function GetPlayerEquipmentPokeInfoPokeball(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoPokeball = Player(Index).EquipPokeInfo(EquipmentSlot).Pokeball
End Function

Sub SetPlayerEquipmentPokeInfoPokeball(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Pokeball = InvNum
End Sub

'PokeInfo Level
Function GetPlayerEquipmentPokeInfoLevel(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoLevel = Player(Index).EquipPokeInfo(EquipmentSlot).Level
End Function

Sub SetPlayerEquipmentPokeInfoLevel(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Level = InvNum
End Sub

'PokeInfo Exp
Function GetPlayerEquipmentPokeInfoExp(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoExp = Player(Index).EquipPokeInfo(EquipmentSlot).Exp
End Function

Sub SetPlayerEquipmentPokeInfoExp(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Exp = InvNum
End Sub

'PokeInfo Vitals
Function GetPlayerEquipmentPokeInfoVital(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoVital = Player(Index).EquipPokeInfo(EquipmentSlot).Vital(VitalType)
End Function

Sub SetPlayerEquipmentPokeInfoVital(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Vital(VitalType) = InvNum
End Sub

'PokeInfo MaxVital
Function GetPlayerEquipmentPokeInfoMaxVital(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoMaxVital = Player(Index).EquipPokeInfo(EquipmentSlot).MaxVital(VitalType)
End Function

Sub SetPlayerEquipmentPokeInfoMaxVital(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).MaxVital(VitalType) = InvNum
End Sub

'PokeInfo Stat
Function GetPlayerEquipmentPokeInfoStat(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal StatNum As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoStat = Player(Index).EquipPokeInfo(EquipmentSlot).Stat(StatNum)
End Function

Sub SetPlayerEquipmentPokeInfoStat(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal StatNum As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Stat(StatNum) = InvNum
End Sub

'PokeInfo Spells
Function GetPlayerEquipmentPokeInfoSpell(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal SpellNum As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoSpell = Player(Index).EquipPokeInfo(EquipmentSlot).Spells(SpellNum)
End Function

Sub SetPlayerEquipmentPokeInfoSpell(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal SpellNum As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Spells(SpellNum) = InvNum
End Sub
'PokeInfo Negatives
Function GetPlayerEquipmentNgt(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal NgtNum As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentNgt = Player(Index).EquipPokeInfo(EquipmentSlot).Negatives(NgtNum)
End Function

Sub SetPlayerEquipmentNgt(ByVal Index As Long, ByVal NgtNum As Long, ByVal EquipmentSlot As Equipment, ByVal NgtValue As Long)
    If NgtNum = 0 Then Exit Sub
    Player(Index).EquipPokeInfo(EquipmentSlot).Negatives(NgtNum) = NgtValue
End Sub

'PokeInfo Sexo
Function GetPlayerEquipmentSexo(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentSexo = Player(Index).EquipPokeInfo(EquipmentSlot).Sexo
End Function

Sub SetPlayerEquipmentSexo(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal Sexo As Long)
    If EquipmentSlot = 0 Then Exit Sub
    Player(Index).EquipPokeInfo(EquipmentSlot).Sexo = Sexo
End Sub

'PokeInfo Felicidade
Function GetPlayerEquipmentFelicidade(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentFelicidade = Player(Index).EquipPokeInfo(EquipmentSlot).Felicidade
End Function

Sub SetPlayerEquipmentFelicidade(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal Felicidade As Long)
    If EquipmentSlot = 0 Then Exit Sub
    Player(Index).EquipPokeInfo(EquipmentSlot).Felicidade = Felicidade
End Sub

'PokeInfo Shiny
Function GetPlayerEquipmentShiny(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentShiny = Player(Index).EquipPokeInfo(EquipmentSlot).Shiny
End Function

Sub SetPlayerEquipmentShiny(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal Shiny As Long)
    If EquipmentSlot = 0 Then Exit Sub
    Player(Index).EquipPokeInfo(EquipmentSlot).Shiny = Shiny
End Sub

'PokeInfo Berry
Function GetPlayerEquipmentBerry(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal BerryStat As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentBerry = Player(Index).EquipPokeInfo(EquipmentSlot).Berry(BerryStat)
End Function

Sub SetPlayerEquipmentBerry(ByVal Index As Long, ByVal Valor As Long, ByVal EquipmentSlot As Equipment, ByVal BerryStat As Long)
    If EquipmentSlot = 0 Then Exit Sub
    Player(Index).EquipPokeInfo(EquipmentSlot).Berry(BerryStat) = Valor
End Sub

'################
'#####QUESTS#####
'################

Sub ClearQuest(ByVal Index As Long)
    Dim i As Byte, X As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = ""
    Quest(Index).Description = ""

    For i = 1 To MAX_QUEST_TASKS
        For X = 1 To 3
            Quest(Index).Task(i).Message(X) = ""
        Next

        Quest(Index).Task(i).num = 1
        Quest(Index).Task(i).value = 1
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearQuests()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearQuests", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function FindQuestSlot(ByVal Index As Long, ByVal QuestNum As Integer) As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(Index) Then Exit Function

    ' Find the quest
    For i = 1 To MAX_QUESTS
        If Player(MyIndex).Quests(i).status = QuestNum Then
            FindQuestSlot = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindQuestSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function HasItem(ByVal Index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function


Function GetPlayerFlying(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerFlying = Player(Index).Flying

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerFlying", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerFlying(ByVal Index As Long, ByVal Flying As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Flying = Flying

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerFlying", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SetPlayerOrg(ByVal Index As Long, ByVal ORG As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).ORG = ORG

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerOrg", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerOrg(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerOrg = Player(Index).ORG

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerOrg", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
