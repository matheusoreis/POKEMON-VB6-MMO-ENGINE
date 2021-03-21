Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
Dim filename As String, IP As String, F As Long, i As Long

    ' Add banned to the player's index
    Player(BanPlayerIndex).isBanned = 1
    SavePlayer BanPlayerIndex

    ' IP banning
    filename = App.Path & "\data\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)
    F = FreeFile
    Open filename For Append As #F
        Print #F, IP
    Close #F
    
    ' Tell them they're banned
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " foi banido do " & Options.Game_Name & ".", White)
    Call AddLog(GetPlayerName(BanPlayerIndex) & " foi banido.", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "Você foi banido.")
End Sub

Public Function isBanned_IP(ByVal IP As String) As Boolean
Dim filename As String, fIP As String, F As Long
    
    filename = App.Path & "\data\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String, ByVal RecoveryKey As String, ByVal Email As String)
    Dim i As Long
    
    ClearPlayer Index
    
    Player(Index).Login = Name
    Player(Index).Password = Password
    Player(Index).SecondPass = RecoveryKey
    Player(Index).Email = Email
    
    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim S As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, S

        If Trim$(LCase$(S)) <> Trim$(LCase$(Name)) Then
            Print #f2, S
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean

    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal Cabelo As Byte)
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean
    Dim i As Long

    If LenB(Trim$(Player(Index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(Index).Name = Name
        Player(Index).Sex = Sex
        Player(Index).Class = ClassNum
        
        If Player(Index).Sex = SEX_MALE Then
            Player(Index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(Index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If
        
        Player(Index).Cabelo = Class(ClassNum).Cabelos(Cabelo)
        Player(Index).Level = 1
        Player(Index).MySprite = Player(Index).Sprite

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(Index).Dir = DIR_DOWN
        Player(Index).Map = START_MAP
        Player(Index).x = START_X
        Player(Index).Y = START_Y
        Player(Index).Dir = DIR_DOWN
        Player(Index).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        Player(Index).TPX = 0
        Player(Index).TPY = 0
        Player(Index).TPDir = 0
        Player(Index).PokeInicial = 1
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(Index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim S As String
    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, S

        If Trim$(LCase$(S)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Player(Index)
    Close #F
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long
    Call ClearPlayer(Index)
    filename = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index)
    Close #F
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Class = 1

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
           
        tmpSprite = GetVar(filename, "CLASS" & i, "Cabelos")
        tmpArray() = Split(tmpSprite, ",")
        ReDim Class(i).Cabelos(0 To UBound(tmpArray))
        For n = 0 To UBound(tmpArray)
            Class(i).Cabelos(n) = Val(tmpArray(n))
        Next
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_PLAYER_SPELLS Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    
    filename = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.Path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.Path & "\data\items\item" & ItemNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\spells\spells" & SpellNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(SpellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Npc(NpcNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Npc(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
    Npc(Index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim Y As Long
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(MapNum).Name
    Put #F, , Map(MapNum).Music
    Put #F, , Map(MapNum).Revision
    Put #F, , Map(MapNum).Moral
    Put #F, , Map(MapNum).Up
    Put #F, , Map(MapNum).Down
    Put #F, , Map(MapNum).Left
    Put #F, , Map(MapNum).Right
    Put #F, , Map(MapNum).BootMap
    Put #F, , Map(MapNum).BootX
    Put #F, , Map(MapNum).BootY
    Put #F, , Map(MapNum).MaxX
    Put #F, , Map(MapNum).MaxY
    Put #F, , Map(MapNum).Weather
    Put #F, , Map(MapNum).Intensity
    
    For x = 1 To 2
        Put #F, , Map(MapNum).LevelPoke(x)
    Next
    
    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            Put #F, , Map(MapNum).Tile(x, Y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(MapNum).Npc(x)
    Next
    Close #F
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim Y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        Get #F, , Map(i).Weather
        Get #F, , Map(i).Intensity
        
        For x = 1 To 2
            Get #F, , Map(i).LevelPoke(x)
        Next
        
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For Y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, Y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).Npc(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next

        Close #F
        
        ClearTempTile i
        CacheResources i
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
    MapItem(MapNum, Index).Num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).x = 0
    MapItem(MapNum, Index).Y = 0
    MapItem(MapNum, Index).canDespawn = False
    MapItem(MapNum, Index).despawnTimer = 0
    MapItem(MapNum, Index).PokeInfo.Pokemon = 0
    MapItem(MapNum, Index).PokeInfo.Pokeball = 0
    MapItem(MapNum, Index).PokeInfo.Level = 0
    MapItem(MapNum, Index).PokeInfo.EXP = 0
    MapItem(MapNum, Index).PokeInfo.Felicidade = 0
    MapItem(MapNum, Index).PokeInfo.Sexo = 0
    MapItem(MapNum, Index).PokeInfo.Shiny = 0
    
    For i = 1 To Vitals.Vital_Count - 1
        MapItem(MapNum, Index).PokeInfo.Vital(i) = 0
        MapItem(MapNum, Index).PokeInfo.MaxVital(i) = 0
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        MapItem(MapNum, Index).PokeInfo.Stat(i) = 0
    Next
    
    For i = 1 To MAX_POKE_SPELL
        MapItem(MapNum, Index).PokeInfo.Spells(i) = 0
    Next
    
    For i = 1 To MAX_NEGATIVES
        MapItem(MapNum, Index).PokeInfo.Negatives(i) = 0
    Next
    
    For i = 1 To MAX_BERRYS
        MapItem(MapNum, Index).PokeInfo.Berry(i) = 0
    Next
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, Y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(Index)), LenB(MapNpc(MapNum).Npc(Index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, Y)
        Next
    Next

End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.Path & "\data\banks\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(Index)
    Close #F
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    Call ClearBank(Index)

    filename = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(Index)
    Close #F

End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
End Sub

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub

'###########################
'######### QUESTS ##########
'###########################

Sub SaveQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next

End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Quest(QuestNum)
    Close #F
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Quest(i)
        Close #F
    Next
End Sub

Sub CheckQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS

        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If

    Next
End Sub

Sub ClearQuest(ByVal Index As Long)
    Dim i As Byte, x As Byte

    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = ""
    Quest(Index).Description = ""
    
    For i = 1 To MAX_QUEST_TASKS
        For x = 1 To 3
            Quest(Index).Task(i).message(x) = ""
        Next
    
        Quest(Index).Task(i).Num = 1
        Quest(Index).Task(i).Value = 1
    Next
End Sub

Sub ClearQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

Sub SaveLeilão(ByVal LeilãoNum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\data\leilão\" & LeilãoNum & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Leilao(LeilãoNum)
    Close #F
End Sub

Sub LoadLeilão(ByVal LeilãoNum As Long)
    Dim filename As String
    Dim F As Long
    Call ClearLeilão(LeilãoNum)
    filename = App.Path & "\data\leilão\" & LeilãoNum & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Leilao(LeilãoNum)
    Close #F
End Sub

Sub ClearLeilão(ByVal LeilãoNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Leilao(LeilãoNum)), LenB(Leilao(LeilãoNum)))
    Set Buffer = New clsBuffer
    
    Leilao(LeilãoNum).Vendedor = vbNullString
    Leilao(LeilãoNum).Tipo = 0
    Leilao(LeilãoNum).ItemNum = 0
    Leilao(LeilãoNum).Tempo = 0
    Leilao(LeilãoNum).Price = 0
    
End Sub

Sub SavePendencia(ByVal PendNum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\data\pendencias\" & PendNum & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Pendencia(PendNum)
    Close #F
End Sub

Sub LoadPendencia(ByVal PendNum As Long)
    Dim filename As String
    Dim F As Long
    Call ClearPendencia(PendNum)
    filename = App.Path & "\data\pendencias\" & PendNum & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Pendencia(PendNum)
    Close #F
End Sub

Sub ClearPendencia(ByVal PendNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Pendencia(PendNum)), LenB(Pendencia(PendNum)))
    Set Buffer = New clsBuffer
    
    Pendencia(PendNum).Vendedor = vbNullString
    Pendencia(PendNum).Tipo = 0
    Pendencia(PendNum).ItemNum = 0
    Pendencia(PendNum).Price = 0
End Sub

Public Sub SaveRankLevel()
Dim filename As String, i As Byte

    filename = App.Path & "\data\ranklevel.ini"
    
    For i = 1 To MAX_RANKS
        PutVar filename, "RANK", "Name" & i, Trim$(RankLevel(i).Name)
        PutVar filename, "RANK", "Level" & i, Val(RankLevel(i).Level)
        PutVar filename, "RANK", "PokeNum" & i, Val(RankLevel(i).PokeNum)
    Next


End Sub

Public Sub LoadRankLevel()
Dim filename As String
Dim i As Byte

    filename = App.Path & "\data\ranklevel.ini"
    
    If Not FileExist(filename, True) Then Exit Sub
    
    For i = 1 To MAX_RANKS
        RankLevel(i).Name = GetVar(filename, "RANK", "Name" & i)
        RankLevel(i).Level = Val(GetVar(filename, "RANK", "Level" & i))
        RankLevel(i).PokeNum = Val(GetVar(filename, "RANK", "PokeNum" & i))
    Next
    
End Sub

Public Function IsValidEmail(strEmail As String) As Boolean
   Dim names, Name, i, C
   IsValidEmail = True

   names = Split(strEmail, "@")

  If UBound(names) <> 1 Then
  IsValidEmail = False
  Exit Function
  End If

  For Each Name In names

  If Len(Name) <= 0 Then
  IsValidEmail = False
  Exit Function
  End If

  For i = 1 To Len(Name)
  C = LCase(Mid(Name, i, 1))

  If InStr("abcdefghijklmnopqrstuvwxyz_-.", C) <= 0 And Not IsNumeric(C) Then
  IsValidEmail = False
  Exit Function
  End If
  Next

  If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
  IsValidEmail = False
  Exit Function
  End If

  Next

  If InStr(names(1), ".") <= 0 Then
  IsValidEmail = False
  Exit Function
  End If

  i = Len(names(1)) - InStrRev(names(1), ".")

  If i <> 2 And i <> 3 Then
  IsValidEmail = False
  Exit Function
  End If

  If InStr(strEmail, "..") > 0 Then
  IsValidEmail = False
  Exit Function
  End If

  End Function

'###########ORG CONFIGURATION########################################

Sub SaveOrgShop()
    Dim filename As String
    Dim i As Long

    ' Declaração
    filename = App.Path & "\data\OrgShop.ini"

    For i = 1 To MAX_ORG_SHOP
        Call PutVar(filename, "Item", "Item" & i, STR$(OrgShop(i).Item))
        Call PutVar(filename, "Quantia", "Quantia" & i, STR$(OrgShop(i).Quantia))
        Call PutVar(filename, "Valor", "Valor" & i, STR$(OrgShop(i).Valor))
        Call PutVar(filename, "Level", "Level" & i, STR$(OrgShop(i).Level))
    Next
End Sub

Sub LoadOrgShop()
    Dim filename As String
    Dim i As Long
    
    ' Declaração
    filename = App.Path & "\data\OrgShop.ini"
    
    ' Verifica se o arquivo existe, se não cria-lo
    If Not FileExist(filename, True) Then Call SaveOrgShop

    For i = 1 To MAX_ORG_SHOP
        OrgShop(i).Item = Val(GetVar(filename, "Item", "Item" & i))
        OrgShop(i).Quantia = Val(GetVar(filename, "Quantia", "Quantia" & i))
        OrgShop(i).Valor = Val(GetVar(filename, "Valor", "Valor" & i))
        OrgShop(i).Level = Val(GetVar(filename, "Level", "Level" & i))
    Next
End Sub

Sub SaveOrgExp(ByVal OrgNum As Byte)
    Dim filename As String
    Dim i As Long, x As Long
    i = OrgNum
    
    ' Declaração
    filename = App.Path & "\data\Orgs.ini"

    Call PutVar(filename, "ORG" & i, "Exp", STR$(Organization(i).EXP))
    Call PutVar(filename, "ORG" & i, "Level", STR$(Organization(i).Level))
    
End Sub

Public Function FindOpenOrgMemberSlot(ByVal OrgNum As Long) As Long
Dim i As Integer
    
    For i = 1 To MAX_ORG_MEMBERS
        If Organization(OrgNum).OrgMember(i).Used = False Then
            FindOpenOrgMemberSlot = i
            Exit Function
        End If
    Next i
    
    'Guild is full sorry bub
    FindOpenOrgMemberSlot = 0

End Function

Public Sub ClearOrgMemberSlot(ByVal OrgNum As Long, ByVal MemberSlot As Long)

        'Evitar OverFlow
        If MemberSlot = 0 Then Exit Sub

        Organization(OrgNum).OrgMember(MemberSlot).Used = False
        Organization(OrgNum).OrgMember(MemberSlot).User_Login = vbNullString
        Organization(OrgNum).OrgMember(MemberSlot).User_Name = vbNullString
        Organization(OrgNum).OrgMember(MemberSlot).Online = False
            
        'Save guild after we remove member
        Call SaveOrg(OrgNum)
End Sub

Public Sub LoadOrg(ByVal OrgNum As Long)
Dim i As Integer
Dim filename As String
Dim F As Long
    
    'Evitar OverFlow
    If OrgNum > 4 Or OrgNum = 0 Then Exit Sub

    'Does this file even exist?
    If Not FileExist("\data\Orgs\Org" & OrgNum & ".dat") Then
        SaveOrg OrgNum
    End If
    
        filename = App.Path & "\data\Orgs\Org" & OrgNum & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Organization(OrgNum).EXP
            Get #F, , Organization(OrgNum).Level
            Get #F, , Organization(OrgNum).Lider
            For i = 1 To MAX_ORG_MEMBERS
                Get #F, , Organization(OrgNum).OrgMember(i)
            Next
        Close #F
        
        'Make sure an online flag didn't manage to slip through
        For i = 1 To MAX_ORG_MEMBERS
            If Organization(OrgNum).OrgMember(i).Online = True Then
                Organization(OrgNum).OrgMember(i).Online = False
            End If
        Next i
        
End Sub

Public Sub SaveOrg(ByVal OrgNum As Long)
Dim filename As String
Dim F As Long, i As Long

    'Evitar OverFlow
    If OrgNum > 3 Or OrgNum = 0 Then Exit Sub
    
    filename = App.Path & "\data\Orgs\Org" & OrgNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Organization(OrgNum).EXP
        Put #F, , Organization(OrgNum).Level
        Put #F, , Organization(OrgNum).Lider
        For i = 1 To MAX_ORG_MEMBERS
            Put #F, , Organization(OrgNum).OrgMember(i)
        Next
    Close #F
    
End Sub

Function GetPlayerOrg(ByVal Index As Long) As Long
    GetPlayerOrg = Player(Index).ORG
End Function

Sub SetPlayerOrg(ByVal Index As Long, ByVal ORG As Long)
    Player(Index).ORG = ORG
End Sub

Public Function CheckForSwears(Index As Long, Msg As String)
Dim SplitStr() As String
Dim SwearWords As String
Dim i As Integer

   ' SwearWords = Trim$(Options.Bloquear)
    SplitStr = Split(SwearWords, ",")
        
    For i = 0 To UBound(SplitStr)
        If InStr(1, LCase(Msg$), SplitStr(i), 1) Then
              Msg = Replace$((Msg), (SplitStr(i)), LCase(String(Len(SplitStr(i)), "*")), , , 1)
        End If
    Next i
End Function
