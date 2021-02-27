Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "leilão"
    ChkDir App.Path & "\Data\", "pendencias"
    LoadRankLevel

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Port = 7001
        Options.MOTD = "Welcome to Eclipse Origins."
        Options.Website = "http://www.touchofdeathforums.com/smf/"
        SaveOptions
    Else
        LoadOptions
    End If
    
    For i = 1 To MAX_LEILAO
        LoadLeilão i
    Next
    
    For i = 1 To MAX_LEILAO
        LoadPendencia i
    Next
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call SetStatus("Unloading ORGS...")
    For i = 1 To MAX_ORGS
        Call SaveOrg(i)
    Next
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing Temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing Maps...")
    Call ClearMaps
    Call SetStatus("Clearing Map items...")
    Call ClearMapItems
    Call SetStatus("Clearing Map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing Npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing Items...")
    Call ClearItems
    Call SetStatus("Clearing Shops...")
    Call ClearShops
    Call SetStatus("Clearing Spells...")
    Call ClearSpells
    Call SetStatus("Clearing Animations...")
    Call ClearAnimations
    Call SetStatus("Clearing Pokémons...")
    Call ClearPokemons
    Call SetStatus("Clearing Quests...")
    Call ClearQuests
End Sub

Private Sub LoadGameData()
Dim i As Long

    Call SetStatus("Loading Classes...")
    Call LoadClasses
    Call SetStatus("Loading Maps...")
    Call LoadMaps
    Call SetStatus("Loading Items...")
    Call LoadItems
    Call SetStatus("Loading Npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading Shops...")
    Call LoadShops
    Call SetStatus("Loading Spells...")
    Call LoadSpells
    Call SetStatus("Loading Animations...")
    Call LoadAnimations
    Call SetStatus("Loading Pokémons...")
    Call LoadPokemons
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Loading Orgs...")
    For i = 1 To MAX_ORGS
    Call LoadOrg(i)
    Next
    Call SetStatus("Loading OrgShop...")
    Call LoadOrgShop
    Call SetStatus("Loading OrgMembers...")
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function
