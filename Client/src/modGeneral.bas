Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Pegar a resolução
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public DX7 As New DirectX7  ' Master Object, early binding

Public Sub Main()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set loading screen
    loadGUI True
    frmLoad.Visible = True

    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions

    ' load main menu
    Call SetStatus("Loading Menu...")
    Load frmMenu

    ' load gui
    Call SetStatus("Loading interface...")
    loadGUI

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\", "main"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "questbuttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    ' Clear game values
    Call SetStatus("Clearing game data...")
    Call ClearGameData

    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34)    ' "

    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name

    ' initialize DirectX
    If Not InitDirectDraw Then
        MsgBox "Error Initializing DirectX7 - DirectDraw."
        DestroyGame
    End If

    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Initializing DirectX...")

    ' DX7 Master Object is already created, early binding
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFacesShiny
    Call CheckPokeIcons
    Call CheckPokeIconShiny
    Call CheckHairNum

    ' load music/sound engine
    InitFmod

    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)

    ' Reset values
    Ping = -1

    'Load frmMainMenu
    Load frmMenu

    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    cacheButtons
    resetButtons_Menu

    ' we can now see it
    frmMenu.Visible = True

    ' set values for directional blocking arrows
    DirArrowX(1) = 12    ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12    ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0    ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23    ' right
    DirArrowY(4) = 12

    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.weapon

    ' hide the load form
    frmLoad.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub loadGUI(Optional ByVal loadingScreen As Boolean = False)
    Dim i As Long

    ' if we can't find the interface
    On Error GoTo errorhandler

    ' loading screen
    If loadingScreen Then
        frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
        Exit Sub
    End If

    ' menu
    'frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\background.jpg")
    'frmMenu.picMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\main.jpg")
    'frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\login.jpg")
    'frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\register.jpg")
    'frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\credits.jpg")
    'frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\character.jpg")
    ' main
    'frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main.jpg")
    'frmMain.picInventory.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\inventory.jpg")
    'frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\character.jpg")
    'frmMain.picSpells.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\skills.jpg")
    'frmMain.picOptions.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\options.jpg")
    'frmMain.picParty.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party.jpg")
    'frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_item.jpg")
    'frmMain.picSpellDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_spell.jpg")
    'frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    'frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    'frmMain.picTempSpell.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    'frmMain.picShop.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\shop.jpg")
    'frmMain.picBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bank.jpg")
    'frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\hotbar.jpg")

    ' main - bars
    frmMain.imgHPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.bmp")
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\spirit.bmp")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.bmp")

    ' main - party bars
    For i = 1 To MAX_PARTY_MEMBERS
        frmMain.imgPartyHealth(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_health.jpg")
        frmMain.imgPartySpirit(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_spirit.jpg")
    Next

    ' store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.Width
    SPRBar_Width = frmMain.imgMPBar.Width
    EXPBar_Width = frmMain.imgEXPBar.Width
    ' party
    Party_HPWidth = frmMain.imgPartyHealth(1).Width
    Party_SPRWidth = frmMain.imgPartySpirit(1).Width
    ORG = frmMain.PicExp.Width

    Exit Sub

    ' let them know we can't load the GUI
errorhandler:
    MsgBox "Cannot find one or more interface images." & vbNewLine & "If they exist then you have not extracted the project properly." & vbNewLine & "Please follow the installation instructions fully.", vbCritical
    DestroyGame
    Exit Sub
End Sub

Public Sub MenuState(ByVal state As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmLoad.Visible = True

    Select Case state
    Case MENU_STATE_ADDCHAR
        frmMenu.Visible = False
        frmMenu.picLogin.Visible = False
        frmMenu.picCharacter.Visible = False
        frmMenu.picRegister.Visible = False
        frmMenu.PicRecover.Visible = False

        If ConnectToServer(1) Then
            Call SetStatus("Connected, sending character addition data...")

            If frmMenu.optMale.value Then
                Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
            Else
                Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
            End If
        End If

    Case MENU_STATE_NEWACCOUNT
        frmMenu.Visible = False
        frmMenu.picLogin.Visible = False
        frmMenu.picCharacter.Visible = False
        frmMenu.picRegister.Visible = False
        frmMenu.PicRecover.Visible = False

        If ConnectToServer(1) Then
            Call SetStatus("Connected, sending new account information...")
            Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text, frmMenu.txtRecovery, frmMenu.txtEmail)
        End If

    Case MENU_STATE_LOGIN
        frmMenu.Visible = False
        frmMenu.picLogin.Visible = False
        frmMenu.picCharacter.Visible = False
        frmMenu.picRegister.Visible = False
        frmMenu.PicRecover.Visible = False

        If ConnectToServer(1) Then
            Call SetStatus("Connected, sending login information...")
            Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
            Exit Sub
        End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picLogin.Visible = True
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False
            frmLoad.Visible = False
            Call MsgBox("Desculpe o jogo parece estar offline, espere alguns minutos ou acesse o nosso site: " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
    Dim buffer As clsBuffer, i As Long

    isLogging = True
    InGame = False
    Set buffer = New clsBuffer
    buffer.WriteLong CQuit
    SendData buffer.ToArray()
    Set buffer = Nothing
    Call DestroyTCP

    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next

    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SpellX = 0
    SpellY = 0
    LastItemDesc = 0
    LastItemPoke = 0    'Limpar ItemPoke
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0

    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    Unload frmEditor_Pokemon
    Unload frmLeilao
    Unload frmEditor_Quest

    ' hide main form stuffs
    frmMain.picBank.Visible = False
    frmMenu.picMain.Visible = False
    frmMenu.picLogin.Visible = True
    frmMain.txtChat.text = vbNullString
    frmMain.txtMyChat.text = vbNullString
    frmMain.picCurrency.Visible = False
    frmMain.picDialogue.Visible = False
    frmMain.picInventory.Visible = False
    frmMain.picTrade.Visible = False
    frmMain.picTrade.Visible = False
    frmMain.lblTradeStatus(0).Caption = "Esperando Confirmação"
    frmMain.lblTradeStatus(1).Caption = "Esperando Confirmação"
    frmMain.lblTradeStatus(0).ForeColor = &HE0E0E0
    frmMain.lblTradeStatus(1).ForeColor = &HE0E0E0
    frmMain.PicTradeOn(0).Visible = False
    frmMain.PicTradeOn(1).Visible = False
    frmMain.picYourTrade.Visible = False
    frmMain.picTheirTrade.Visible = False
    frmMain.picCover.Visible = False
    frmMain.picSpells.Visible = False
    frmMain.picOptions.Visible = False
    frmMain.picParty.Visible = False
    frmMain.picPokedex.Visible = False
    frmMain.picQuest.Visible = False
    frmMain.imgClose(9).Visible = True
    frmMain.imgButton(14).Visible = True
    frmMain.PicPokeInicial.Visible = False
    frmMain.picRank.Visible = False
End Sub

Sub GameInit()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False

    ' bring all the main gui components to the front
    frmMain.picShop.ZOrder (0)
    frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)

    ' hide gui
    frmMain.picCover.Visible = False
    InBank = False
    InShop = False
    InTrade = False

    ' show the main form
    frmLoad.Visible = False
    frmMain.Show

    frmMain.Font = FONT_NAME
    frmMain.FontSize = FONT_SIZE - 5

    ' Set the focus
    Call SetFocusOnChat
    frmMain.picScreen.Visible = True

    'PicLoading
    LoadingPing = 15000 + GetTickCount

    ' Blt inv
    BltInventory

    ' blt hotbar
    blthotbar

    ' get ping
    GetPing
    DrawPing

    'stop the song playing
    StopMusic

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' break out of GameLoop
    InGame = False
    Call DestroyTCP

    'destroy objects in reverse order
    Call DestroyDirectDraw

    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If
    DestroyFmod
    Call UnloadAllForms
    End

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
    Dim Frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each Frm In VB.Forms
        Unload Frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmLoad.lblStatus.Caption = Caption
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.text = Txt.text + Msg + vbCrLf
    Else
        Txt.text = Txt.text + Msg
    End If

    Txt.SelStart = Len(Txt.text) - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetFocusOnChat()
    If chaton = False Then
        SetFocusOnGame
        Exit Sub
    End If

    On Error Resume Next    'prevent RTE5, no way to handle error
    frmMain.txtMyChat.Visible = True
    frmMain.PicChat.Visible = True

    frmMain.txtMyChat.SetFocus
End Sub

Public Sub SetFocusOnGame()

    If chaton = True Then
        SetFocusOnChat
        Exit Sub
    End If

    On Error Resume Next    'prevent RTE5, no way to handle error
    frmMain.txtMyChat.Visible = False
    frmMain.PicChat.Visible = False
    frmMain.picScreen.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Rand = Int((High - Low + 1) * Rnd) + Low

    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim GlobalX As Long
    Dim GlobalY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GlobalX = PB.Left
    GlobalY = PB.top

    If Button = 1 Then
        PB.Left = GlobalX + X - SOffsetX
        PB.top = GlobalY + Y - SOffsetY
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Botão Fechar
    With MenuButton(1)
        .filename = "fechar"
        .state = 0
    End With

    ' menu - login
    With MenuButton(2)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    ' menu - register
    With MenuButton(3)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    ' menu - credits
    With MenuButton(4)
        .filename = "fechar"
        .state = 0    ' normal
    End With

    With MenuButton(5)
        .filename = "confirmar"
        .state = 0
    End With

    With MenuButton(6)
        .filename = "fechar"
        .state = 0    ' normal
    End With

    With MenuButton(7)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MenuButton(8)
        .filename = "fechar"
        .state = 0    ' normal
    End With

    With MenuButton(9)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MenuButton(10)
        .filename = "fechar"
        .state = 0    ' normal
    End With

    With MenuButton(11)
        .filename = "fechar"
        .state = 0    ' normal
    End With

    With MenuButton(12)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    ' main - inv
    With MainButton(1)
        .filename = "inv"
        .state = 0    ' normal
    End With

    ' main - skills
    With MainButton(2)
        .filename = "leilao"
        .state = 0    ' normal
    End With

    ' main - char
    With MainButton(3)
        .filename = "quest"
        .state = 0    ' normal
    End With

    ' main - opt
    With MainButton(4)
        .filename = "opt"
        .state = 0    ' normal
    End With

    ' main - trade
    With MainButton(5)
        .filename = "temdex"
        .state = 0    ' normal
    End With

    ' main - pokdex
    With MainButton(6)
        .filename = "temdex"
        .state = 0    ' normal
    End With

    ' main - pokdex
    With MainButton(7)
        .filename = "private"
        .state = 0    ' normal
    End With

    ' main - pokdex
    With MainButton(8)
        .filename = "vip"
        .state = 0    ' normal
    End With

    ' main - pokdex
    With MainButton(9)
        .filename = "org"
        .state = 0    ' normal
    End With

    ' main - pokdex
    With MainButton(10)
        .filename = "personagem"
        .state = 0    ' normal
    End With

    With MainButton(11)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MainButton(12)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MainButton(13)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MainButton(14)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MainButton(15)
        .filename = "doar"
        .state = 0    ' normal
    End With

    With MainButton(16)
        .filename = "escolher"
        .state = 0    ' normal
    End With

    With MainButton(17)
        .filename = "escolher"
        .state = 0    ' normal
    End With

    With MainButton(18)
        .filename = "escolher"
        .state = 0    ' normal
    End With

    With MainButton(19)
        .filename = "escolher"
        .state = 0    ' normal
    End With

    With MainButton(20)
        .filename = "escolher"
        .state = 0    ' normal
    End With

    With MainButton(21)
        .filename = "escolher"
        .state = 0    ' normal
    End With

    With MainButton(22)
        .filename = "sair"
        .state = 0    ' normal
    End With

    With MainButton(23)
        .filename = "cancelar"
        .state = 0    ' normal
    End With

    With MainButton(24)
        .filename = "confirmar"
        .state = 0    ' normal
    End With

    With MainButton(25)
        .filename = "loja"
        .state = 0    ' normal
    End With

    With MainButton(26)
        .filename = "comprar"
        .state = 0    ' normal
    End With
    
    With MainButton(27)
        .filename = "confirmar"
        .state = 0    ' normal
    End With
    
    With MainButton(28)
        .filename = "comprar"
        .state = 0    ' normal
    End With
    
    With MainButton(29)
        .filename = "comprar"
        .state = 0    ' normal
    End With
    
    With MainButton(30)
        .filename = "vender"
        .state = 0    ' normal
    End With
        
        ' Chat enviar
    With MainButton(31)
        .filename = "enviar"
        .state = 0    ' normal
    End With
    
        ' Chat Sistem
    With MainButton(32)
        .filename = "sistema"
        .state = 0    ' normal
    End With
    
        ' Chat Mapa
    With MainButton(33)
        .filename = "mapa"
        .state = 0    ' normal
    End With
    
        ' Chat Todos
    With MainButton(34)
        .filename = "todos"
        .state = 0    ' normal
    End With
    
        ' Chat Grupo
    With MainButton(35)
        .filename = "grupo"
        .state = 0    ' normal
    End With
    
        ' Chat Org
    With MainButton(36)
        .filename = "corg"
        .state = 0    ' normal
    End With
    
    '###########Quest Buttons##############
    ' quest - informations
    With QuestButton(1)
        .filename = "list"
        .state = 0    ' normal
    End With

    ' quest - rewards
    With QuestButton(2)
        .filename = "rewards"
        .state = 0    ' normal
    End With

    ' quest - informations
    With QuestButton(3)
        .filename = "info"
        .state = 0    ' normal
    End With

    '############# Close Buttons #############'

    'Quest List
    With CloseButton(1)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(2)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(3)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(4)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(5)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(6)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(7)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(8)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(9)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(10)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(11)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(12)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(13)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(14)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(15)
        .filename = "fechar"
        .state = 0
    End With

    With CloseButton(16)
        .filename = "fechar"
        .state = 0
    End With
    
    With CloseButton(17)
        .filename = "fechar"
        .state = 0
    End With
    
    With CloseButton(18)
        .filename = "fechar"
        .state = 0
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).state = 0    'normal
            renderButton_Menu i
        End If
    Next

    If exceptionNum = 0 Then LastButtonSound_Menu = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal buttonNum As Long)
    Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' get the suffix
    Select Case MenuButton(buttonNum).state
    Case 0    ' normal
        bSuffix = "_norm"
    Case 1    ' hover
        bSuffix = "_hover"
    Case 2    ' click
        bSuffix = "_click"
    End Select

    ' render the button
    frmMenu.imgButton(buttonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonNum).filename & bSuffix & ".bmp")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal buttonNum As Long, ByVal bState As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MenuButton(buttonNum).state = bState
        renderButton_Menu buttonNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' main specific buttons
Public Sub resetButtons_Main(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' loop through entire array
    For i = 1 To MAX_MAINBUTTONS
        ' only change if different and not exception
        If Not MainButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MainButton(i).state = 0    'normal
            renderButton_Main i
        End If
    Next

    If exceptionNum = 0 Then LastButtonSound_Main = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Main(ByVal buttonNum As Long)
    Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' get the suffix
    Select Case MainButton(buttonNum).state
    Case 0    ' normal
        bSuffix = "_norm"
    Case 1    ' hover
        bSuffix = "_hover"
    Case 2    ' click
        bSuffix = "_click"
    End Select

    ' render the button
    frmMain.imgButton(buttonNum).Picture = LoadPicture(App.Path & MAINBUTTON_PATH & MainButton(buttonNum).filename & bSuffix & ".bmp")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Main(ByVal buttonNum As Long, ByVal bState As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MainButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MainButton(buttonNum).state = bState
        renderButton_Main buttonNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Botões fechar
Public Sub resetButtons_Close(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' loop through entire array
    For i = 1 To MAX_CLOSEBUTTONS
        ' only change if different and not exception
        If Not CloseButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            CloseButton(i).state = 0    'normal
            renderButton_Close i
        End If
    Next

    If exceptionNum = 0 Then LastButtonSound_Close = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Close", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Close(ByVal buttonNum As Long)
    Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' get the suffix
    Select Case CloseButton(buttonNum).state
    Case 0    ' normal
        bSuffix = "_norm"
    Case 1    ' hover
        bSuffix = "_hover"
    Case 2    ' click
        bSuffix = "_click"
    End Select

    ' render the button
    frmMain.imgClose(buttonNum).Picture = LoadPicture(App.Path & CLOSEBUTTON_PATH & CloseButton(buttonNum).filename & bSuffix & ".bmp")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Close", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Close(ByVal buttonNum As Long, ByVal bState As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If CloseButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        CloseButton(buttonNum).state = bState
        renderButton_Close buttonNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Close", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
    Dim strLoad As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Cache music list
    strLoad = Dir(App.Path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop

    ' Cache sound list
    strLoad = Dir(App.Path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearGameData()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ClearNpcs
    Call ClearResources
    Call ClearItems
    Call ClearShops
    Call ClearSpells
    Call ClearAnimations
    Call ClearPokemons
    Call ClearQuests

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearGameData", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsPerPixels(pixels As Integer) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TwipsPerPixels = pixels * Screen.TwipsPerPixelX

    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsPerPixels", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ChangeToWindowed()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ChangingResolution = True

    ' Restaurar
    DD.SetDisplayMode OldResolutionX, OldResolutionY, 32, 0, DDSDM_DEFAULT

    ' Ajustar a janela
    frmMain.WindowState = vbNormal
    frmMain.BorderStyle = 1
    frmMain.Caption = frmMain.Caption
    frmMain.Width = 12090
    frmMain.Height = 9600
    frmMain.Left = OldMainLeft
    frmMain.top = OldMainTop

    ResetDDS_Primary

    ' Terminamos, volte a desenhar
    ChangingResolution = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeToWindowed", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeToFullScreen()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Ei, pare um pouco de desenhar
    ChangingResolution = True

    ' Armazenar informações da resolução
    OldResolutionX = GetDeviceCaps(frmMain.hDC, 8)
    OldResolutionY = GetDeviceCaps(frmMain.hDC, 10)
    OldMainLeft = frmMain.Left
    OldMainTop = frmMain.top

    ' Apagar borda
    frmMain.BorderStyle = 0
    frmMain.Caption = frmMain.Caption

    ' Valores do fullscreen
    DD.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT

    frmMain.WindowState = vbNormal
    frmMain.Left = 0
    frmMain.top = 0
    frmMain.WindowState = vbMaximized

    ' Resetar desenhista
    ResetDDS_Primary

    ChangingResolution = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeToFullScreen", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' quest specific buttons
Public Sub resetButtons_Quest(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' loop through entire array
    For i = 1 To MAX_QUESTBUTTONS
        ' only change if different and not exception
        If Not QuestButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            QuestButton(i).state = 0    'normal
            renderButton_Quest i
        End If
    Next

    If exceptionNum = 0 Then LastButtonSound_Quest = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Quest", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Quest(ByVal buttonNum As Long)
    Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' get the suffix
    Select Case QuestButton(buttonNum).state
    Case 0    ' normal
        bSuffix = "_norm"
    Case 1    ' hover
        bSuffix = "_hover"
    Case 2    ' click
        bSuffix = "_click"
    End Select

    ' render the button
    'frmMain.imgQuest(buttonNum).Picture = LoadPicture(App.Path & QUESTBUTTON_PATH & QuestButton(buttonNum).filename & bSuffix & ".jpg")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Quest", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Quest(ByVal buttonNum As Long, ByVal bState As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If QuestButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        QuestButton(buttonNum).state = bState
        renderButton_Quest buttonNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Quest", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
