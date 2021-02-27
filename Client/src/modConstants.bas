Attribute VB_Name = "modConstants"
Option Explicit

' Procura as janelas abertas
Declare Function FindWindow Lib "user32" Alias _
                            "FindWindowA" (ByVal lpClassName As String, _
                                           ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias _
                             "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                                             ByVal wParam As Long, lParam As Any) As Long

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' animated buttons
Public Const MAX_MENUBUTTONS As Long = 30
Public Const MAX_MAINBUTTONS As Long = 30
Public Const MAX_QUESTBUTTONS As Long = 3
Public Const MAX_CLOSEBUTTONS As Long = 30

Public Const MENUBUTTON_PATH As String = "\Data Files\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\Data Files\graphics\gui\main\buttons\"
Public Const QUESTBUTTON_PATH As String = "\Data Files\graphics\gui\main\questbuttons\"
Public Const CLOSEBUTTON_PATH As String = "\Data Files\graphics\gui\main\closebuttons\"

' Hotbar
Public Const HotbarTop As Long = 6
Public Const HotbarLeft As Long = 6
Public Const HotbarOffsetX As Long = 8

' Inventory constants
Public Const InvTop As Long = 38
Public Const InvLeft As Long = 11
Public Const InvOffsetY As Long = 8
Public Const InvOffsetX As Long = 8
Public Const InvColumns As Long = 5

' Leilao constants
Public Const LeilaoTop As Long = 10
Public Const LeilaoLeft As Long = 6
Public Const LeilaoOffsetY As Long = 9
Public Const LeilaoOffsetX As Long = 9
Public Const LeilaoColumns As Long = 5

' Bank constants
Public Const BankTop As Long = 35
Public Const BankLeft As Long = 8
Public Const BankOffsetY As Long = 9
Public Const BankOffsetX As Long = 9
Public Const BankColumns As Long = 13

' spells constants
Public Const SpellTop As Long = 24
Public Const SpellLeft As Long = 12
Public Const SpellOffsetY As Long = 49
Public Const SpellOffsetX As Long = 50
Public Const SpellColumns As Long = 5

' shop constants
Public Const ShopTop As Long = 9
Public Const ShopLeft As Long = 9
Public Const ShopOffsetY As Long = 8
Public Const ShopOffsetX As Long = 8
Public Const ShopColumns As Long = 5

' Character consts
Public Const EqTop As Long = 224
Public Const EqLeft As Long = 18
Public Const EqOffsetX As Long = 10
Public Const EqColumns As Long = 4

' Trade consts
Public Const TradeTop As Long = 21
Public Const TradeLeft As Long = 7
Public Const TradeOffsetY As Long = 9
Public Const TradeOffsetX As Long = 9
Public Const TradeColumns As Long = 5

' OrgShop consts
Public Const OrgTop As Long = 23
Public Const OrgLeft As Long = 13
Public Const OrgOffsetX As Long = 9
Public Const OrgOffsetY As Long = 9
Public Const OrgColumns As Long = 1

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"

' Font variables
Public Const FONT_NAME As String = "Georgia"
Public Const FONT_SIZE As Byte = 14
Public Const FONT_NAME2 As String = "Georgia"
Public Const FONT_SIZE2 As Byte = 14
Public Const FONT_NAME3 As String = "Arial Black"
Public Const FONT_SIZE3 As Byte = 15

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\Data Files\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\Data Files\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\Data Files\graphics\"
Public Const GFX_EXT As String = ".bmp"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Speed moving vars
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 6

' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Long = 32
Public Const SIZE_Y As Long = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const MAX_PLAYERS As Long = 70
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 999
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 20
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 15
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 100
Public Const MAX_BANK As Long = 130
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_QUESTS As Byte = 100
Public Const MAX_PLAYER_QUESTS As Byte = 100
Public Const MAX_QUEST_TASKS As Byte = 5
Public Const MAX_NPC_QUESTS As Byte = 5
Public Const MAX_LEILAO As Byte = 255
Public Const MAX_INSIGNIAS As Byte = 100
Public Const MAX_RANKS As Byte = 10
Public Const MAX_BERRYS As Byte = 5
Public Const MAX_NEGATIVES As Long = 11
Public Const MAX_NOTICIAS As Byte = 15
Public Const MAX_ORG_SHOP As Byte = 20
Public Const MAX_BAU As Byte = 15
Public Const MAX_ORG_MEMBERS As Byte = 36
Public Const QUEST_MAX_REWARDS As Byte = 10
Public Const UNLOCKED_POKEMONS As Integer = 252

' Website
Public Const GAME_WEBSITE As String = ""

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = (1152 / 32 - 1)
Public Const MAX_MAPY As Byte = (640 / 32 - 1)
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_ARENA As Byte = 2
Public Const MAP_MORAL_INTERIOR As Byte = 3
Public Const MAP_MORAL_ROCKTUNEL As Byte = 4

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_GRASS As Byte = 15
Public Const TILE_TYPE_WATER As Byte = 16
Public Const TILE_TYPE_FISHING As Byte = 17
Public Const TILE_TYPE_SCRIPT As Byte = 18
Public Const TILE_TYPE_FLYAVOID As Byte = 19
Public Const TILE_TYPE_SIGN As Byte = 20

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_CONSUME As Byte = 5
Public Const ITEM_TYPE_KEY As Byte = 6
Public Const ITEM_TYPE_CURRENCY As Byte = 7
Public Const ITEM_TYPE_SPELL As Byte = 8
Public Const ITEM_TYPE_STONE As Byte = 9
Public Const ITEM_TYPE_ROD As Byte = 10
Public Const ITEM_TYPE_UP As Byte = 11
Public Const ITEM_TYPE_POKEDEX As Byte = 12
Public Const ITEM_TYPE_BAU As Byte = 13

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement: Tiles per Second
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4
Public Const NPC_BEHAVIOUR_QUEST As Byte = 5

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_LINEAR As Byte = 5
Public Const SPELL_TYPE_SCRIPT As Byte = 6
Public Const SPELL_TYPE_FLY As Byte = 7

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
'7 EDITOR_POKEMON
Public Const EDITOR_QUEST As Byte = 8

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3
Public Const DIALOGUE_TYPE_QUEST As Byte = 4
Public Const DIALOGUE_TYPE_PM As Byte = 5
Public Const DIALOGUE_TYPE_LT As Byte = 6

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2
Public Const ACTIONMSG_STATICLOCKED As Long = 3

' stuffs
Public Const HalfX As Integer = ((MAX_MAPX + 1) / 2) * PIC_X
Public Const HalfY As Integer = ((MAX_MAPY + 1) / 2) * PIC_Y
Public Const ScreenX As Integer = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Integer = (MAX_MAPY + 1) * PIC_Y
Public Const StartXValue As Integer = ((MAX_MAPX + 1) / 2)
Public Const StartYValue As Integer = ((MAX_MAPY + 1) / 2)
Public Const EndXValue As Integer = (MAX_MAPX + 1) + 1
Public Const EndYValue As Integer = (MAX_MAPY + 1) + 1
Public Const Half_PIC_X As Integer = PIC_X / 2
Public Const Half_PIC_Y As Integer = PIC_Y / 2

' weather
Public Const WEATHER_RAINING As Long = 1
Public Const WEATHER_SNOWING As Long = 2
Public Const WEATHER_BIRD As Long = 3
Public Const WEATHER_SAND As Long = 4

' Chat Bubble Mondo
Public Const ChatBubbleWidth As Long = 200
Public Const Font_Default As String = "Default"
