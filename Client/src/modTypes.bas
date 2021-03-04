Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Quest(1 To MAX_QUESTS) As QuestRec
Public Leilao(1 To MAX_LEILAO) As LeilaoRec
Public RankLevel(1 To MAX_RANKS) As RankRec
Public Organization(1 To 4) As OrganizationRec
Public OrgShop(1 To MAX_ORG_SHOP) As OrgShopRec

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public QuestButton(1 To MAX_QUESTBUTTONS) As ButtonRec
Public CloseButton(1 To MAX_CLOSEBUTTONS) As ButtonRec
Public Party As PartyRec
Public MapSounds() As MapSoundRec
Public MapSoundCount As Long

' options
Public Options As OptionsRec

'Sounds
Public Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    channel As Long
End Type

'Org Shop
Private Type OrgShopRec
    Item As Long
    Quantia As Long
    Valor As Long
    Level As Long
End Type

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    sound As Byte
    wasd As Byte
    Debug As Byte
    MiniMap As Byte
    Quest As Byte
End Type

Private Type PokeRec
    Pokemon As Integer
    Pokeball As Byte
    Level As Byte
    Exp As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    Stat(1 To Stats.Stat_Count - 1) As Long
    Spells(1 To 4) As Byte
    Negatives(1 To 11) As Long
    Felicidade As Long
    Sexo As Byte
    Shiny As Byte
    Berry(1 To MAX_BERRYS) As Long
End Type

Private Type LeilaoRec
    Vendedor As String
    ItemNum As Long
    Poke As PokeRec
    Price As Long
    Tempo As Long
    Tipo As Long
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    value As Long
    PokeInfo As PokeRec
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    SpellNum As Long
    timer As Long
    FramePointer As Long
End Type

Private Type OrgMemberRec
    Used As Boolean
    User_Name As String * ACCOUNT_LENGTH
    Online As Boolean
End Type

Private Type OrganizationRec
    Exp As Long
    Level As Long
    OrgMember(1 To MAX_ORG_MEMBERS) As OrgMemberRec
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Long
    Sprite As Long
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Sex As Byte

    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long

    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long

    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    EquipPokeInfo(1 To Equipment.Equipment_Count - 1) As PokeRec

    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    MySprite As Long

    ' TrainerPoint
    TPX As Long
    TPY As Long
    TPDir As Long
    TPSprite As Long

    'Customização
    HairModel As Integer
    HairColor As Byte
    HairNum As Integer
    ClothModel As Integer
    ClothColor As Byte
    ClothNum As Integer
    LegsModel As Integer
    LegsColor As Byte
    LegsNum As Integer

    'Pokédex
    Pokedex(1 To MAX_POKEMONS) As Byte
    EvolPermition As Byte
    InFishing As Long
    InSurf As Byte
    Quests(1 To MAX_QUESTS) As PlayerQuestRec
    KillNpcs(1 To MAX_NPCS) As Integer
    KillPlayers As Integer
    MyLeiloes(1 To MAX_LEILAO) As Long
    ScanTime As Long
    Parado As Long
    Flying As Byte
    Insignia(1 To MAX_INSIGNIAS) As Byte
    AnimFrame As Byte
    Arena(1 To 10) As Long
    LearnSpell(1 To 3) As Integer    '1 Aprendendo Spell Nova, 2 Numero da Spell Nova,3 Numero do Item
    PokeLight As Boolean
    Vitorias As Long
    Derrotas As Long
    ORG As Byte
    Honra As Long
    MyVip As Byte    'Vip
    VipInName As Boolean    'Vip
    VipPoints As Integer    'Vip
    EvoId As Integer
    PuloSlide As Byte
    PuloStatus As Byte
    Running As Boolean
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH

    Revision As Long
    Moral As Byte

    Up As Long
    Down As Long
    Left As Long
    Right As Long

    BootMap As Long
    BootX As Byte
    BootY As Byte

    MaxX As Byte
    MaxY As Byte

    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long

    Weather As Long
    Intensity As Long
    LevelPoke(1 To 2) As Integer
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH

    Pic As Long
Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long

    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    Pokemon As Long
    Spell(1 To 4) As Long
    Tipo As Long
    vel As Long
    NDrop As Boolean
    NTrade As Boolean
    NDeposit As Boolean
    Berry As Long
    YesNo(1 To 5) As Boolean
    BauItem(1 To MAX_BAU) As Long
    BauValue(1 To MAX_BAU) As Long
    GiveAll As Boolean
End Type

Private Type MapItemRec
    playerName As String
    num As Long
    value As Long
    Frame As Byte
    X As Byte
    Y As Byte
    PokeInfo As PokeRec
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH

    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    Exp As Long
    Animation As Long
    Damage As Long
    Level As Long
    Chance As Long
    Pokemon As Long
    Quest(1 To MAX_NPC_QUESTS) As Integer
End Type

Private Type MapNpcRec
    num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    Desmaiado As Boolean
    Sexo As Boolean
    Shiny As Boolean
    AnimFrame As Byte
    Level As Integer
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH

Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Elemento As Long
    BaseStat As Byte
    Tamanho As Long
    AnimL As Long
    Element As Long
    Script As Long
    Voar1 As Long
    Voar2 As Long
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte    ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH

    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
    Spell As Long
End Type

Private Type ActionMsgRec
    message As String
    Created As Long
Type As Long
    color As Long
    Scroll As Long
    X As Long
    Y As Long
    timer As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    sound As String * NAME_LENGTH

    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
    Pokemon As Long
    Pokeball As Long
End Type

Public Type ButtonRec
    filename As String
    state As Byte
End Type

Type DropRec
    X As Long
    Y As Long
    ySpeed As Long
    xSpeed As Long
    Init As Boolean
End Type

Public Type RankRec
    Name As String
    PokeNum As Long
    Level As Long
End Type

' Chat Bubble Mondo
Public Type ChatBubbleRec
    Msg As String
    colour As Long
    target As Long
    targetType As Byte
    timer As Long
    active As Boolean
End Type
