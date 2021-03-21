Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Quest(1 To MAX_QUESTS) As QuestRec
Public Leilao(1 To MAX_LEILAO) As LeilaoRec
Public Pendencia(1 To MAX_LEILAO) As PendRec
Public RankLevel(1 To MAX_RANKS) As RankRec
Public Organization(1 To 4) As OrganizationRec
Public OrgShop(1 To MAX_ORG_SHOP) As OrgShopRec

Public Options As OptionsRec

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
End Type

'Org Shop
Private Type OrgShopRec
    Item As Long
    Quantia As Long
    Valor As Long
    Level As Long
End Type

Private Type PokeRec
    Pokemon As Integer
    Pokeball As Byte
    Level As Byte
    EXP As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    Stat(1 To Stats.Stat_Count - 1) As Long
    Spells(1 To MAX_POKE_SPELL) As Byte
    Negatives(1 To 11) As Long
    Felicidade As Long
    Sexo As Byte
    Shiny As Byte
    Berry(1 To MAX_BERRYS) As Long
End Type

Private Type PendRec
    Vendedor As String
    ItemNum As Long
    Poke As PokeRec
    Price As Long
    Tipo As Long
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
    PT As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    PokeInfo As PokeRec
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
    Pokemon As Long
    Pokeball As Long
End Type

Private Type OrgMemberRec
    Used As Boolean
    User_Login As String * ACCOUNT_LENGTH
    User_Name As String * ACCOUNT_LENGTH
    Online As Boolean
End Type

Private Type OrganizationRec
    EXP As Long
    Level As Long
    Lider As String * ACCOUNT_LENGTH
    OrgMember(1 To MAX_ORG_MEMBERS) As OrgMemberRec
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    Email As String * 255
    SecondPass As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    UltConexao(1 To 3) As Integer '1 - Dia, 2 - Mês & 3 - Ano
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    EquipPokeInfo(1 To Equipment.Equipment_Count - 1) As PokeRec
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    Y As Byte
    Dir As Byte
    MySprite As Long
    
    'Informações de Treinador
    Pokedex(1 To MAX_POKEMONS) As Byte
    PokeQntia As Byte 'Max 6
    Vitorias As Long
    Derrotas As Long
    Insignia(1 To MAX_INSIGNIAS) As Byte
    Quests(1 To MAX_QUESTS) As PlayerQuestRec
    
    'Trainer Point
    TPX As Long
    TPY As Long
    TPDir As Long
    TPSprite As Long

    'Evolução
    EvolPermition As Byte
    EvolTimerStone As Long
    EvolStone As Byte
    EvoId As Integer
    
    'Habilidade
    LearnSpell(1 To 3) As Integer '1 Aprendendo Spell Nova, 2 Numero da Spell Nova,3 Numero do Item
    LearnFila(1 To 10) As Integer 'Fila de habilidades
    
    'Vip
    MyVip As Byte
    VipStart As String 'Dia/Mês/Ano
    VipDays(1 To 6) As Integer
    VipInName As Boolean
    VipPoints As Integer
    Pontos As Long
    
    'Organização
    ORG As Byte
    Honra As Long
    OrgMember As Byte
    
    'Outras Informações
    PokeInicial As Byte
    MutedTime As Long
    VitalTemp As Long
    InFishing As Long
    InSurf As Byte
    UltRodLevel As Long
    Flying As Byte
    PokeLight As Boolean
    MyMap(1 To 3) As Long
    NgtDamage(1 To MAX_NEGATIVES) As Long
    Visuais(1 To 50) As Long
    Teleport(1 To 30) As Long
    Cabelo As Byte
    isBanned As Byte
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    Pokemon As Long
    EvolTimer As Long
    'Quest
    QuestInvite As Integer
    QuestSelect As Integer
    Conversando As Long
    ConversandoC As Long
    ScanTime As Long
    ScanPokemon As Long
    ' Surf Adicional
    SurfSlideTo As Byte
    Lutando As Long
    LutandoA As Long
    LutandoT As Long
    LutQntPoke As Long
    SwitPoke As Long
    NgtTick(1 To MAX_NEGATIVES) As Long
    Running As Boolean
    GymQntPoke As Byte
    InBattleGym As Byte
    GymTimer As Long
    GymLeaderPoke(1 To 2) As Long
End Type

Private Type TileDataRec
    x As Long
    Y As Long
    Tileset As Long
End Type

Private Type TileRec
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
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
    Cabelos() As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
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
    Speed As Long
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
    Spell(1 To MAX_POKE_SPELL) As Long
    Tipo As Long
    Vel As Long
    NDrop As Boolean
    NTrade As Boolean
    NDeposit As Boolean
    Berry As Long
    YesNo(1 To 5) As Boolean
    BauItem(1 To MAX_BAU) As Long
    BauValue(1 To MAX_BAU) As Long
    GiveAll As Boolean
    VNum As Byte
    VSlot As Byte
End Type

Private Type MapItemRec
    Num As Long
    Value As Long
    x As Byte
    Y As Byte
    ' despawn
    canDespawn As Boolean
    despawnTimer As Long
    PokeInfo As PokeRec
End Type
 
Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    Chance As Long
    Pokemon As Long
    Quest(1 To MAX_NPC_QUESTS) As Integer
    Money As Long
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    Y As Byte
    Dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    Desmaiado As Boolean
    Pescado As Boolean
    Sexo As Boolean
    Shiny As Boolean
    Level As Integer
    Stat(1 To Stats.Stat_Count - 1) As Long
    InBattle As Boolean
    Spell(1 To 4) As Integer
    GymBattle As Byte
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
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
    Teste As Boolean
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    Y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
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

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Public Type RankRec
    Name As String
    PokeNum As Long
    Level As Long
End Type
