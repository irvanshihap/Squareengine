Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public TempEventMap(1 To MAX_MAPS) As GlobalEventsRec
Public MapCache(1 To MAX_MAPS) As Cache
Public temptile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public AEditor As PlayerRec

Public Rank(1 To MAX_RANK) As RankRec
Public Options As OptionsRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapBlocks(1 To MAX_MAPS) As MapBlockRec


Private Type MoveRouteRec
    index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    data5 As Long
    data6 As Long
End Type

Private Type GlobalEventRec
    x As Long
    y As Long
    Dir As Long
    active As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    Position As Long
    
    GraphicType As Long
    GraphicNum As Long
    GraphicX As Long
    GraphicX2 As Long
    GraphicY As Long
    GraphicY2 As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
End Type

Public Type GlobalEventsRec
    EventCount As Long
    Events() As GlobalEventRec
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
    Buy_Cost As Long
    Buy_Lvl As Long
    Buy_Item As Long
    Join_Cost As Long
    Join_Lvl As Long
    Join_Item As Long
    DayNight As Byte
    PathfindingType As Byte
End Type

Private Type RankRec

Name As String * ACCOUNT_LENGTH

Level As Long

End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
    
    PTCheckLevel As Long
    PTlevel As Long
    PTEXP As Long
    ' Props to Ryoku Hassu
End Type

Public Type PlayerInvRec
    num As Long
    Value As Long
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
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Spriteold As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Guilds
    GuildFileId As Long
    GuildMemberId As Long
    
    ' Admins
    Visible As Long
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Long
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    EquipmentOld(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    
    StartFlash As Long
    
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    SpellUses(1 To MAX_PLAYER_SPELLS) As Long
    
    'Color
    A As Byte
    R As Byte
    b As Byte
    G As Byte
    
ExpMultiplier As Long
ExpMultiplierTime As Long
    Transformed As Boolean
    
    IsMember As Byte
    DateCount As String
    
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Private Type EventCommandRec
    index As Byte
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    data5 As Long
    data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Private Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Private Type EventPageRec
    'These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    'End Conditions
    
    'Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    'Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    'Guidelines for the event
    WalkAnim As Long
    DirFix As Long
    WalkThrough As Long
    ShowName As Long
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    'For EventMap
    x As Long
    y As Long
End Type

Private Type EventRec
    Name As String
    Global As Byte
    PageCount As Long
    Pages() As EventPageRec
    x As Long
    y As Long
    'Self Switches re-set on restart.
    SelfSwitches(0 To 4) As Long
End Type

Public Type GlobalMapEvents
    eventID As Long
    pageID As Long
    x As Long
    y As Long
End Type

Private Type MapEventRec
    Dir As Long
    x As Long
    y As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    
    movementspeed As Long
    Position As Long
    Visible As Long
    eventID As Long
    pageID As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
    SelfSwitches(0 To 4) As Long
End Type

Private Type EventMapRec
    CurrentEvents As Long
    EventPages() As MapEventRec
End Type

Private Type EventProcessingRec
    CurList As Long
    CurSlot As Long
    eventID As Long
    pageID As Long
    WaitingForResponse As Long
    ActionTimer As Long
    ListLeftOff() As Long
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
    Target As Long
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
    ' guild
    tmpGuildSlot As Long
    tmpGuildInviteSlot As Long
    tmpGuildInviteTimer As Long
    tmpGuildInviteId As Long

    EventMap As EventMapRec
    EventProcessingCount As Long
    EventProcessing() As EventProcessingRec
    
    BlindTimer As Long
    BlindDuration As Long
    
    StealthTimer As Long
    StealthDuration As Long
    
    isHealing As Boolean
    
    Buffs(1 To 10) As Long
    BuffTimer(1 To 10) As Long
    BuffValue(1 To 10) As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec
    
    IsMember As Byte
    DayNight As Byte
    Panorama As Byte
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    stat(1 To Stats.Stat_Count - 1) As Long
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
    
    Next As Byte
    Advancement(1 To 3) As Byte
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
    ClassReq(1 To MAX_CLASSREQ) As Byte
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Long
    Rarity As Byte
    speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Long
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    Stackable As Byte
    
    gift(1 To 10) As Long
    gchance(1 To 10) As Byte
    gvalue(1 To 10) As Long
    multigift As Boolean
    
    istwohander As Boolean
    
    Projectile As Long
    Range As Byte
    Rotation As Integer
    Ammo As Long
    
    Element As Long
    
    A As Byte
    R As Byte
    G As Byte
    b As Byte
    
    CarryWeight As Long
    Toolpower As Long
    
    'crafting
    Tool As Long
    ToolReq As Long
    Recipe(1 To MAX_RECIPE_ITEMS) As Long
    MageProjectile As Boolean
    
    addExpMultiplier As Long
addExpMultiplierTime As Long
    Sprite As Long
    
     IsMember As Byte
    
End Type

Private Type MapItemRec
    num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Long
    DropItemValue(1 To MAX_NPC_DROPS) As Long
    stat(1 To Stats.Stat_Count - 1) As Long
    HP As Long
    exp As Long
    Animation As Long
    Damage As Long
    Level As Long
    speed As Long
    Quest As Byte
    QuestNum As Long
    
    Projectile As Long
    ProjectileRange As Byte
    Rotation As Integer
    
    AttackSpeed As Integer
    
    Spell(1 To MAX_NPC_SPELLS) As Long
    
    Element As Long
    
    A As Byte
    R As Byte
    G As Byte
    b As Byte
    
    isBoss As Byte
    SpawnAtDay As Byte
    SpawnAtNight As Byte
End Type

Private Type MapNpcRec
    num As Long
    Target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
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
    StartFlash As Long
    
    ' Npc spells
    SpellTimer As Long
    Heals As Integer
    FirstCast As Boolean
    
    BlindDuration As Long
    BlindTimer As Long
    
    StealthDuration As Long
    StealthTimer As Long
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
    y As Long
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
    BlindDuration As Long
    StealthDuration As Long
    NextRank As Long
    NextUses As Long
    
    'Stat Bases
    StrBase As Boolean
    AgiBase As Boolean
    WillBase As Boolean
    IntBase As Boolean
    EndBase As Boolean
    
    BuffType As Long
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    NPC() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
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
    WalkThrough As Boolean
    Animation As Long
    
    ToolpowerReq As Long
    exp As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Public Type Vector
    x As Long
    y As Long
End Type

Public Type MapBlockRec
    Blocks() As Long
End Type
