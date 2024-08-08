Attribute VB_Name = "Protocol"
Option Explicit

#If False Then
    Dim Map, X, Y, n, Mapa, race, helmet, weapon, shield, color, Value, errHandler, punishments, Length, obj, index As Variant
#End If

Private Const SEPARATOR As String * 1 = vbNullChar
Private auxiliarBuffer  As clsByteQueue

Private Enum ServerPacketID
    Logged = 1
    RemoveDialogs = 2
    RemoveCharDialog = 3
    NavigateToggle = 4
    Disconnect = 5
    CommerceEnd = 6
    BankEnd = 7
    CommerceInit = 8
    BankInit = 9
    UserCommerceInit = 10
    UserCommerceEnd = 11
    UserOfferConfirm = 12
    CommerceChat = 13
    UpdateSta = 14
    UpdateMana = 15
    UpdateHP = 16
    UpdateGold = 17
    UpdateBankGold = 18
    UpdateExp = 19
    ChangeMap = 20
    PosUpdate = 21
    ChatOverHead = 22
    ConsoleMsg = 23
    GuildChat = 24
    ShowMessageBox = 25
    UserIndexInServer = 26
    UserCharIndexInServer = 27
    CharacterCreate = 28
    CharacterRemove = 29
    CharacterChangeNick = 30
    CharacterMove = 31
    ForceCharMove = 32
    CharacterChange = 33
    HeadingChange = 34
    ObjectCreate = 35
    ObjectDelete = 36
    BlockPosition = 37
    PlayMp3 = 38
    PlayMidi = 39
    PlayWave = 40
    guildList = 41
    AreaChanged = 42
    PauseToggle = 43
    RainToggle = 44
    CreateFX = 45
    UpdateUserStats = 46
    ChangeInventorySlot = 47
    ChangeBankSlot = 48
    ChangeSpellSlot = 49
    Atributes = 50
    BlacksmithWeapons = 51
    BlacksmithArmors = 52
    InitCarpenting = 53
    RestOK = 54
    errorMsg = 55
    Blind = 56
    Dumb = 57
    ShowSignal = 58
    ChangeNPCInventorySlot = 59
    UpdateHungerAndThirst = 60
    Fame = 61
    MiniStats = 62
    LevelUp = 63
    AddForumMsg = 64
    ShowForumForm = 65
    SetInvisible = 66
    DiceRoll = 67
    MeditateToggle = 68
    BlindNoMore = 69
    DumbNoMore = 70
    SendSkills = 71
    TrainerCreatureList = 72
    guildNews = 73
    OfferDetails = 74
    AlianceProposalsList = 75
    PeaceProposalsList = 76
    CharacterInfo = 77
    GuildLeaderInfo = 78
    GuildMemberInfo = 79
    GuildDetails = 80
    ShowGuildFundationForm = 81
    ParalizeOK = 82
    ShowUserRequest = 83
    ChangeUserTradeSlot = 84
    SendNight = 85
    Pong = 86
    UpdateTagAndStatus = 87
    SpawnList = 88
    ShowSOSForm = 89
    ShowMOTDEditionForm = 90
    ShowGMPanelForm = 91
    UserNameList = 92
    ShowDenounces = 93
    RecordList = 94
    RecordDetails = 95
    ShowGuildAlign = 96
    ShowPartyForm = 97
    UpdateStrenghtAndDexterity = 98
    UpdateStrenght = 99
    UpdateDexterity = 100
    AddSlots = 101
    MultiMessage = 102
    StopWorking = 103
    CancelOfferItem = 104
    PalabrasMagicas = 105
    PlayAttackAnim = 106
    FXtoMap = 107
    AccountLogged = 108
    SearchList = 109
    QuestDetails = 110
    QuestListSend = 111
    CreateDamage = 112
    UserInEvent = 113
    RenderMsg = 114
    DeletedChar = 115
    EquitandoToggle = 116
    EnviarDatosServer = 117
    InitCraftman = 118
    EnviarListDeAmigos = 119
    SeeInProcess = 120
    ShowProcess = 121
    proyectil = 122
    PlayIsInChatMode = 123
End Enum

Private Enum ClientPacketID
    LoginExistingChar = 1
    ThrowDices = 2
    LoginNewChar = 3
    Talk = 4
    Yell = 5
    Whisper = 6
    Walk = 7
    RequestPositionUpdate = 8
    Attack = 9
    PickUp = 10
    SafeToggle = 11
    ResuscitationSafeToggle = 12
    RequestGuildLeaderInfo = 13
    RequestAtributes = 14
    RequestFame = 15
    RequestSkills = 16
    RequestMiniStats = 17
    CommerceEnd = 18
    UserCommerceEnd = 19
    UserCommerceConfirm = 20
    CommerceChat = 21
    BankEnd = 22
    UserCommerceOk = 23
    UserCommerceReject = 24
    Drop = 25
    CastSpell = 26
    LeftClick = 27
    DoubleClick = 28
    Work = 29
    UseSpellMacro = 30
    UseItem = 31
    CraftBlacksmith = 32
    CraftCarpenter = 33
    WorkLeftClick = 34
    CreateNewGuild = 35
    sadasdA = 36
    EquipItem = 37
    ChangeHeading = 38
    ModifySkills = 39
    Train = 40
    CommerceBuy = 41
    BankExtractItem = 42
    CommerceSell = 43
    BankDeposit = 44
    ForumPost = 45
    MoveSpell = 46
    MoveBank = 47
    ClanCodexUpdate = 48
    UserCommerceOffer = 49
    GuildAcceptPeace = 50
    GuildRejectAlliance = 51
    GuildRejectPeace = 52
    GuildAcceptAlliance = 53
    GuildOfferPeace = 54
    GuildOfferAlliance = 55
    GuildAllianceDetails = 56
    GuildPeaceDetails = 57
    GuildRequestJoinerInfo = 58
    GuildAlliancePropList = 59
    GuildPeacePropList = 60
    GuildDeclareWar = 61
    GuildNewWebsite = 62
    GuildAcceptNewMember = 63
    GuildRejectNewMember = 64
    GuildKickMember = 65
    GuildUpdateNews = 66
    GuildMemberInfo = 67
    GuildOpenElections = 68
    GuildRequestMembership = 69
    GuildRequestDetails = 70
    Online = 71
    Quit = 72
    GuildLeave = 73
    RequestAccountState = 74
    PetStand = 75
    PetFollow = 76
    ReleasePet = 77
    TrainList = 78
    Rest = 79
    Meditate = 80
    Resucitate = 81
    Heal = 82
    Help = 83
    RequestStats = 84
    CommerceStart = 85
    BankStart = 86
    Enlist = 87
    Information = 88
    Reward = 89
    RequestMOTD = 90
    UpTime = 91
    PartyLeave = 92
    PartyCreate = 93
    PartyJoin = 94
    Inquiry = 95
    GuildMessage = 96
    PartyMessage = 97
    GuildOnline = 98
    PartyOnline = 99
    CouncilMessage = 100
    RoleMasterRequest = 101
    GMRequest = 102
    bugReport = 103
    ChangeDescription = 104
    GuildVote = 105
    punishments = 106
    ChangePassword = 107
    Gamble = 108
    InquiryVote = 109
    LeaveFaction = 110
    BankExtractGold = 111
    BankDepositGold = 112
    Denounce = 113
    GuildFundate = 114
    GuildFundation = 115
    PartyKick = 116
    PartySetLeader = 117
    PartyAcceptMember = 118
    Ping = 119
    RequestPartyForm = 120
    ItemUpgrade = 121
    GMCommands = 122
    InitCrafting = 123
    Home = 124
    ShowGuildNews = 125
    ShareNpc = 126
    StopSharingNpc = 127
    Consultation = 128
    moveItem = 129
    LoginExistingAccount = 130
    LoginNewAccount = 131
    CentinelReport = 132
    Ecvc = 133
    Acvc = 134
    IrCvc = 135
    DragAndDropHechizos = 136
    Quest = 137
    QuestAccept = 138
    QuestListRequest = 139
    QuestDetailsRequest = 140
    QuestAbandon = 141
    CambiarContrasena = 142
    FightSend = 143
    FightAccept = 144
    CloseGuild = 145
    Discord = 146
    DeleteChar = 147
    ObtenerDatosServer = 148
    CraftsmanCreate = 149
    AddAmigos = 150
    DelAmigos = 151
    OnAmigos = 152
    MsgAmigos = 153
    LookProcess = 154
    SendProcessList = 155
    SendIfCharIsInChatMode = 156
End Enum

Private Const LAST_CLIENT_PACKET_ID As Byte = 156

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_CRIMINAL
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss
End Enum

Public Sub InitAuxiliarBuffer()
    Set auxiliarBuffer = New clsByteQueue
End Sub

Public Function HandleIncomingData(ByVal Userindex As Integer) As Boolean
    On Error Resume Next
    With UserList(Userindex)
        .Counters.PacketsTick = .Counters.PacketsTick + 1
        Dim packetID As Long: packetID = CLng(.incomingData.PeekByte())
        If Not (packetID = ClientPacketID.ThrowDices _
                Or packetID = ClientPacketID.LoginExistingChar _
                Or packetID = ClientPacketID.LoginNewChar _
                Or packetID = ClientPacketID.LoginNewAccount _
                Or packetID = ClientPacketID.LoginExistingAccount _
                Or packetID = ClientPacketID.DeleteChar _
                Or packetID = ClientPacketID.ObtenerDatosServer _
                Or packetID = ClientPacketID.CambiarContrasena) Then
            If Not .flags.UserLogged Then
                Call CloseSocket(Userindex)
                Exit Function
            ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
                .Counters.IdleCount = 0
            End If
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            .Counters.IdleCount = 0
            If .flags.UserLogged Then
                Call CloseSocket(Userindex)
                Exit Function
            End If
        End If
        .flags.NoPuedeSerAtacado = False
    End With
    Select Case packetID
        Case ClientPacketID.SendIfCharIsInChatMode
            Call HandleSendIfCharIsInChatMode(Userindex)
            
        Case ClientPacketID.LoginExistingChar
            Call HandleLoginExistingChar(Userindex)
        
        Case ClientPacketID.ThrowDices
            Call HandleThrowDices(Userindex)
        
        Case ClientPacketID.LoginNewChar
            Call HandleLoginNewChar(Userindex)

        Case ClientPacketID.DeleteChar
            Call HandleDeleteChar(Userindex)
        
        Case ClientPacketID.Talk
            Call HandleTalk(Userindex)
        
        Case ClientPacketID.Yell
            Call HandleYell(Userindex)
        
        Case ClientPacketID.Whisper
            Call HandleWhisper(Userindex)
        
        Case ClientPacketID.Walk
            Call HandleWalk(Userindex)
        
        Case ClientPacketID.RequestPositionUpdate
            Call HandleRequestPositionUpdate(Userindex)
        
        Case ClientPacketID.Attack
            Call HandleAttack(Userindex)
        
        Case ClientPacketID.PickUp
            Call HandlePickUp(Userindex)
        
        Case ClientPacketID.SafeToggle
            Call HandleSafeToggle(Userindex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(Userindex)
        
        Case ClientPacketID.RequestGuildLeaderInfo
            Call HandleRequestGuildLeaderInfo(Userindex)
        
        Case ClientPacketID.RequestAtributes
            Call HandleRequestAtributes(Userindex)
        
        Case ClientPacketID.RequestFame
            Call HandleRequestFame(Userindex)
        
        Case ClientPacketID.RequestSkills
            Call HandleRequestSkills(Userindex)
        
        Case ClientPacketID.RequestMiniStats
            Call HandleRequestMiniStats(Userindex)
        
        Case ClientPacketID.CommerceEnd
            Call HandleCommerceEnd(Userindex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(Userindex)
        
        Case ClientPacketID.UserCommerceEnd
            Call HandleUserCommerceEnd(Userindex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(Userindex)
        
        Case ClientPacketID.BankEnd
            Call HandleBankEnd(Userindex)
        
        Case ClientPacketID.UserCommerceOk
            Call HandleUserCommerceOk(Userindex)
        
        Case ClientPacketID.UserCommerceReject
            Call HandleUserCommerceReject(Userindex)
        
        Case ClientPacketID.Drop
            Call HandleDrop(Userindex)
        
        Case ClientPacketID.CastSpell
            Call HandleCastSpell(Userindex)
        
        Case ClientPacketID.LeftClick
            Call HandleLeftClick(Userindex)
        
        Case ClientPacketID.DoubleClick
            Call HandleDoubleClick(Userindex)
        
        Case ClientPacketID.Work
            Call HandleWork(Userindex)
        
        Case ClientPacketID.UseSpellMacro
            Call HandleUseSpellMacro(Userindex)
        
        Case ClientPacketID.UseItem
            Call HandleUseItem(Userindex)
        
        Case ClientPacketID.CraftBlacksmith
            Call HandleCraftBlacksmith(Userindex)
        
        Case ClientPacketID.CraftCarpenter
            Call HandleCraftCarpenter(Userindex)
        
        Case ClientPacketID.WorkLeftClick
            Call HandleWorkLeftClick(Userindex)
        
        Case ClientPacketID.CreateNewGuild
            Call HandleCreateNewGuild(Userindex)
        
        Case ClientPacketID.EquipItem
            Call HandleEquipItem(Userindex)
        
        Case ClientPacketID.ChangeHeading
            Call HandleChangeHeading(Userindex)
        
        Case ClientPacketID.ModifySkills
            Call HandleModifySkills(Userindex)
        
        Case ClientPacketID.Train
            Call HandleTrain(Userindex)
        
        Case ClientPacketID.CommerceBuy
            Call HandleCommerceBuy(Userindex)
        
        Case ClientPacketID.BankExtractItem
            Call HandleBankExtractItem(Userindex)
        
        Case ClientPacketID.CommerceSell
            Call HandleCommerceSell(Userindex)
        
        Case ClientPacketID.BankDeposit
            Call HandleBankDeposit(Userindex)
        
        Case ClientPacketID.ForumPost
            Call HandleForumPost(Userindex)
        
        Case ClientPacketID.MoveSpell
            Call HandleMoveSpell(Userindex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(Userindex)
        
        Case ClientPacketID.ClanCodexUpdate
            Call HandleClanCodexUpdate(Userindex)
        
        Case ClientPacketID.UserCommerceOffer
            Call HandleUserCommerceOffer(Userindex)
        
        Case ClientPacketID.GuildAcceptPeace
            Call HandleGuildAcceptPeace(Userindex)
        
        Case ClientPacketID.GuildRejectAlliance
            Call HandleGuildRejectAlliance(Userindex)
        
        Case ClientPacketID.GuildRejectPeace
            Call HandleGuildRejectPeace(Userindex)
        
        Case ClientPacketID.GuildAcceptAlliance
            Call HandleGuildAcceptAlliance(Userindex)
        
        Case ClientPacketID.GuildOfferPeace
            Call HandleGuildOfferPeace(Userindex)
        
        Case ClientPacketID.GuildOfferAlliance
            Call HandleGuildOfferAlliance(Userindex)
        
        Case ClientPacketID.GuildAllianceDetails
            Call HandleGuildAllianceDetails(Userindex)
        
        Case ClientPacketID.GuildPeaceDetails
            Call HandleGuildPeaceDetails(Userindex)
        
        Case ClientPacketID.GuildRequestJoinerInfo
            Call HandleGuildRequestJoinerInfo(Userindex)
        
        Case ClientPacketID.GuildAlliancePropList
            Call HandleGuildAlliancePropList(Userindex)
        
        Case ClientPacketID.GuildPeacePropList
            Call HandleGuildPeacePropList(Userindex)
        
        Case ClientPacketID.GuildDeclareWar
            Call HandleGuildDeclareWar(Userindex)
        
        Case ClientPacketID.GuildNewWebsite
            Call HandleGuildNewWebsite(Userindex)
        
        Case ClientPacketID.GuildAcceptNewMember
            Call HandleGuildAcceptNewMember(Userindex)
        
        Case ClientPacketID.GuildRejectNewMember
            Call HandleGuildRejectNewMember(Userindex)
        
        Case ClientPacketID.GuildKickMember
            Call HandleGuildKickMember(Userindex)
        
        Case ClientPacketID.GuildUpdateNews
            Call HandleGuildUpdateNews(Userindex)
        
        Case ClientPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo(Userindex)
        
        Case ClientPacketID.GuildOpenElections
            Call HandleGuildOpenElections(Userindex)
        
        Case ClientPacketID.GuildRequestMembership
            Call HandleGuildRequestMembership(Userindex)
        
        Case ClientPacketID.GuildRequestDetails
            Call HandleGuildRequestDetails(Userindex)
                  
        Case ClientPacketID.Online
            Call HandleOnline(Userindex)
        
        Case ClientPacketID.Quit
            Call HandleQuit(Userindex)
        
        Case ClientPacketID.GuildLeave
            Call HandleGuildLeave(Userindex)
        
        Case ClientPacketID.RequestAccountState
            Call HandleRequestAccountState(Userindex)
        
        Case ClientPacketID.PetStand
            Call HandlePetStand(Userindex)
        
        Case ClientPacketID.PetFollow
            Call HandlePetFollow(Userindex)
            
        Case ClientPacketID.ReleasePet
            Call HandleReleasePet(Userindex)
        
        Case ClientPacketID.TrainList
            Call HandleTrainList(Userindex)
        
        Case ClientPacketID.Rest
            Call HandleRest(Userindex)
        
        Case ClientPacketID.Meditate
            Call HandleMeditate(Userindex)
        
        Case ClientPacketID.Resucitate
            Call HandleResucitate(Userindex)
        
        Case ClientPacketID.Heal
            Call HandleHeal(Userindex)
        
        Case ClientPacketID.Help
            Call HandleHelp(Userindex)
        
        Case ClientPacketID.RequestStats
            Call HandleRequestStats(Userindex)
        
        Case ClientPacketID.CommerceStart
            Call HandleCommerceStart(Userindex)
        
        Case ClientPacketID.BankStart
            Call HandleBankStart(Userindex)
        
        Case ClientPacketID.Enlist
            Call HandleEnlist(Userindex)
        
        Case ClientPacketID.Information
            Call HandleInformation(Userindex)
        
        Case ClientPacketID.Reward
            Call HandleReward(Userindex)
        
        Case ClientPacketID.RequestMOTD
            Call HandleRequestMOTD(Userindex)
        
        Case ClientPacketID.UpTime
            Call HandleUpTime(Userindex)
        
        Case ClientPacketID.PartyLeave
            Call HandlePartyLeave(Userindex)
        
        Case ClientPacketID.PartyCreate
            Call HandlePartyCreate(Userindex)
        
        Case ClientPacketID.PartyJoin
            Call HandlePartyJoin(Userindex)
        
        Case ClientPacketID.Inquiry
            Call HandleInquiry(Userindex)
        
        Case ClientPacketID.GuildMessage
            Call HandleGuildMessage(Userindex)
        
        Case ClientPacketID.PartyMessage
            Call HandlePartyMessage(Userindex)
        
        Case ClientPacketID.GuildOnline
            Call HandleGuildOnline(Userindex)
        
        Case ClientPacketID.PartyOnline
            Call HandlePartyOnline(Userindex)
        
        Case ClientPacketID.CouncilMessage
            Call HandleCouncilMessage(Userindex)
        
        Case ClientPacketID.RoleMasterRequest
            Call HandleRoleMasterRequest(Userindex)
        
        Case ClientPacketID.GMRequest
            Call HandleGMRequest(Userindex)
        
        Case ClientPacketID.bugReport
            Call HandleBugReport(Userindex)
        
        Case ClientPacketID.ChangeDescription
            Call HandleChangeDescription(Userindex)
        
        Case ClientPacketID.GuildVote
            Call HandleGuildVote(Userindex)
        
        Case ClientPacketID.punishments
            Call HandlePunishments(Userindex)
        
        Case ClientPacketID.ChangePassword
            Call HandleChangePassword(Userindex)
        
        Case ClientPacketID.Gamble
            Call HandleGamble(Userindex)
        
        Case ClientPacketID.InquiryVote
            Call HandleInquiryVote(Userindex)
        
        Case ClientPacketID.LeaveFaction
            Call HandleLeaveFaction(Userindex)
        
        Case ClientPacketID.BankExtractGold
            Call HandleBankExtractGold(Userindex)
        
        Case ClientPacketID.BankDepositGold
            Call HandleBankDepositGold(Userindex)
        
        Case ClientPacketID.Denounce
            Call HandleDenounce(Userindex)
        
        Case ClientPacketID.GuildFundate
            Call HandleGuildFundate(Userindex)
            
        Case ClientPacketID.GuildFundation
            Call HandleGuildFundation(Userindex)
        
        Case ClientPacketID.PartyKick
            Call HandlePartyKick(Userindex)
        
        Case ClientPacketID.PartySetLeader
            Call HandlePartySetLeader(Userindex)
        
        Case ClientPacketID.PartyAcceptMember
            Call HandlePartyAcceptMember(Userindex)
        
        Case ClientPacketID.Ping
            Call HandlePing(Userindex)
            
        Case ClientPacketID.RequestPartyForm
            Call HandlePartyForm(Userindex)
            
        Case ClientPacketID.ItemUpgrade
            Call HandleItemUpgrade(Userindex)
        
        Case ClientPacketID.GMCommands
            Call HandleGMCommands(Userindex)
            
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(Userindex)
        
        Case ClientPacketID.Home
            Call HandleHome(Userindex)
        
        Case ClientPacketID.ShowGuildNews
            Call HandleShowGuildNews(Userindex)
            
        Case ClientPacketID.ShareNpc
            Call HandleShareNpc(Userindex)
            
        Case ClientPacketID.StopSharingNpc
            Call HandleStopSharingNpc(Userindex)
            
        Case ClientPacketID.Consultation
            Call HandleConsultation(Userindex)
        
        Case ClientPacketID.moveItem
            Call HandleMoveItem(Userindex)

        Case ClientPacketID.LoginExistingAccount
            Call HandleLoginExistingAccount(Userindex)

        Case ClientPacketID.LoginNewAccount
            Call HandleLoginNewAccount(Userindex)
        
        Case ClientPacketID.CentinelReport
            Call HandleCentinelReport(Userindex)
            
        Case ClientPacketID.Ecvc
            Call HandleEnviaCvc(Userindex)

        Case ClientPacketID.Acvc
            Call HandleAceptarCvc(Userindex)

        Case ClientPacketID.IrCvc
            Call HandleIrCvc(Userindex)
            
        Case ClientPacketID.DragAndDropHechizos
            Call HandleDragAndDropHechizos(Userindex)
  
        Case ClientPacketID.Quest
            Call Quests.HandleQuest(Userindex)
            
        Case ClientPacketID.QuestAccept
            Call Quests.HandleQuestAccept(Userindex)
        
        Case ClientPacketID.QuestListRequest
            Call Quests.HandleQuestListRequest(Userindex)
        
        Case ClientPacketID.QuestDetailsRequest
            Call Quests.HandleQuestDetailsRequest(Userindex)
        
        Case ClientPacketID.QuestAbandon
            Call Quests.HandleQuestAbandon(Userindex)
        
        Case ClientPacketID.CambiarContrasena
            Call HandleCambiarContrasena(Userindex)

        Case ClientPacketID.FightSend
            Call HandleFightSend(Userindex)
            
        Case ClientPacketID.FightAccept
            Call HandleFightAccept(Userindex)
        
        Case ClientPacketID.CloseGuild
            Call HandleCloseGuild(Userindex)
        
        Case ClientPacketID.Discord
            Call HandleDiscord(Userindex)

        Case ClientPacketID.ObtenerDatosServer
            Call HandleObtenerDatosServer(Userindex)
            
        Case ClientPacketID.CraftsmanCreate
            Call HandleCraftsmanCreate(Userindex)
      
        Case ClientPacketID.AddAmigos
            Call Amigos.HandleAddAmigo(Userindex)

        Case ClientPacketID.DelAmigos
            Call Amigos.HandleDelAmigo(Userindex)

        Case ClientPacketID.OnAmigos
            Call Amigos.HandleOnAmigo(Userindex)

        Case ClientPacketID.MsgAmigos
            Call Amigos.HandleMsgAmigo(Userindex)

        Case ClientPacketID.LookProcess
            Call HandleLookProcess(Userindex)

        Case ClientPacketID.SendProcessList
            Call HandleSendProcessList(Userindex)
            
        Case Else
            Call CloseSocket(Userindex)
    End Select
    If UserList(Userindex).incomingData.Length > 0 And Err.Number = 0 Then
        Err.Clear
        HandleIncomingData = True
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(Userindex).incomingData.NotEnoughDataErrCode Then
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & vbTab & " LastDllError: " & Err.LastDllError & vbTab & " - UserIndex: " & Userindex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(Userindex)
        HandleIncomingData = False
    Else
        Call FlushBuffer(Userindex)
        HandleIncomingData = False
    End If
End Function

Public Sub WriteMultiMessage(ByVal Userindex As Integer, ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        Select Case MessageIndex

            Case eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteByte(Arg1)
                Call .WriteInteger(Arg2)
                
            Case eMessages.UserHitNPC
                Call .WriteLong(Arg1)
                
            Case eMessages.UserAttackedSwing
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInteger(Arg1)
                Call .WriteByte(Arg2)
                Call .WriteInteger(Arg3)
                
            Case eMessages.UserHittedUser
                Call .WriteInteger(Arg1)
                Call .WriteByte(Arg2)
                Call .WriteInteger(Arg3)
                
            Case eMessages.WorkRequestTarget
                Call .WriteByte(Arg1)
            
            Case eMessages.HaveKilledUser
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                Call .WriteLong(Arg2)
            
            Case eMessages.UserKill
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)

            Case eMessages.EarnExp
                Call .WriteLong(Arg1)
                
            Case eMessages.Home
                Call .WriteByte(CByte(Arg1))
                Call .WriteInteger(CInt(Arg2))
                Call .WriteASCIIString(StringArg1)
                
            Case eMessages.UserMuerto
            
            Case eMessages.NpcInmune
            
            Case eMessages.Hechizo_HechiceroMSG_NOMBRE
                Call .WriteByte(CByte(Arg1))
                Call .WriteASCIIString(StringArg1)
             
            Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
                Call .WriteByte(CByte(Arg1))
             
            Case eMessages.Hechizo_HechiceroMSG_CRIATURA
                Call .WriteByte(CByte(Arg1))
             
            Case eMessages.Hechizo_PropioMSG
                Call .WriteByte(CByte(Arg1))
         
            Case eMessages.Hechizo_TargetMSG
                Call .WriteByte(CByte(Arg1))
                Call .WriteASCIIString(StringArg1)
        End Select
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub HandleGMCommands(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim Command As Byte
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Command = .incomingData.PeekByte
        Select Case Command
            Case eGMCommands.GMMessage
                Call HandleGMMessage(Userindex)
        
            Case eGMCommands.showName
                Call HandleShowName(Userindex)
        
            Case eGMCommands.OnlineRoyalArmy
                Call HandleOnlineRoyalArmy(Userindex)
        
            Case eGMCommands.OnlineChaosLegion
                Call HandleOnlineChaosLegion(Userindex)
        
            Case eGMCommands.GoNearby
                Call HandleGoNearby(Userindex)
        
            Case eGMCommands.comment
                Call HandleComment(Userindex)
        
            Case eGMCommands.serverTime
                Call HandleServerTime(Userindex)
        
            Case eGMCommands.Where
                Call HandleWhere(Userindex)
        
            Case eGMCommands.CreaturesInMap
                Call HandleCreaturesInMap(Userindex)
        
            Case eGMCommands.WarpMeToTarget
                Call HandleWarpMeToTarget(Userindex)
        
            Case eGMCommands.WarpChar
                Call HandleWarpChar(Userindex)
        
            Case eGMCommands.Silence
                Call HandleSilence(Userindex)
        
            Case eGMCommands.SOSShowList
                Call HandleSOSShowList(Userindex)
            
            Case eGMCommands.SOSRemove
                Call HandleSOSRemove(Userindex)
        
            Case eGMCommands.GoToChar
                Call HandleGoToChar(Userindex)
        
            Case eGMCommands.invisible
                Call HandleInvisible(Userindex)
        
            Case eGMCommands.GMPanel
                Call HandleGMPanel(Userindex)
        
            Case eGMCommands.RequestUserList
                Call HandleRequestUserList(Userindex)
        
            Case eGMCommands.Working
                Call HandleWorking(Userindex)
        
            Case eGMCommands.Hiding
                Call HandleHiding(Userindex)
        
            Case eGMCommands.Jail
                Call HandleJail(Userindex)
        
            Case eGMCommands.KillNPC
                Call HandleKillNPC(Userindex)
        
            Case eGMCommands.WarnUser
                Call HandleWarnUser(Userindex)
        
            Case eGMCommands.EditChar
                Call HandleEditChar(Userindex)
        
            Case eGMCommands.RequestCharInfo
                Call HandleRequestCharInfo(Userindex)
        
            Case eGMCommands.RequestCharStats
                Call HandleRequestCharStats(Userindex)
        
            Case eGMCommands.RequestCharGold
                Call HandleRequestCharGold(Userindex)
        
            Case eGMCommands.RequestCharInventory
                Call HandleRequestCharInventory(Userindex)
        
            Case eGMCommands.RequestCharBank
                Call HandleRequestCharBank(Userindex)
        
            Case eGMCommands.RequestCharSkills
                Call HandleRequestCharSkills(Userindex)
        
            Case eGMCommands.ReviveChar
                Call HandleReviveChar(Userindex)
        
            Case eGMCommands.OnlineGM
                Call HandleOnlineGM(Userindex)
        
            Case eGMCommands.OnlineMap
                Call HandleOnlineMap(Userindex)
        
            Case eGMCommands.Forgive
                Call HandleForgive(Userindex)
        
            Case eGMCommands.Kick
                Call HandleKick(Userindex)
        
            Case eGMCommands.Execute
                Call HandleExecute(Userindex)
        
            Case eGMCommands.BanChar
                Call HandleBanChar(Userindex)
        
            Case eGMCommands.UnbanChar
                Call HandleUnbanChar(Userindex)
        
            Case eGMCommands.NPCFollow
                Call HandleNPCFollow(Userindex)
        
            Case eGMCommands.SummonChar
                Call HandleSummonChar(Userindex)
        
            Case eGMCommands.SpawnListRequest
                Call HandleSpawnListRequest(Userindex)
        
            Case eGMCommands.SpawnCreature
                Call HandleSpawnCreature(Userindex)
        
            Case eGMCommands.ResetNPCInventory
                Call HandleResetNPCInventory(Userindex)
        
            Case eGMCommands.ServerMessage
                Call HandleServerMessage(Userindex)
        
            Case eGMCommands.MapMessage
                Call HandleMapMessage(Userindex)
            
            Case eGMCommands.NickToIP
                Call HandleNickToIP(Userindex)
        
            Case eGMCommands.IPToNick
                Call HandleIPToNick(Userindex)
        
            Case eGMCommands.GuildOnlineMembers
                Call HandleGuildOnlineMembers(Userindex)
        
            Case eGMCommands.TeleportCreate
                Call HandleTeleportCreate(Userindex)
        
            Case eGMCommands.TeleportDestroy
                Call HandleTeleportDestroy(Userindex)
        
            Case eGMCommands.RainToggle
                Call HandleRainToggle(Userindex)
        
            Case eGMCommands.SetCharDescription
                Call HandleSetCharDescription(Userindex)

            Case eGMCommands.ForceMP3ToMap
                Call HanldeForceMP3ToMap(Userindex)
        
            Case eGMCommands.ForceMIDIToMap
                Call HanldeForceMIDIToMap(Userindex)
        
            Case eGMCommands.ForceWAVEToMap
                Call HandleForceWAVEToMap(Userindex)
        
            Case eGMCommands.RoyalArmyMessage
                Call HandleRoyalArmyMessage(Userindex)
        
            Case eGMCommands.ChaosLegionMessage
                Call HandleChaosLegionMessage(Userindex)
        
            Case eGMCommands.CitizenMessage
                Call HandleCitizenMessage(Userindex)
        
            Case eGMCommands.CriminalMessage
                Call HandleCriminalMessage(Userindex)
        
            Case eGMCommands.TalkAsNPC
                Call HandleTalkAsNPC(Userindex)
        
            Case eGMCommands.DestroyAllItemsInArea
                Call HandleDestroyAllItemsInArea(Userindex)
        
            Case eGMCommands.AcceptRoyalCouncilMember
                Call HandleAcceptRoyalCouncilMember(Userindex)
        
            Case eGMCommands.AcceptChaosCouncilMember
                Call HandleAcceptChaosCouncilMember(Userindex)
        
            Case eGMCommands.ItemsInTheFloor
                Call HandleItemsInTheFloor(Userindex)
        
            Case eGMCommands.MakeDumb
                Call HandleMakeDumb(Userindex)
        
            Case eGMCommands.MakeDumbNoMore
                Call HandleMakeDumbNoMore(Userindex)
        
            Case eGMCommands.DumpIPTables
                Call HandleDumpIPTables(Userindex)
        
            Case eGMCommands.CouncilKick
                Call HandleCouncilKick(Userindex)
        
            Case eGMCommands.SetTrigger
                Call HandleSetTrigger(Userindex)
        
            Case eGMCommands.AskTrigger
                Call HandleAskTrigger(Userindex)
        
            Case eGMCommands.BannedIPList
                Call HandleBannedIPList(Userindex)
        
            Case eGMCommands.BannedIPReload
                Call HandleBannedIPReload(Userindex)
        
            Case eGMCommands.GuildMemberList
                Call HandleGuildMemberList(Userindex)
        
            Case eGMCommands.GuildBan
                Call HandleGuildBan(Userindex)
        
            Case eGMCommands.BanIP
                Call HandleBanIP(Userindex)
        
            Case eGMCommands.UnbanIP
                Call HandleUnbanIP(Userindex)
        
            Case eGMCommands.CreateItem
                Call HandleCreateItem(Userindex)
        
            Case eGMCommands.DestroyItems
                Call HandleDestroyItems(Userindex)
        
            Case eGMCommands.ChaosLegionKick
                Call HandleChaosLegionKick(Userindex)
        
            Case eGMCommands.RoyalArmyKick
                Call HandleRoyalArmyKick(Userindex)

            Case eGMCommands.ForceMP3All
                Call HandleForceMP3All(Userindex)
        
            Case eGMCommands.ForceMIDIAll
                Call HandleForceMIDIAll(Userindex)
        
            Case eGMCommands.ForceWAVEAll
                Call HandleForceWAVEAll(Userindex)
        
            Case eGMCommands.RemovePunishment
                Call HandleRemovePunishment(Userindex)
        
            Case eGMCommands.TileBlockedToggle
                Call HandleTileBlockedToggle(Userindex)
        
            Case eGMCommands.KillNPCNoRespawn
                Call HandleKillNPCNoRespawn(Userindex)
        
            Case eGMCommands.KillAllNearbyNPCs
                Call HandleKillAllNearbyNPCs(Userindex)
        
            Case eGMCommands.LastIP
                Call HandleLastIP(Userindex)
        
            Case eGMCommands.ChangeMOTD
                Call HandleChangeMOTD(Userindex)
        
            Case eGMCommands.SetMOTD
                Call HandleSetMOTD(Userindex)
        
            Case eGMCommands.SystemMessage
                Call HandleSystemMessage(Userindex)
        
            Case eGMCommands.CreateNPC
                Call HandleCreateNPC(Userindex)
        
            Case eGMCommands.ImperialArmour
                Call HandleImperialArmour(Userindex)
        
            Case eGMCommands.ChaosArmour
                Call HandleChaosArmour(Userindex)
        
            Case eGMCommands.NavigateToggle
                Call HandleNavigateToggle(Userindex)
        
            Case eGMCommands.ServerOpenToUsersToggle
                Call HandleServerOpenToUsersToggle(Userindex)
        
            Case eGMCommands.TurnOffServer
                Call HandleTurnOffServer(Userindex)
        
            Case eGMCommands.TurnCriminal
                Call HandleTurnCriminal(Userindex)
        
            Case eGMCommands.ResetFactions
                Call HandleResetFactions(Userindex)
        
            Case eGMCommands.RemoveCharFromGuild
                Call HandleRemoveCharFromGuild(Userindex)
        
            Case eGMCommands.RequestCharMail
                Call HandleRequestCharMail(Userindex)
        
            Case eGMCommands.AlterPassword
                Call HandleAlterPassword(Userindex)
        
            Case eGMCommands.AlterMail
                Call HandleAlterMail(Userindex)
        
            Case eGMCommands.AlterName
                Call HandleAlterName(Userindex)
        
            Case Declaraciones.eGMCommands.DoBackUp
                Call HandleDoBackUp(Userindex)
        
            Case eGMCommands.ShowGuildMessages
                Call HandleShowGuildMessages(Userindex)
        
            Case eGMCommands.SaveMap
                Call HandleSaveMap(Userindex)
        
            Case eGMCommands.ChangeMapInfoPK
                Call HandleChangeMapInfoPK(Userindex)
            
            Case eGMCommands.ChangeMapInfoBackup
                Call HandleChangeMapInfoBackup(Userindex)
        
            Case eGMCommands.ChangeMapInfoRestricted
                Call HandleChangeMapInfoRestricted(Userindex)
        
            Case eGMCommands.ChangeMapInfoNoMagic
                Call HandleChangeMapInfoNoMagic(Userindex)
        
            Case eGMCommands.ChangeMapInfoNoInvi
                Call HandleChangeMapInfoNoInvi(Userindex)
        
            Case eGMCommands.ChangeMapInfoNoResu
                Call HandleChangeMapInfoNoResu(Userindex)
        
            Case eGMCommands.ChangeMapInfoLand
                Call HandleChangeMapInfoLand(Userindex)
        
            Case eGMCommands.ChangeMapInfoZone
                Call HandleChangeMapInfoZone(Userindex)
        
            Case eGMCommands.ChangeMapInfoStealNpc
                Call HandleChangeMapInfoStealNpc(Userindex)
            
            Case eGMCommands.ChangeMapInfoNoOcultar
                Call HandleChangeMapInfoNoOcultar(Userindex)
            
            Case eGMCommands.ChangeMapInfoNoInvocar
                Call HandleChangeMapInfoNoInvocar(Userindex)
            
            Case eGMCommands.SaveChars
                Call HandleSaveChars(Userindex)
        
            Case eGMCommands.CleanSOS
                Call HandleCleanSOS(Userindex)
        
            Case eGMCommands.ShowServerForm
                Call HandleShowServerForm(Userindex)
        
            Case eGMCommands.night
                Call HandleNight(Userindex)
        
            Case eGMCommands.KickAllChars
                Call HandleKickAllChars(Userindex)
        
            Case eGMCommands.ReloadNPCs
                Call HandleReloadNPCs(Userindex)
        
            Case eGMCommands.ReloadServerIni
                Call HandleReloadServerIni(Userindex)
        
            Case eGMCommands.ReloadSpells
                Call HandleReloadSpells(Userindex)
        
            Case eGMCommands.ReloadObjects
                Call HandleReloadObjects(Userindex)
        
            Case eGMCommands.Restart
                Call HandleRestart(Userindex)
        
            Case eGMCommands.ResetAutoUpdate
                Call HandleResetAutoUpdate(Userindex)
        
            Case eGMCommands.ChatColor
                Call HandleChatColor(Userindex)
        
            Case eGMCommands.Ignored
                Call HandleIgnored(Userindex)
        
            Case eGMCommands.CheckSlot
                Call HandleCheckSlot(Userindex)
        
            Case eGMCommands.SetIniVar
                Call HandleSetIniVar(Userindex)
            
            Case eGMCommands.CreatePretorianClan
                Call HandleCreatePretorianClan(Userindex)
         
            Case eGMCommands.RemovePretorianClan
                Call HandleDeletePretorianClan(Userindex)
                
            Case eGMCommands.EnableDenounces
                Call HandleEnableDenounces(Userindex)
            
            Case eGMCommands.ShowDenouncesList
                Call HandleShowDenouncesList(Userindex)
        
            Case eGMCommands.SetDialog
                Call HandleSetDialog(Userindex)
            
            Case eGMCommands.Impersonate
                Call HandleImpersonate(Userindex)
            
            Case eGMCommands.Imitate
                Call HandleImitate(Userindex)
            
            Case eGMCommands.RecordAdd
                Call HandleRecordAdd(Userindex)
            
            Case eGMCommands.RecordAddObs
                Call HandleRecordAddObs(Userindex)
            
            Case eGMCommands.RecordRemove
                Call HandleRecordRemove(Userindex)
            
            Case eGMCommands.RecordListRequest
                Call HandleRecordListRequest(Userindex)
            
            Case eGMCommands.RecordDetailsRequest
                Call HandleRecordDetailsRequest(Userindex)
            
            Case eGMCommands.ExitDestroy
                Call HandleExitDestroy(Userindex)

            Case eGMCommands.ToggleCentinelActivated
                Call HandleToggleCentinelActivated(Userindex)
        
            Case eGMCommands.SearchNpc
                Call HandleSearchNpc(Userindex)
           
            Case eGMCommands.SearchObj
                Call HandleSearchObj(Userindex)
                                           
            Case eGMCommands.LimpiarMundo
                Call HandleLimpiarMundo(Userindex)
        End Select
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & ". Paquete: " & Command)
End Sub

Private Sub HandleHome(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.TargetNpcTipo = eNPCType.Gobernador Then
            Call setHome(Userindex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
        Else
            If .flags.Muerto = 1 Then
                If (MapInfo(.Pos.Map).Restringir = eRestrict.restrict_no) And (.Counters.Pena = 0) Then
                    If .flags.Traveling = 0 Then
                        If Ciudades(.Hogar).Map <> .Pos.Map Then
                            Call goHome(Userindex)
                        Else
                            Call WriteConsoleMsg(Userindex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        Call WriteMultiMessage(Userindex, eMessages.CancelHome)
                        .flags.Traveling = 0
                        .Counters.goHome = 0
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes usar este comando aqui.", FontTypeNames.FONTTYPE_FIGHT)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

Private Sub HandleDeleteChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte
    Dim UserName    As String
    Dim AccountHash As String
    UserName = buffer.ReadASCIIString()
    AccountHash = buffer.ReadASCIIString()
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    Call BorrarUsuario(Userindex, UserName, AccountHash)
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.DeletedChar)
    Exit Sub
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleLoginExistingChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte
    Dim UserName    As String
    Dim AccountHash As String
    Dim version     As String
    UserName = buffer.ReadASCIIString()
    AccountHash = buffer.ReadASCIIString()
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(Userindex, "Nombre invalido.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "El personaje no existe.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If BANCheck(UserName) Then
        Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada a Argentum Online debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.argentumonline.org")
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call ConnectUser(Userindex, UserName, AccountHash)
    End If
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleThrowDices(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    With UserList(Userindex).Stats
        .UserAtributos(eAtributos.Fuerza) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Agilidad) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Inteligencia) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Carisma) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Constitucion) = RandomNumber(DiceMinimum, DiceMaximum)
    End With
    Call WriteDiceRoll(Userindex)
End Sub

Private Sub HandleLoginNewChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 15 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte
    Dim UserName    As String
    Dim AccountHash As String
    Dim version     As String
    Dim race        As eRaza
    Dim gender      As eGenero
    Dim homeland    As eCiudad
    Dim Class As eClass
    Dim Head As Integer
    UserName = buffer.ReadASCIIString()
    AccountHash = buffer.ReadASCIIString()
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    Head = buffer.ReadInteger
    homeland = buffer.ReadByte()
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(Userindex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(Userindex, "Servidor restringido a administradores. Consulte la pagina oficial o el foro oficial para mas informacion.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If aClon.MaxPersonajes(UserList(Userindex).IP) Then
        Call WriteErrorMsg(Userindex, "Has creado demasiados personajes.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call ConnectNewUser(Userindex, UserName, AccountHash, race, gender, Class, homeland, Head)
    End If
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleTalk(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
        End If
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        If LenB(Chat) <> 0 Then
            Call Statistics.ParseChat(Chat)
            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
                End If
            Else
                If RTrim(Chat) <> "" Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleYell(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)
        End If
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        If LenB(Chat) <> 0 Then
            Call Statistics.ParseChat(Chat)
            If .flags.Privilegios And PlayerType.User Then
                If UserList(Userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))
                End If
            Else
                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleWhisper(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat            As String
        Dim TargetUserIndex As Integer
        Dim TargetPriv      As PlayerType
        Dim UserPriv        As PlayerType
        Dim TargetName      As String
        TargetName = buffer.ReadASCIIString()
        Chat = buffer.ReadASCIIString()
        UserPriv = .flags.Privilegios
        If .flags.Muerto Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
        Else
            TargetUserIndex = NameIndex(TargetName)
            If TargetUserIndex = INVALID_INDEX Then
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                TargetPriv = UserList(TargetUserIndex).flags.Privilegios
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (UserPriv And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 And Not .flags.EnConsulta Then
                    Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                ElseIf (UserPriv And PlayerType.User) <> 0 And (Not TargetPriv And PlayerType.User) <> 0 And Not .flags.EnConsulta Then
                    Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                ElseIf Not EstaPCarea(Userindex, TargetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                    If (TargetPriv And (PlayerType.User)) = 0 And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                        Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                        Call WriteConsoleMsg(Userindex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If UserPriv And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le susurro a '" & UserList(TargetUserIndex).Name & "' " & Chat)
                    ElseIf (UserPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call LogGM(UserList(TargetUserIndex).Name, .Name & " le susurro en consulta: " & Chat)
                    End If
                    If LenB(Chat) <> 0 Then
                        Call Statistics.ParseChat(Chat)
                        If Not EstaPCarea(Userindex, TargetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                            Call WriteConsoleMsg(Userindex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(Userindex, Chat, .Char.CharIndex, vbBlue)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, vbBlue)
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, Userindex, PrepareMessageChatOverHead("A " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                            End If
                        Else
                            Call WriteConsoleMsg(Userindex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            If Userindex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, Userindex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleWalk(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim dummy    As Long
    Dim TempTick As Long
    Dim heading  As eHeading
    With UserList(Userindex)
        Call .incomingData.ReadByte
        heading = .incomingData.ReadByte()
        Dim TiempoDeWalk As Byte
        If .flags.Equitando = 1 Then
            TiempoDeWalk = 36
        Else
            TiempoDeWalk = 30
        End If
        If .flags.TimesWalk >= TiempoDeWalk Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then dummy = 126000 \ dummy
                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(Userindex)
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        .flags.TimesWalk = .flags.TimesWalk + 1
        Call CancelExit(Userindex)
        If .flags.Comerciando Then Exit Sub
        If .flags.Traveling = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes moverte mientras estas viajando a tu hogar con el comando /HOGAR.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
                Call MoveUserChar(Userindex, heading)
            Else
                Call MoveUserChar(Userindex, heading)
                If .flags.Descansar Then
                    .flags.Descansar = False
                    Call WriteRestOK(Userindex)
                    Call WriteConsoleMsg(Userindex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                Call WriteConsoleMsg(Userindex, "No puedes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            .flags.CountSH = 0
        End If
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .Clase <> eClass.Thief And .Clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.Navegando = 1 Then
                    If .Clase = eClass.Pirat Then
                        Call ToggleBoatBody(Userindex)
                        Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                    End If
                Else
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub HandleRequestPositionUpdate(ByVal Userindex As Integer)
    UserList(Userindex).incomingData.ReadByte
    Call WritePosUpdate(Userindex)
End Sub

Private Sub HandleAttack(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.Meditando Then
            Exit Sub
        End If
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(Userindex, "No puedes usar asi este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        Call CancelExit(Userindex)
        Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterAttackAnim(.Char.CharIndex))
        Call UsuarioAtaca(Userindex)
        .flags.NoPuedeSerAtacado = False
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
End Sub

Private Sub HandlePickUp(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then Exit Sub
        If .flags.Comerciando Then Exit Sub
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(Userindex, "No puedes tomar ningUn objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        Call GetObj(Userindex)
    End With
End Sub

Private Sub HandleSafeToggle(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Seguro Then
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOff)
        Else
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOn)
        End If
        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

Private Sub HandleResuscitationToggle(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        .flags.SeguroResu = Not .flags.SeguroResu
        If .flags.SeguroResu Then
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOn)
        Else
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOff)
        End If
    End With
End Sub

Private Sub HandleRequestGuildLeaderInfo(ByVal Userindex As Integer)
    UserList(Userindex).incomingData.ReadByte
    Call modGuilds.SendGuildLeaderInfo(Userindex)
End Sub

Private Sub HandleRequestAtributes(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call WriteAttributes(Userindex)
End Sub

Private Sub HandleRequestFame(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call EnviarFama(Userindex)
End Sub

Private Sub HandleRequestSkills(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call WriteSendSkills(Userindex)
End Sub

Private Sub HandleRequestMiniStats(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call WriteMiniStats(Userindex)
End Sub

Private Sub HandleCommerceEnd(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    UserList(Userindex).flags.Comerciando = False
    Call WriteCommerceEnd(Userindex)
End Sub

Private Sub HandleUserCommerceEnd(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)
            End If
        End If
        Call FinComerciarUsu(Userindex)
        Call WriteConsoleMsg(Userindex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
    End With
End Sub

Private Sub HandleUserCommerceConfirm(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    If PuedeSeguirComerciando(Userindex) Then
        Call WriteUserOfferConfirm(UserList(Userindex).ComUsu.DestUsu)
        UserList(Userindex).ComUsu.Confirmo = True
    End If
End Sub

Private Sub HandleCommerceChat(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(Userindex) Then
                Call Statistics.ParseChat(Chat)
                Chat = UserList(Userindex).Name & "> " & Chat
                Call WriteCommerceChat(Userindex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(Userindex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleBankEnd(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        .flags.Comerciando = False
        Call WriteBankEnd(Userindex)
    End With
End Sub

Private Sub HandleUserCommerceOk(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call AceptarComercioUsu(Userindex)
End Sub

Private Sub HandleUserCommerceReject(ByVal Userindex As Integer)
    Dim otherUser As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        otherUser = .ComUsu.DestUsu
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
            End If
        End If
        Call WriteConsoleMsg(Userindex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(Userindex)
    End With
End Sub

Private Sub HandleDrop(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim Slot   As Byte
    Dim Amount As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        If .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub
        If .flags.Comerciando Then Exit Sub
        If .flags.Navegando = 1 And Not .Clase = eClass.Pirat Then
            Call WriteConsoleMsg(Userindex, "Solo los Piratas pueden tirar items en altamar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub
            Call TirarOro(Amount, Userindex)
            Call WriteUpdateGold(Userindex)
        Else
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                Call DropObj(Userindex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

Private Sub HandleCastSpell(ByVal Userindex As Integer)

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Spell As Byte
        Spell = .incomingData.ReadByte()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        .flags.NoPuedeSerAtacado = False
        If Spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub
        End If
        .flags.Hechizo = .Stats.UserHechizos(Spell)
    End With
End Sub

Private Sub HandleLeftClick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim X As Byte
        Dim Y As Byte
        X = .ReadByte()
        Y = .ReadByte()
        Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)
    End With
End Sub

Private Sub HandleDoubleClick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim X As Byte
        Dim Y As Byte
        X = .ReadByte()
        Y = .ReadByte()
        Call Accion(Userindex, UserList(Userindex).Pos.Map, X, Y)
    End With
End Sub

Private Sub HandleWork(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Skill As eSkill
        Skill = .incomingData.ReadByte()
        If UserList(Userindex).flags.Muerto = 1 Then Exit Sub
        Call CancelExit(Userindex)
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, Skill)
                
            Case Ocultarse
                If (MapInfo(.Pos.Map).OcultarSinEfecto = 1) Or (MapInfo(.Pos.Map).InviSinEfecto = 1) Then
                    Call WriteConsoleMsg(Userindex, "Ocultarse no funciona aqui!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(Userindex, "No puedes ocultarte si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.Navegando = 1 Then
                    If .Clase <> eClass.Pirat Then
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(Userindex, "No puedes ocultarte si estas navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3
                        End If
                        Exit Sub
                    End If
                End If
                If .flags.Oculto = 1 Then
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(Userindex, "Ya estas oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    Exit Sub
                End If
                Call DoOcultarse(Userindex)
        End Select
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en HandleWork en " & Erl & " - Skill: " & Skill & ". Err: " & Err.Number & " " & Err.description)
End Sub

Private Sub HandleInitCrafting(ByVal Userindex As Integer)
    Dim TotalItems    As Long
    Dim ItemsPorCiclo As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger
        If TotalItems > 0 Then
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(Userindex), ItemsPorCiclo)
        End If
    End With
End Sub

Private Sub HandleUseSpellMacro(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Call SendData(SendTarget.ToAdmins, Userindex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_FIGHT))
        Call WriteErrorMsg(Userindex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call CloseSocket(Userindex)
    End With
End Sub

Private Sub HandleUseItem(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Slot As Byte
        Slot = .incomingData.ReadByte()
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        If .flags.Meditando Then
            Exit Sub
        End If
        Call UseInvItem(Userindex, Slot)
    End With
End Sub

Private Sub HandleCraftBlacksmith(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim Item As Integer
        Item = .ReadInteger()
        If Item < 1 Then Exit Sub
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
        Call HerreroConstruirItem(Userindex, Item)
    End With
End Sub

Private Sub HandleCraftCarpenter(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim Item As Integer
        Item = .ReadInteger()
        If Item < 1 Then Exit Sub
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
        Call CarpinteroConstruirItem(Userindex, Item)
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en HandleCraftcarpenter en " & Erl & " - Item: " & Item & ". Err " & Err.Number & " " & Err.description)
End Sub

Private Sub HandleWorkLeftClick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim X           As Byte
        Dim Y           As Byte
        Dim Skill       As eSkill
        Dim DummyInt    As Integer
        Dim tU          As Integer
        Dim tN          As Integer
        Dim WeaponIndex As Integer
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Skill = .incomingData.ReadByte()
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
        If Not InRangoVision(Userindex, X, Y) Then
            Call WritePosUpdate(Userindex)
            Exit Sub
        End If
        Call CancelExit(Userindex)
        Select Case Skill
            Case eSkill.Proyectiles
                If Not IntervaloPermiteAtacar(Userindex, False) Then Exit Sub
                If Not IntervaloPermiteLanzarSpell(Userindex, False) Then Exit Sub
                If Not IntervaloPermiteUsarArcos(Userindex) Then Exit Sub
                Call LanzarProyectil(Userindex, X, Y)
                            
            Case eSkill.Magia
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(Userindex, "Una fuerza oscura te impide canalizar tu energia.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                If Not IntervaloPermiteUsarArcos(Userindex, False) Then Exit Sub
                If Not IntervaloPermiteGolpeMagia(Userindex) Then
                    If Not IntervaloPermiteLanzarSpell(Userindex) Then
                        Exit Sub
                    End If
                End If
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, Userindex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(Userindex, "Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.pesca
                WeaponIndex = .Invent.WeaponEqpObjIndex
                If WeaponIndex = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.CASA Then
                    Call WriteConsoleMsg(Userindex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If HayAgua(.Pos.Map, X, Y) Then
                    Select Case WeaponIndex
                        Case CANA_PESCA, CANA_PESCA_NEWBIE
                            Call DoPescar(Userindex)
                        
                        Case RED_PESCA
                            DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            If DummyInt = 0 Then
                                Call WriteConsoleMsg(Userindex, "No hay un yacimiento de peces donde pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(Userindex, "Estas demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            If .Pos.X = X And .Pos.Y = Y Then
                                Call WriteConsoleMsg(Userindex, "No puedes pescar desde alli.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            If ObjData(DummyInt).OBJType = eOBJType.otYacimientoPez Then
                                Call DoPescarRed(Userindex)
                            Else
                                Call WriteConsoleMsg(Userindex, "No hay un yacimiento de peces donde pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        Case Else
                            Exit Sub
                    End Select
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(Userindex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                If MapInfo(.Pos.Map).Pk Then
                    If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                    Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)
                    tU = .flags.TargetUser
                    If tU > 0 And tU <> Userindex Then
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                    Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                End If
                                If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(Userindex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub
                                End If
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(Userindex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub
                                End If
                                Call DoRobar(Userindex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eSkill.Talar
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                WeaponIndex = .Invent.WeaponEqpObjIndex
                If WeaponIndex = 0 Then
                    Call WriteConsoleMsg(Userindex, "Deberias equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If WeaponIndex <> HACHA_LENADOR And WeaponIndex <> HACHA_LENA_ELFICA And WeaponIndex <> HACHA_LENADOR_NEWBIE Then
                    Exit Sub
                End If
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(Userindex, "No puedes talar desde alli.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                        If WeaponIndex = HACHA_LENADOR Or WeaponIndex = HACHA_LENADOR_NEWBIE Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(Userindex)
                        Else
                            Call WriteConsoleMsg(Userindex, "No puedes extraer lena de este arbol con este hacha.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    ElseIf ObjData(DummyInt).OBJType = eOBJType.otArbolElfico Then
                        If WeaponIndex = HACHA_LENA_ELFICA Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(Userindex, True)
                        Else
                            Call WriteConsoleMsg(Userindex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No hay ningUn arbol ahi.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                WeaponIndex = .Invent.WeaponEqpObjIndex
                If WeaponIndex = 0 Then Exit Sub
                If WeaponIndex <> PIQUETE_MINERO And WeaponIndex <> PIQUETE_MINERO_NEWBIE Then
                    Exit Sub
                End If
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ahi no hay ningUn yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_MINERO, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(Userindex, "Ahi no hay ningUn yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eSkill.Domar
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(Userindex, "No puedes domar una criatura que esta luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call DoDomar(Userindex, tN)
                    Else
                        Call WriteConsoleMsg(Userindex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No hay ninguna criatura alli!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
                            Exit Sub
                        End If
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(Userindex, "No tienes mas minerales.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            Call WriteErrorMsg(Userindex, "Has sido expulsado por el sistema anti cheats.")
                            Call CloseSocket(Userindex)
                            Exit Sub
                        End If
                        If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
                            Call FundirMineral(Userindex)
                        ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                            Call FundirArmas(Userindex)
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "Ahi no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Ahi no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eSkill.Herreria
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(Userindex)
                        Call EnivarArmadurasConstruibles(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

Private Sub HandleCreateNewGuild(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 9 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Desc      As String
        Dim GuildName As String
        Dim Site      As String
        Dim codex()   As String
        Dim errorStr  As String
        Desc = buffer.ReadASCIIString()
        GuildName = Trim$(buffer.ReadASCIIString())
        Site = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        If modGuilds.CrearNuevoClan(Userindex, Desc, GuildName, Site, codex, .FundandoGuildAlineacion, errorStr) Then
            Dim Message As String
            Message = .Name & " fundo el clan " & GuildName & " de alineacion " & modGuilds.GuildAlignment(.GuildIndex)
            Call SendData(SendTarget.ToAll, Userindex, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            Call RefreshCharStatus(Userindex)
            If ConexionAPI Then
                Call ApiEndpointSendNewGuildCreatedMessageDiscord(Message, Desc, GuildName, Site)
            End If
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleEquipItem(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim itemSlot As Byte
        itemSlot = .incomingData.ReadByte()
        If .flags.Muerto = 1 Then Exit Sub
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        Call EquiparInvItem(Userindex, itemSlot)
    End With
End Sub

Private Sub HandleChangeHeading(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim heading As eHeading
        Dim posX    As Integer
        Dim posY    As Integer
        heading = .incomingData.ReadByte()
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1

                Case eHeading.EAST
                    posX = 1

                Case eHeading.SOUTH
                    posY = 1

                Case eHeading.WEST
                    posX = -1
            End Select
            If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub
            End If
        End If
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

Private Sub HandleModifySkills(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 1 + NUMSKILLS Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim i                      As Long
        Dim Count                  As Integer
        Dim points(1 To NUMSKILLS) As Byte
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            If points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .IP & " trato de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(Userindex)
                Exit Sub
            End If
            Count = Count + points(i)
        Next i
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .IP & " trato de hackear los skills.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        With .Stats
            For i = 1 To NUMSKILLS
                If points(i) > 0 Then
                    .SkillPts = .SkillPts - points(i)
                    .UserSkills(i) = .UserSkills(i) + points(i)
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100
                    End If
                    Call CheckEluSkill(Userindex, i, True)
                End If
            Next i
        End With
    End With
End Sub

Private Sub HandleTrain(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim SpawnedNpc As Integer
        Dim PetIndex   As Byte
        PetIndex = .incomingData.ReadByte()
        If .flags.TargetNPC = 0 Then Exit Sub
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No puedo traer mas criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

Private Sub HandleCommerceBuy(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Slot   As Byte
        Dim Amount As Integer
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC < 1 Then Exit Sub
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "No estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call Comercio(eModoComercio.Compra, Userindex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

Private Sub HandleBankExtractItem(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Slot   As Byte
        Dim Amount As Integer
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC < 1 Then Exit Sub
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        Call UserRetiraItem(Userindex, Slot, Amount)
    End With
End Sub

Private Sub HandleCommerceSell(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Slot   As Byte
        Dim Amount As Integer
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC < 1 Then Exit Sub
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        Call Comercio(eModoComercio.Venta, Userindex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

Private Sub HandleBankDeposit(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Slot   As Byte
        Dim Amount As Integer
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC < 1 Then Exit Sub
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        Call UserDepositaItem(Userindex, Slot, Amount)
    End With
End Sub

Private Sub HandleForumPost(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim ForumMsgType As eForumMsgType
        Dim File         As String
        Dim Title        As String
        Dim Post         As String
        Dim ForumIndex   As Integer
        Dim postFile     As String
        Dim ForumType    As Byte
        ForumMsgType = buffer.ReadByte()
        Title = buffer.ReadASCIIString()
        Post = buffer.ReadASCIIString()
        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)
            Select Case ForumType
                Case eForumType.ieGeneral
                    ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)
                    
                Case eForumType.ieREAL
                    ForumIndex = GetForumIndex(FORO_REAL_ID)
                    
                Case eForumType.ieCAOS
                    ForumIndex = GetForumIndex(FORO_CAOS_ID)
            End Select
            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleMoveSpell(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim Dir As Integer
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1
        End If
        Call DesplazarHechizo(Userindex, Dir, .ReadByte())
    End With
End Sub

Private Sub HandleMoveBank(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim Dir      As Integer
        Dim Slot     As Byte
        Dim TempItem As obj
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1
        End If
        Slot = .ReadByte()
    End With
    With UserList(Userindex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        If Dir = 1 Then
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
            Call UpdateBanUserInv(False, Userindex, Slot - 1)
        Else
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
            Call UpdateBanUserInv(False, Userindex, Slot + 1)
        End If
        Call UpdateBanUserInv(False, Userindex, Slot)
    End With
End Sub

Private Sub HandleClanCodexUpdate(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Desc    As String
        Dim codex() As String
        Desc = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        Call modGuilds.ChangeCodexAndDesc(Desc, codex, .GuildIndex)
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleUserCommerceOffer(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Amount    As Long
        Dim Slot      As Byte
        Dim tUser     As Integer
        Dim OfferSlot As Byte
        Dim ObjIndex  As Integer
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        tUser = .ComUsu.DestUsu
        If UserList(Userindex).ComUsu.Confirmo = True Then
            Call FinComerciarUsu(Userindex)
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
            End If
            Exit Sub
        End If
        If ((Slot < 0 Or Slot > UserList(Userindex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        If Amount = 0 Then Exit Sub
        If Slot = FLAGORO Then
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(Userindex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.GoldAmount Then
                    Amount = .ComUsu.GoldAmount * (-1)
                End If
            End If
        Else
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex
            If Not HasEnoughItems(Userindex, ObjIndex, TotalOfferItems(ObjIndex, Userindex) + Amount) Then
                Call WriteCommerceChat(Userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)
                End If
            End If
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(Userindex, OfferSlot)
                Exit Sub
            End If
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(Userindex, "No puedes vender tu barco mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            If .flags.Equitando = 1 Then
                If .Invent.MonturaEqpSlot = Slot Then
                    Call WriteConsoleMsg(Userindex, "No podes vender tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteCommerceChat(Userindex, "No puedes vender tu alforja o mochila mientras la estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
        End If
        Call AgregarOferta(Userindex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
        Call EnviarOferta(tUser, OfferSlot)
    End With
End Sub

Private Sub HandleGuildAcceptPeace(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        Guild = buffer.ReadASCIIString()
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(Userindex, Guild, errorStr)
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildRejectAlliance(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        Guild = buffer.ReadASCIIString()
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(Userindex, Guild, errorStr)
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & Guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildRejectPeace(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        Guild = buffer.ReadASCIIString()
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(Userindex, Guild, errorStr)
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildAcceptAlliance(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        Guild = buffer.ReadASCIIString()
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(Userindex, Guild, errorStr)
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildOfferPeace(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild    As String
        Dim proposal As String
        Dim errorStr As String
        Guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        If modGuilds.r_ClanGeneraPropuesta(Userindex, Guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildOfferAlliance(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild    As String
        Dim proposal As String
        Dim errorStr As String
        Guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        If modGuilds.r_ClanGeneraPropuesta(Userindex, Guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub
Private Sub HandleGuildAllianceDetails(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild    As String
        Dim errorStr As String
        Dim details  As String
        Guild = buffer.ReadASCIIString()
        details = modGuilds.r_VerPropuesta(Userindex, Guild, RELACIONES_GUILD.ALIADOS, errorStr)
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildPeaceDetails(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild    As String
        Dim errorStr As String
        Dim details  As String
        Guild = buffer.ReadASCIIString()
        details = modGuilds.r_VerPropuesta(Userindex, Guild, RELACIONES_GUILD.PAZ, errorStr)
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildRequestJoinerInfo(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim User    As String
        Dim details As String
        User = buffer.ReadASCIIString()
        details = modGuilds.a_DetallesAspirante(Userindex, User)
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, "El personaje no ha mandado solicitud, o no estas habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(Userindex, details)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildAlliancePropList(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call WriteAlianceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.ALIADOS))
End Sub

Private Sub HandleGuildPeacePropList(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call WritePeaceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.PAZ))
End Sub

Private Sub HandleGuildDeclareWar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild           As String
        Dim errorStr        As String
        Dim otherGuildIndex As Integer
        Guild = buffer.ReadASCIIString()
        otherGuildIndex = modGuilds.r_DeclararGuerra(Userindex, Guild, errorStr)
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildNewWebsite(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Call modGuilds.ActualizarWebSite(Userindex, buffer.ReadASCIIString())
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildAcceptNewMember(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim errorStr As String
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If Not modGuilds.a_AceptarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildRejectNewMember(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim errorStr As String
        Dim UserName As String
        Dim Reason   As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        If Not modGuilds.a_RechazarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildKickMember(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName   As String
        Dim GuildIndex As Integer
        UserName = buffer.ReadASCIIString()
        GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(Userindex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildUpdateNews(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Call modGuilds.ActualizarNoticias(Userindex, buffer.ReadASCIIString())
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildMemberInfo(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Call modGuilds.SendDetallesPersonaje(Userindex, buffer.ReadASCIIString())
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildOpenElections(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Error As String
        If Not modGuilds.v_AbrirElecciones(Userindex, Error) Then
            Call WriteConsoleMsg(Userindex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

Private Sub HandleGuildRequestMembership(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild       As String
        Dim application As String
        Dim errorStr    As String
        Guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        If Not modGuilds.a_NuevoAspirante(Userindex, Guild, application, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Tu solicitud ha sido enviada. Espera prontas noticias del lider de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildRequestDetails(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Call modGuilds.SendGuildDetails(Userindex, buffer.ReadASCIIString())
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub WriteConsoleServerUpTimeMsg(ByVal Userindex As Integer)
    Dim time As Long
    Dim UpTimeStr As String
    time = ((GetTickCount()) - tInicioServer) \ 1000
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    If time = 1 Then
        UpTimeStr = time & " dia, " & UpTimeStr
    Else
        UpTimeStr = time & " dias, " & UpTimeStr
    End If
    Call WriteConsoleMsg(Userindex, "Tiempo del Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Sub HandleOnline(ByVal Userindex As Integer)
    Dim SB As cStringBuilder
    Set SB = New cStringBuilder
    Dim i     As Long
    Dim Count As Long
    Dim CountTrabajadores As Long
    With UserList(Userindex)
        Call .incomingData.ReadByte
        For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                Call SB.Append(UserList(i).Name)
                If UserList(i).Clase = eClass.Worker Then
                    If EsGm(Userindex) Or (.Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Supervivencia) = 100) Then
                        CountTrabajadores = CountTrabajadores + 1
                        Call SB.Append(" [T]")
                    End If
                End If
                If i <> LastUser Then
                    Call SB.Append(", ")
                End If
                Count = Count + 1
            End If
        Next i
        Call WriteConsoleMsg(Userindex, SB.toString, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, "Usuarios en linea: " & CStr(Count), FontTypeNames.FONTTYPE_INFOBOLD)
        If EsGm(Userindex) Or (.Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Supervivencia) = 100) Then
            Call WriteConsoleMsg(Userindex, "Trabajadores en linea:" & CStr(CountTrabajadores), FontTypeNames.FONTTYPE_INFOBOLD)
        End If
        Set SB = Nothing
    End With
    Call WriteConsoleServerUpTimeMsg(Userindex)
End Sub
Private Sub HandleQuit(ByVal Userindex As Integer)
    Dim tUser        As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = Userindex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_WARNING)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            Call WriteConsoleMsg(Userindex, "Comercio cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Call FinComerciarUsu(Userindex)
        End If
        Call Cerrar_Usuario(Userindex)
    End With
End Sub

Private Sub HandleGuildLeave(ByVal Userindex As Integer)
    Dim GuildIndex As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        GuildIndex = m_EcharMiembroDeClan(Userindex, .Name)
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(Userindex, "Tu no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

Private Sub HandleRequestAccountState(ByVal Userindex As Integer)
    Dim earnings   As Integer
    Dim Percentage As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                Call WriteChatOverHead(Userindex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    Call WriteConsoleMsg(Userindex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

Private Sub HandlePetStand(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

Private Sub HandlePetFollow(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        Call FollowAmo(.flags.TargetNPC)
        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

Private Sub HandleReleasePet(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar una mascota, haz click izquierdo sobre ella.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call QuitarPet(Userindex, .flags.TargetNPC)
    End With
End Sub

Private Sub HandleTrainList(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        Call WriteTrainerCreatureList(Userindex, .flags.TargetNPC)
    End With
End Sub

Private Sub HandleRest(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(Userindex)
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(Userindex, "Te acomodas junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(Userindex)
                Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                .flags.Descansar = False
                Exit Sub
            End If
            Call WriteConsoleMsg(Userindex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleMeditate(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.Equitando Then
            Call WriteConsoleMsg(Userindex, "No puedes meditar mientras si estas montado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Stats.MaxMAN = 0 Then
            Call WriteConsoleMsg(Userindex, "Solo las clases magicas conocen el arte de la meditacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(Userindex, "Mana restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(Userindex)
            Exit Sub
        End If
        Call WriteMeditateToggle(Userindex)
        If .flags.Meditando Then Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        .flags.Meditando = Not .flags.Meditando
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount()
            Call WriteConsoleMsg(Userindex, "Te estas concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzaras a meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.loops = INFINITE_LOOPS
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            ElseIf .Stats.ELV < 35 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
            End If
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

Private Sub HandleResucitate(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Or .flags.Muerto = 0 Then Exit Sub
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 5 Then
            Call WriteConsoleMsg(Userindex, "El sacerdote no puede resucitarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call SacerdoteResucitateUser(Userindex)
    End With
End Sub

Private Sub HandleConsultation(ByVal Userindex As String)
    Dim UserConsulta As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If Not EsGm(Userindex) Then Exit Sub
        UserConsulta = .flags.TargetUser
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserConsulta = Userindex Then Exit Sub
        If EsGm(UserConsulta) Then
            Call WriteConsoleMsg(Userindex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim UserName As String
        UserName = UserList(UserConsulta).Name
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(Userindex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & UserName)
            UserList(UserConsulta).flags.EnConsulta = False
        Else
            Call WriteConsoleMsg(Userindex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & UserName)
            With UserList(UserConsulta)
                .flags.EnConsulta = True
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    If UserList(UserConsulta).flags.Navegando = 0 Then
                        Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
                    End If
                End If
            End With
        End If
        Call UsUaRiOs.SetConsulatMode(UserConsulta)
    End With
End Sub

Private Sub HandleHeal(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor) Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call SacerdoteHealUser(Userindex)
    End With
End Sub

Private Sub HandleRequestStats(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call SendUserStatsTxt(Userindex, Userindex)
End Sub

Private Sub HandleHelp(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call SendHelp(Userindex)
End Sub

Private Sub HandleCommerceStart(ByVal Userindex As Integer)
    Dim i As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.TargetNPC > 0 Then
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(Userindex, "No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                Exit Sub
            End If
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call IniciarComercioNPC(Userindex)
        ElseIf .flags.TargetUser > 0 Then
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(Userindex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If .flags.TargetUser = Userindex Then
                Call WriteConsoleMsg(Userindex, "No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> Userindex Then
                Call WriteConsoleMsg(Userindex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name
            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i
            .ComUsu.GoldAmount = 0
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            Call IniciarComercioConUsuario(Userindex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(Userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleBankStart(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(Userindex)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleEnlist(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Debes acercarte mas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(Userindex)
        Else
            Call EnlistarCaos(Userindex)
        End If
    End With
End Sub

Private Sub HandleInformation(ByVal Userindex As Integer)
    Dim Matados    As Integer
    Dim NextRecom  As Integer
    Dim Diferencia As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        NextRecom = .Faccion.NextRecompensa
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(Userindex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados
            If Diferencia > 0 Then
                Call WriteChatOverHead(Userindex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(Userindex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        Else
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(Userindex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados
            If Diferencia > 0 Then
                Call WriteChatOverHead(Userindex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(Userindex, "Tu deber es sembrar el caos y la desesperanza, y creo que estas en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        End If
    End With
End Sub

Private Sub HandleReward(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(Userindex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            Call RecompensaArmadaReal(Userindex)
        Else
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(Userindex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            Call RecompensaCaos(Userindex)
        End If
    End With
End Sub

Private Sub HandleRequestMOTD(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call SendMOTD(Userindex)
End Sub

Private Sub HandleUpTime(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Dim time      As Long
    Dim UpTimeStr As String
    Call WriteConsoleServerUpTimeMsg(Userindex)
End Sub

Private Sub HandlePartyLeave(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call mdParty.SalirDeParty(Userindex)
End Sub

Private Sub HandlePartyCreate(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    If Not mdParty.PuedeCrearParty(Userindex) Then Exit Sub
    Call mdParty.CrearParty(Userindex)
End Sub

Private Sub HandlePartyJoin(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call mdParty.SolicitarIngresoAParty(Userindex)
End Sub

Private Sub HandleShareNpc(ByVal Userindex As Integer)
    Dim TargetUserIndex  As Integer
    Dim SharingUserIndex As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        TargetUserIndex = .flags.TargetUser
        If TargetUserIndex = 0 Then Exit Sub
        If EsGm(TargetUserIndex) Then
            Call WriteConsoleMsg(Userindex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If criminal(Userindex) Then
            If esCaos(Userindex) Then
                If Not esCaos(TargetUserIndex) Then
                    Call WriteConsoleMsg(Userindex, "Solo puedes compartir npcs con miembros de tu misma faccion!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        Else
            If criminal(TargetUserIndex) Then
                Call WriteConsoleMsg(Userindex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        SharingUserIndex = .flags.ShareNpcWith
        If SharingUserIndex = TargetUserIndex Then Exit Sub
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        End If
        .flags.ShareNpcWith = TargetUserIndex
        Call WriteConsoleMsg(TargetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, "Ahora compartes tus npcs con " & UserList(TargetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleStopSharingNpc(ByVal Userindex As Integer)
    Dim SharingUserIndex As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        SharingUserIndex = .flags.ShareNpcWith
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            .flags.ShareNpcWith = 0
        End If
    End With
End Sub

Private Sub HandleInquiry(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    ConsultaPopular.SendInfoEncuesta (Userindex)
End Sub

Private Sub HandleGuildMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If LenB(Chat) <> 0 Then
            Call Statistics.ParseChat(Chat)
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
                If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToClanArea, Userindex, PrepareMessageChatOverHead("< " & Chat & " >", .Char.CharIndex, vbYellow))
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandlePartyMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If LenB(Chat) <> 0 Then
            Call Statistics.ParseChat(Chat)
            Call mdParty.BroadCastParty(Userindex, Chat)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCentinelReport(ByVal Userindex As Integer)
    Dim NotBuff As New clsByteQueue
    With UserList(Userindex)
        Call NotBuff.CopyBuffer(.incomingData)
        Call NotBuff.ReadByte
        Call modCentinela.IngresaClave(Userindex, NotBuff.ReadASCIIString())
        Call .incomingData.CopyBuffer(NotBuff)
    End With
End Sub

Private Sub HandleGuildOnline(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim onlineList As String
        onlineList = modGuilds.m_ListaDeMiembrosOnline(Userindex, .GuildIndex)
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(Userindex, "Companeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(Userindex, "No pertences a ningUn clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub

Private Sub HandlePartyOnline(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call mdParty.OnlineParty(Userindex)
End Sub

Private Sub HandleCouncilMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If LenB(Chat) <> 0 Then
            Call Statistics.ParseChat(Chat)
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRoleMasterRequest(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim request As String
        request = buffer.ReadASCIIString()
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(Userindex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGMRequest(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(Userindex, "El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(Userindex, "Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name + " ha solicitado la ayuda de algun GM con /GM. Podes usar el comando /SHOW SOS para ver quienes necesitan ayuda", FontTypeNames.FONTTYPE_INFO))
    End With
End Sub

Private Sub HandleBugReport(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Dim n As Integer
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim bugReport As String
        bugReport = buffer.ReadASCIIString()
        n = FreeFile
        Open App.Path & "\LOGS\BUGs.log" For Append Shared As n
        Print #n, "Usuario:" & .Name & "  Fecha:" & Date & "    Hora:" & time
        Print #n, "BUG:"
        Print #n, bugReport
        Print #n, "########################################################################"
        Close #n
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleChangeDescription(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim description As String
        description = buffer.ReadASCIIString()
        If Not AsciiValidos(description) Then
            Call WriteConsoleMsg(Userindex, "La descripcion tiene caracteres invalidos.", FontTypeNames.FONTTYPE_INFO)
        Else
            .Desc = Trim$(description)
            Call WriteConsoleMsg(Userindex, "La descripcion ha cambiado.", FontTypeNames.FONTTYPE_INFO)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildVote(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim vote     As String
        Dim errorStr As String
        vote = buffer.ReadASCIIString()
        If Not modGuilds.v_UsuarioVota(Userindex, vote, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleShowGuildNews(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Call modGuilds.SendGuildNews(Userindex)
    End With
End Sub

Private Sub HandlePunishments(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Name  As String
        Dim Count As Integer
        Name = buffer.ReadASCIIString()
        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")
            End If
            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")
            End If
            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")
            End If
            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")
            End If
            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(Userindex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(Userindex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else
                If PersonajeExiste(Name) Then
                    Count = GetUserAmountOfPunishments(Name)
                    If Count = 0 Then
                        Call WriteConsoleMsg(Userindex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call SendUserPunishments(Userindex, Name, Count)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleChangePassword(ByVal Userindex As Integer)
    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Dim oldSalt    As String
        Dim Salt       As String
        Dim oldPass    As String
        Dim newPass    As String
        Dim storedPass As String
        Call buffer.ReadByte
        oldSalt = GetUserSalt(UserList(Userindex).Name)
        oldPass = oSHA256.SHA256(buffer.ReadASCIIString() & oldSalt)
        Salt = RandomString(10)
        newPass = oSHA256.SHA256(buffer.ReadASCIIString() & Salt)
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(Userindex, "Debes especificar una contrasena nueva, intentalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            storedPass = GetUserPassword(UserList(Userindex).Name)
            If storedPass <> oldPass Then
                Call WriteConsoleMsg(Userindex, "La contrasena actual proporcionada no es correcta. La contrasena no ha sido cambiada, intentalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call StorePasswordSalt(UserList(Userindex).Name, newPass, Salt)
                Call WriteConsoleMsg(Userindex, "La contrasena fue cambiada con exito.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGamble(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Amount  As Integer
        Dim TypeNpc As eNPCType
        Amount = .incomingData.ReadInteger()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Dim TargetNpcType As eNPCType
            TargetNpcType = Npclist(.flags.TargetNPC).NPCtype
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
                Call WriteChatOverHead(Userindex, "No tengo ningUn interes en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(Userindex, "El minimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(Userindex, "El maximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(Userindex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(Userindex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
                Call WriteChatOverHead(Userindex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            Call WriteUpdateGold(Userindex)
        End If
    End With
End Sub

Private Sub HandleInquiryVote(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim opt As Byte
        opt = .incomingData.ReadByte()
        Call WriteConsoleMsg(Userindex, ConsultaPopular.doVotar(Userindex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

Private Sub HandleBankExtractGold(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Amount As Long
        Amount = .incomingData.ReadLong()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Amount > 0 And Amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.Gld = .Stats.Gld + Amount
            Call WriteChatOverHead(Userindex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(Userindex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        Call WriteUpdateGold(Userindex)
        Call WriteUpdateBankGold(Userindex)
    End With
End Sub

Private Sub HandleLeaveFaction(ByVal Userindex As Integer)
    Dim TalkToKing  As Boolean
    Dim TalkToDemon As Boolean
    Dim NpcIndex    As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        NpcIndex = .flags.TargetNPC
        If NpcIndex <> 0 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
                If Npclist(NpcIndex).flags.Faccion = 0 Then
                    TalkToKing = True
                Else
                    TalkToDemon = True
                End If
            End If
        End If
        If .Faccion.ArmadaReal = 1 Then
            If TalkToDemon Then
                Call WriteChatOverHead(Userindex, "Sal de aqui bufon!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                If TalkToKing Then
                    Call WriteChatOverHead(Userindex, "Seras bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                Call ExpulsarFaccionReal(Userindex, False)
            End If
        ElseIf .Faccion.FuerzasCaos = 1 Then
            If TalkToKing Then
                Call WriteChatOverHead(Userindex, "Sal de aqui maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                If TalkToDemon Then
                    Call WriteChatOverHead(Userindex, "Ya volveras arrastrandote.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                Call ExpulsarFaccionCaos(Userindex, False)
            End If
        Else
            If (TalkToDemon And criminal(Userindex)) Or (TalkToKing And Not criminal(Userindex)) Then
                Call WriteChatOverHead(Userindex, "No perteneces a nuestra faccion. Si deseas unirte, di /ENLISTAR", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            ElseIf (TalkToDemon And Not criminal(Userindex)) Then
                Call WriteChatOverHead(Userindex, "Sal de aqui bufon!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            ElseIf (TalkToKing And criminal(Userindex)) Then
                Call WriteChatOverHead(Userindex, "Sal de aqui maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                Call WriteConsoleMsg(Userindex, "No perteneces a ninguna faccion!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End With
End Sub

Private Sub HandleBankDepositGold(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Amount As Long
        Amount = .incomingData.ReadLong()
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        Dim RemainingAmountToMaximumGold As Long
        RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld
        If .Stats.Banco >= 2147483647 And RemainingAmountToMaximumGold <= Amount Then
            Call WriteChatOverHead(Userindex, "No puedes depositar el oro por que tendrias mas del maximo permitido (2147483647)", Npclist(.flags.TargetNPC).Char.CharIndex, vbRed)
        ElseIf Amount > 0 And Amount <= .Stats.Gld Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.Gld = .Stats.Gld - Amount
            Call WriteChatOverHead(Userindex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Call WriteUpdateGold(Userindex)
            Call WriteUpdateBankGold(Userindex)
        Else
            Call WriteChatOverHead(Userindex, "No tenes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub
Private Sub HandleDenounce(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Text As String
        Dim Msg  As String
        Text = buffer.ReadASCIIString()
        If .flags.Silenciado = 0 Then
            Call Statistics.ParseChat(Text)
            Msg = LCase$(.Name) & " DENUNCIA: " & Text
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
            Call Denuncias.Push(Msg, False)
            Call WriteConsoleMsg(Userindex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildFundate(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 1 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If HasFound(.Name) Then
            Call WriteConsoleMsg(Userindex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
        Call WriteShowGuildAlign(Userindex)
    End With
End Sub
    
Private Sub HandleGuildFundation(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim clanType As eClanType
        Dim Error    As String
        clanType = .incomingData.ReadByte()
        If HasFound(.Name) Then
            Call WriteConsoleMsg(Userindex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .IP)
            Exit Sub
        End If
        Select Case UCase$(Trim(clanType))
            Case eClanType.ct_RoyalArmy
                .FundandoGuildAlineacion = ALINEACION_ARMADA

            Case eClanType.ct_Evil
                .FundandoGuildAlineacion = ALINEACION_LEGION

            Case eClanType.ct_Neutral
                .FundandoGuildAlineacion = ALINEACION_NEUTRO

            Case eClanType.ct_GM
                .FundandoGuildAlineacion = ALINEACION_MASTER

            Case eClanType.ct_Legal
                .FundandoGuildAlineacion = ALINEACION_CIUDA

            Case eClanType.ct_Criminal
                .FundandoGuildAlineacion = ALINEACION_CRIMINAL

            Case Else
                Call WriteConsoleMsg(Userindex, "Alineacion invalida.", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
        End Select
        If modGuilds.PuedeFundarUnClan(Userindex, .FundandoGuildAlineacion, Error) Then
            Call WriteShowGuildFundationForm(Userindex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(Userindex, Error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

Private Sub HandlePartyKick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If UserPuedeEjecutarComandos(Userindex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(Userindex, tUser)
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                Call WriteConsoleMsg(Userindex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandlePartySetLeader(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim rank     As Integer
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        If UserPuedeEjecutarComandos(Userindex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.TransformarEnLider(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, LCase(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                Call WriteConsoleMsg(Userindex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandlePartyAcceptMember(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName  As String
        Dim tUser     As Integer
        Dim rank      As Integer
        Dim bUserVivo As Boolean
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        If UserList(Userindex).flags.Muerto Then
            Call WriteConsoleMsg(Userindex, "Estas muerto!!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True
        End If
        If mdParty.UserPuedeEjecutarComandos(Userindex) And bUserVivo Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                If (UserList(tUser).flags.Privilegios And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.AprobarIngresoAParty(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(Userindex, LCase(UserName) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGuildMemberList(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild       As String
        Dim memberCount As Integer
        Dim i           As Long
        Dim UserName    As String
        Guild = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(Guild, "\") <> 0) Then
                Guild = Replace(Guild, "\", "")
            End If
            If (InStrB(Guild, "/") <> 0) Then
                Guild = Replace(Guild, "/", "")
            End If
            If Not FileExist(App.Path & "\guilds\" & Guild & "-members.mem") Then
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & Guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "INIT", "NroMembers"))
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "Members", "Member" & i)
                    Call WriteConsoleMsg(Userindex, UserName & "<" & Guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGMMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & Message)
            If LenB(Message) <> 0 Then
                Call Statistics.ParseChat(Message)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleShowName(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName
            Call RefreshCharStatus(Userindex)
        End If
    End With
End Sub

Private Sub HandleOnlineRoyalArmy(ByVal Userindex As Integer)
    With UserList(Userindex)
        .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Dim i    As Long
        Dim list As String
        Dim priv As PlayerType
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin
        End If
        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With
    If Len(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Private Sub HandleOnlineChaosLegion(ByVal Userindex As Integer)
    With UserList(Userindex)
        .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Dim i    As Long
        Dim list As String
        Dim priv As PlayerType
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin
        End If
        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With
    If Len(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Private Sub HandleGoNearby(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        Dim tIndex As Integer
        Dim X      As Long
        Dim Y      As Long
        Dim i      As Long
        Dim Found  As Boolean
        tIndex = NameIndex(UserName)
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.Map, X, Y).Userindex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(Userindex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        Found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            If Found Then Exit For
                        Next X
                        If Found Then Exit For
                    Next i
                    If Not Found Then
                        Call WriteConsoleMsg(Userindex, "Todos los lugares estan ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleComment(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim comment As String
        comment = buffer.ReadASCIIString()
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(Userindex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleServerTime(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call LogGM(.Name, "Hora.")
    End With
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

Private Sub HandleWhere(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim miPos    As String
        UserName = buffer.ReadASCIIString()
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If PersonajeExiste(UserName) Then
                    Dim CharPrivs As PlayerType
                    CharPrivs = GetCharPrivs(UserName)
                    If (CharPrivs And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((CharPrivs And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                        miPos = GetUserPos(UserName)
                        Call WriteConsoleMsg(Userindex, "Ubicacion  " & UserName & " (Offline): " & miPos & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(Userindex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(Userindex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(Userindex, "Ubicacion  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call LogGM(.Name, "/Donde " & UserName)
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCreaturesInMap(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1()    As String
        Dim List2()    As String
        Map = .incomingData.ReadInteger()
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        If MapaValido(Map) Then
            For i = 1 To LastNPC
                If Npclist(i).Pos.Map = Map Then
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            Call WriteConsoleMsg(Userindex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(Userindex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next
            End If
            Call WriteConsoleMsg(Userindex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay mas NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(Userindex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & Map)
        End If
    End With
End Sub

Private Sub HandleWarpMeToTarget(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim X As Integer
        Dim Y As Integer
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        X = .flags.TargetX
        Y = .flags.TargetY
        Call FindLegalPos(Userindex, .flags.TargetMap, X, Y)
        Call WarpUserChar(Userindex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)
    End With
End Sub

Private Sub HandleWarpChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim Map      As Integer
        Dim X        As Integer
        Dim Y        As Integer
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        Map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = Userindex
                End If
                If tUser <= 0 Then
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(Userindex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                ElseIf Not ((UserList(tUser).flags.Privilegios And PlayerType.Dios) <> 0 Or (UserList(tUser).flags.Privilegios And PlayerType.Admin) <> 0) Or tUser = Userindex Then
                    If InMapBounds(Map, X, Y) Then
                        Call FindLegalPos(tUser, Map, X, Y)
                        Call WarpUserChar(tUser, Map, X, Y, True, True)
                        If Userindex <> tUser Then
                            Call WriteConsoleMsg(Userindex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Transporto a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleSilence(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(Userindex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias seran ignoradas por el servidor de aqui en mas. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(Userindex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleSOSShowList(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(Userindex)
    End With
End Sub

Private Sub HandlePartyForm(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No perteneces a ningun grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With
End Sub

Private Sub HandleItemUpgrade(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim ItemIndex As Integer
        Call .incomingData.ReadByte
        ItemIndex = .incomingData.ReadInteger()
        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, Userindex) Then Exit Sub
        If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
        Call DoUpgrade(Userindex, ItemIndex)
    End With
End Sub

Private Sub HandleSOSRemove(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        If Not .flags.Privilegios And PlayerType.User Then Call Ayuda.Quitar(UserName)
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleGoToChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim X        As Integer
        Dim Y        As Integer
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(Userindex, UserList(tUser).Pos.Map, X, Y)
                    Call WarpUserChar(Userindex, UserList(tUser).Pos.Map, X, Y, True)
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleInvisible(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call DoAdminInvisible(Userindex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
End Sub

Private Sub HandleGMPanel(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowGMPanelForm(Userindex)
    End With
End Sub

Private Sub HandleRequestUserList(ByVal Userindex As Integer)
    Dim i       As Long
    Dim names() As String
    Dim Count   As Long
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        ReDim names(1 To LastUser) As String
        Count = 1
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).Name
                    Count = Count + 1
                End If
            End If
        Next i
        If Count > 1 Then Call WriteUserNameList(Userindex, names(), Count - 1)
    End With
End Sub

Private Sub HandleWorking(ByVal Userindex As Integer)
    Dim i     As Long
    Dim Users As String
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name
            End If
        Next i
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(Userindex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleHiding(ByVal Userindex As Integer)
    Dim i     As Long
    Dim Users As String
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "
            End If
        Next i
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(Userindex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleJail(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim Reason   As String
        Dim jailTime As Byte
        Dim Count    As Byte
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                If tUser <= 0 Then
                    If (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(Userindex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(Userindex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(Userindex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > (60) Then
                        Call WriteConsoleMsg(Userindex, "No puedes encarcelar por mas de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        If PersonajeExiste(UserName) Then
                            Count = GetUserAmountOfPunishments(UserName)
                            Call SaveUserPunishment(UserName, Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & time)
                        End If
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarcelo a " & UserName)
                    End If
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleKillNPC(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Dim tNPC   As Integer
        Dim auxNPC As npc
        #If ProteccionGM = 1 Then
            Call WriteConsoleMsg(Userindex, "El comando /RMATA se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.Map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(Userindex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        tNPC = .flags.TargetNPC
        If tNPC > 0 Then
            Call WriteConsoleMsg(Userindex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(Userindex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleWarnUser(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim Reason   As String
        Dim Privs    As PlayerType
        Dim Count    As Byte
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(UserName)
                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(Userindex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                    End If
                    If PersonajeExiste(UserName) Then
                        Count = GetUserAmountOfPunishments(UserName)
                        Call SaveUserPunishment(UserName, Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & time)
                        Call WriteConsoleMsg(Userindex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleEditChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 8 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName      As String
        Dim tUser         As Integer
        Dim opcion        As Byte
        Dim Arg1          As String
        Dim Arg2          As String
        Dim valido        As Boolean
        Dim LoopC         As Byte
        Dim CommandString As String
        Dim n             As Byte
        Dim UserCharPath  As String
        Dim Var           As Long
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        If UCase$(UserName) = "YO" Then
            tUser = Userindex
        Else
            tUser = NameIndex(UserName)
        End If
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    valido = tUser = Userindex And (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida)
                
                Case PlayerType.SemiDios
                    ' Los RMs solo se pueden editar su level o vida y el head y body de cualquiera
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = Userindex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                    
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level o vida solo lo puede hacer sobre si mismo
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = Userindex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_CiticensKilled Or opcion = eEditOptions.eo_CriminalsKilled Or opcion = eEditOptions.eo_Class Or opcion = eEditOptions.eo_Skills Or opcion = eEditOptions.eo_addGold
            End Select
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If opcion = eEditOptions.eo_Vida Then
                valido = (tUser = Userindex)
            Else
                valido = True
            End If
        ElseIf .flags.PrivEspecial Then
            valido = (opcion = eEditOptions.eo_CiticensKilled) Or (opcion = eEditOptions.eo_CriminalsKilled)
        End If
        If Database_Enabled And tUser <= 0 Then
            valido = False
            Call WriteConsoleMsg(Userindex, "El usuario esta offline.", FontTypeNames.FONTTYPE_INFO)
        End If
        If valido Then
            UserCharPath = CharPath & UserName & ".chr"
            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(Userindex, "Estas intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intento editar un usuario inexistente.")
            Else
                CommandString = "/MOD "
                Select Case opcion
                    Case eEditOptions.eo_Gold
                        If val(Arg1) <= MAX_ORO_EDIT Then
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).Stats.Gld = val(Arg1)
                                Call WriteUpdateGold(tUser)
                            End If
                        Else
                            Call WriteConsoleMsg(Userindex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        CommandString = CommandString & "ORO "
                        
                    Case eEditOptions.eo_Experience
                        If val(Arg1) > 20000000 Then
                            Arg1 = 20000000
                        End If
                        If tUser <= 0 Then
                            Var = GetVar(UserCharPath, "STATS", "EXP")
                            Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(Arg1))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        End If
                        CommandString = CommandString & "EXP "
                        
                    Case eEditOptions.eo_Body
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        CommandString = CommandString & "HEAD "
                        
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Faccion.CriminalesMatados = Var
                        End If
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Faccion.CiudadanosMatados = Var
                        End If
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(Userindex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        If val(Arg1) >= 25 Then
                            Dim GI As Integer
                            If tUser <= 0 Then
                                GI = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
                            Else
                                GI = UserList(tUser).GuildIndex
                            End If
                            If GI > 0 Then
                                If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                                    Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                                    If tUser > 0 Then Call WriteConsoleMsg(tUser, "Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearas! Por esta razon, hasta tanto no te enlistes en la faccion bajo la cual tu clan esta alineado, estaras excluido del mismo.", FontTypeNames.FONTTYPE_GUILD)
                                End If
                            End If
                        End If
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "STATS", "ELV", val(Arg1))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Stats.ELV = val(Arg1)
                            Call WriteUpdateUserStats(tUser)
                        End If
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(Userindex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).Clase = LoopC
                            End If
                        End If
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NUMSKILLS
                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(Userindex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)
                                Call WriteVar(UserCharPath, "Skills", "EXPSK" & LoopC, 0)
                                If Arg2 < MAXSKILLPOINTS Then
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, ELU_SKILL_INICIAL * 1.05 ^ Arg2)
                                Else
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, 0)
                                End If
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                                Call CheckEluSkill(tUser, LoopC, True)
                            End If
                        End If
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "STATS", "SkillPtsLibres", Arg1)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Stats.SkillPts = val(Arg1)
                        End If
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                    Case eEditOptions.eo_Nobleza
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "REP", "Nobles", Var)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Reputacion.NobleRep = Var
                        End If
                        CommandString = CommandString & "NOB "
                        
                    Case eEditOptions.eo_Asesino
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "REP", "Asesino", Var)
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Reputacion.AsesinoRep = Var
                        End If
                        CommandString = CommandString & "ASE "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0)
                        Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex)
                        If Sex <> 0 Then
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).Genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(Userindex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                raza = eRaza.Humano

                            Case "ELFO"
                                raza = eRaza.Elfo

                            Case "DROW"
                                raza = eRaza.Drow

                            Case "ENANO"
                                raza = eRaza.Enano

                            Case "GNOMO"
                                raza = eRaza.Gnomo

                            Case Else
                                raza = 0
                        End Select
                        If raza = 0 Then
                            Call WriteConsoleMsg(Userindex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Raza", raza)
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza
                            End If
                        End If
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                        Dim bankGold As Long
                        If Abs(Arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(Userindex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                bankGold = GetVar(UserCharPath, "STATS", "BANCO")
                                Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(Arg1) <= 0, 0, bankGold + val(Arg1)))
                                Call WriteConsoleMsg(Userindex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                            Else
                                UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + val(Arg1))
                                Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)
                            End If
                        End If
                        CommandString = CommandString & "AGREGAR "
                    
                    Case eEditOptions.eo_Vida
                        If val(Arg1) > MAX_VIDA_EDIT Then
                            Arg1 = CStr(MAX_VIDA_EDIT)
                            Call WriteConsoleMsg(Userindex, "No puedes tener vida superior a " & MAX_VIDA_EDIT & ".", FONTTYPE_INFO)
                        End If
                        UserList(tUser).Stats.MaxHp = val(Arg1)
                        UserList(tUser).Stats.MinHp = val(Arg1)
                        Call WriteUpdateUserStats(tUser)
                        CommandString = CommandString & "VIDA "
                        
                    Case eEditOptions.eo_Poss
                        Dim Map As Integer
                        Dim X   As Integer
                        Dim Y   As Integer
                        Map = val(ReadField(1, Arg1, 45))
                        X = val(ReadField(2, Arg1, 45))
                        Y = val(ReadField(3, Arg1, 45))
                        If InMapBounds(Map, X, Y) Then
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "POSITION", Map & "-" & X & "-" & Y)
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WarpUserChar(tUser, Map, X, Y, True, True)
                                Call WriteConsoleMsg(Userindex, "Usuario teletransportado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(Userindex, "Posicion invalida", FONTTYPE_INFO)
                        End If
                        CommandString = CommandString & "POSS "
                    Case Else
                        Call WriteConsoleMsg(Userindex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                End Select
                CommandString = CommandString & Arg1 & " " & Arg2
                If Userindex <> tUser Then
                    Call LogGM(.Name, CommandString & " " & UserName)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRequestCharInfo(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim TargetName  As String
        Dim targetIndex As Integer
        TargetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(TargetName)
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            If targetIndex <= 0 Then
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline, buscando...", FontTypeNames.FONTTYPE_INFO)
                    If Not Database_Enabled Then
                        Call SendUserStatsTxtCharfile(Userindex, TargetName)
                    Else
                        Call SendUserStatsTxtDatabase(Userindex, TargetName)
                    End If
                End If
            Else
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(Userindex, targetIndex)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRequestCharStats(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName         As String
        Dim tUser            As Integer
        Dim UserIsAdmin      As Boolean
        Dim OtherUserIsAdmin As Boolean
        UserName = buffer.ReadASCIIString()
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And ((.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin) Then
            Call LogGM(.Name, "/STAT " & UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_INFO)
                    If Not Database_Enabled Then
                        Call SendUserMiniStatsTxtFromCharfile(Userindex, UserName)
                    Else
                        Call SendUserMiniStatsTxtFromDatabase(Userindex, UserName)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserMiniStatsTxt(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRequestCharGold(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName         As String
        Dim tUser            As Integer
        Dim UserIsAdmin      As Boolean
        Dim OtherUserIsAdmin As Boolean
        UserName = buffer.ReadASCIIString()
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        If (.flags.Privilegios And PlayerType.SemiDios) Or UserIsAdmin Then
            Call LogGM(.Name, "/BAL " & UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)
                    If Not Database_Enabled Then
                        Call SendUserOROTxtFromCharfile(Userindex, UserName)
                    Else
                        Call SendUserOROTxtFromDatabase(Userindex, UserName)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(Userindex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRequestCharInventory(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName         As String
        Dim tUser            As Integer
        Dim UserIsAdmin      As Boolean
        Dim OtherUserIsAdmin As Boolean
        UserName = buffer.ReadASCIIString()
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline. Buscando...", FontTypeNames.FONTTYPE_TALK)
                    If Not Database_Enabled Then
                        Call SendUserInvTxtFromCharfile(Userindex, UserName)
                    Else
                        Call SendUserInvTxtFromDatabase(Userindex, UserName)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserInvTxt(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRequestCharBank(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName         As String
        Dim tUser            As Integer
        Dim UserIsAdmin      As Boolean
        Dim OtherUserIsAdmin As Boolean
        UserName = buffer.ReadASCIIString()
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.Name, "/BOV " & UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)
                    If Not Database_Enabled Then
                        Call SendUserBovedaTxtFromCharfile(Userindex, UserName)
                    Else
                        Call SendUserBovedaTxtFromDatabase(Userindex, UserName)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserBovedaTxt(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRequestCharSkills(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Long
        Dim Message  As String
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")
                End If
                For LoopC = 1 To NUMSKILLS
                    Message = Message & GetUserSkills(UserName)
                Next LoopC
                Call WriteConsoleMsg(Userindex, Message & "CHAR> Libres: " & GetUserFreeSkills(UserName), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(Userindex, tUser)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleReviveChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Byte
        UserName = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = Userindex
            End If
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)
                        End If
                        If .flags.Traveling = 1 Then
                            .flags.Traveling = 0
                            .Counters.goHome = 0
                            Call WriteMultiMessage(tUser, eMessages.CancelHome)
                        End If
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    .Stats.MinHp = .Stats.MaxHp
                    If .flags.Traveling = 1 Then
                        .Counters.goHome = 0
                        .flags.Traveling = 0
                        Call WriteMultiMessage(tUser, eMessages.CancelHome)
                    End If
                End With
                Call WriteUpdateHP(tUser)
                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleOnlineGM(ByVal Userindex As Integer)
    Dim i    As Long
    Dim list As String
    Dim priv As PlayerType
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).Name & ", "
            End If
        Next i
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(Userindex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleOnlineMap(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Map As Integer
        Map = .incomingData.ReadInteger
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        Dim LoopC As Long
        Dim list  As String
        Dim priv  As PlayerType
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        Call WriteConsoleMsg(Userindex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.Name, "/ONLINEMAP " & Map)
    End With
End Sub

Private Sub HandleForgive(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(Userindex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleKick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim rank     As Integer
        Dim IsAdmin  As Boolean
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(Userindex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(Userindex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echo a " & UserName)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleExecute(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(Userindex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                End If
            Else
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(Userindex, "No esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleBanChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim Reason   As String
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(Userindex, UserName, Reason)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleUnbanChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName  As String
        Dim cantPenas As Byte
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If Not PersonajeExiste(UserName) Then
                Call WriteConsoleMsg(Userindex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else
                If BANCheck(UserName) Then
                    Call UnBan(UserName)
                    cantPenas = GetUserAmountOfPunishments(UserName)
                    Call SaveUserPunishment(UserName, cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(Userindex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " no esta baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleNPCFollow(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        #If ProteccionGM = 1 Then
            Call WriteConsoleMsg(Userindex, "El comando /SEGUIR se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

Private Sub HandleSummonChar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim X        As Integer
        Dim Y        As Integer
        UserName = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If EsDios(UserName) Or EsAdmin(UserName) Then
                    Call WriteConsoleMsg(Userindex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleSpawnListRequest(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        Call EnviarSpawnList(Userindex)
    End With
End Sub

Private Sub HandleSpawnCreature(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
End Sub

Private Sub HandleResetNPCInventory(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

Private Sub HandleServerMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(Userindex).Name & "> " & Message, FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleMapMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Dim Mapa As Integer
                Mapa = .Pos.Map
                Call LogGM(.Name, "Mensaje a mapa " & Mapa & ":" & Message)
                Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleNickToIP(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim priv     As PlayerType
        Dim IsAdmin  As Boolean
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)
            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0
            If IsAdmin Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(Userindex, "El ip de " & UserName & " es " & UserList(tUser).IP, FontTypeNames.FONTTYPE_INFO)
                    Dim IP    As String
                    Dim lista As String
                    Dim LoopC As Long
                    IP = UserList(tUser).IP
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).IP = IP Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(Userindex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(Userindex, "No hay ningUn personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleIPToNick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim IP    As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv  As PlayerType
        IP = .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & IP)
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If
        For LoopC = 1 To LastUser
            If UserList(LoopC).IP = IP Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(Userindex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleGuildOnlineMembers(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim GuildName As String
        Dim tGuild    As Integer
        GuildName = buffer.ReadASCIIString()
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            If tGuild > 0 Then
                Call WriteConsoleMsg(Userindex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(Userindex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleTeleportCreate(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Mapa  As Integer
        Dim X     As Byte
        Dim Y     As Byte
        Dim Radio As Byte
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Radio = .incomingData.ReadByte()
        Radio = MinimoInt(Radio, 6)
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then Exit Sub
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(Userindex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim ET As obj
        ET.Amount = 1
        ET.ObjIndex = TELEP_OBJ_INDEX + Radio
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

Private Sub HandleTeleportDestroy(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim Mapa As Integer
        Dim X    As Byte
        Dim Y    As Byte
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        With MapData(Mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                                Call LogGM(UserList(Userindex).Name, "/DT: " & Mapa & "," & X & "," & Y)
                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                End If
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

Private Sub HandleExitDestroy(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim Mapa As Integer
        Dim X    As Byte
        Dim Y    As Byte
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        With MapData(Mapa, X, Y)
            If .TileExit.Map = 0 Then Exit Sub
            If .ObjInfo.ObjIndex > 0 Then
                If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then Exit Sub
            End If
            Call LogGM(UserList(Userindex).Name, "/DE: " & Mapa & "," & X & "," & Y)
            .TileExit.Map = 0
            .TileExit.X = 0
            .TileExit.Y = 0
        End With
    End With
End Sub

Private Sub HandleRainToggle(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

Private Sub HandleEnableDenounces(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If Not EsGm(Userindex) Then Exit Sub
        Dim Activado As Boolean
        Dim Msg      As String
        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado
        Msg = "Denuncias por consola " & IIf(Activado, "ativadas", "desactivadas") & "."
        Call LogGM(.Name, Msg)
        Call WriteConsoleMsg(Userindex, Msg, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleShowDenouncesList(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowDenounces(Userindex)
    End With
End Sub

Private Sub HandleSetCharDescription(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim tUser As Integer
        Dim Desc  As String
        Desc = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(Userindex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HanldeForceMP3ToMap(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Mp3Id As Byte
        Dim Mapa   As Integer
        Mp3Id = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map
            End If
            If Mp3Id = 0 Then
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMp3(MapInfo(.Pos.Map).Music))
            Else
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMp3(Mp3Id))
            End If
        End If
    End With
End Sub

Private Sub HanldeForceMIDIToMap(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim midiID As Byte
        Dim Mapa   As Integer
        midiID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map
            End If
            If midiID = 0 Then
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).Music))
            Else
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

Private Sub HandleForceWAVEToMap(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim waveID As Byte
        Dim Mapa   As Integer
        Dim X      As Byte
        Dim Y      As Byte
        waveID = .incomingData.ReadByte()
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            If Not InMapBounds(Mapa, X, Y) Then
                Mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y
            End If
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

Private Sub HandleRoyalArmyMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJERCITO REAL> " & Message, FontTypeNames.FONTTYPE_TALK))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleChaosLegionMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, FontTypeNames.FONTTYPE_TALK))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCitizenMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Message, FontTypeNames.FONTTYPE_TALK))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCriminalMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Message, FontTypeNames.FONTTYPE_TALK))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleTalkAsNPC(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(Userindex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleDestroyAllItemsInArea(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Dim X       As Long
        Dim Y       As Long
        Dim bIsExit As Boolean
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        Call LogGM(UserList(Userindex).Name, "/MASSDEST")
    End With
End Sub

Private Sub HandleAcceptRoyalCouncilMember(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Byte
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleAcceptChaosCouncilMember(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim LoopC    As Byte
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleItemsInTheFloor(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Dim tObj  As Integer
        Dim lista As String
        Dim X     As Long
        Dim Y     As Long
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(Userindex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

Private Sub HandleMakeDumb(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleMakeDumbNoMore(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleDumpIPTables(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Call SecurityIp.DumpTables
    End With
End Sub

Private Sub HandleCouncilKick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If PersonajeExiste(UserName) Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call KickUserCouncils(UserName)
                Else
                    Call WriteConsoleMsg(Userindex, "No se encuentra el charfile " & CharPath & UserName, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleSetTrigger(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim tTrigger As Byte
        Dim tLog     As String
        tTrigger = .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(Userindex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleAskTrigger(ByVal Userindex As Integer)
    Dim tTrigger As Byte
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        Call WriteConsoleMsg(Userindex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleBannedIPList(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Dim lista As String
        Dim LoopC As Long
        Call LogGM(.Name, "/BANIPLIST")
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleBannedIPReload(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

Private Sub HandleGuildBan(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim GuildName   As String
        Dim cantMembers As Integer
        Dim LoopC       As Long
        Dim member      As String
        Dim Count       As Byte
        Dim tIndex      As Integer
        Dim tFile       As String
        GuildName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneo al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_GUILD))
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If
                    Call SaveBan(member, "BAN AL CLAN: " & GuildName, LCase$(.Name))
                Next LoopC
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleBanIP(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim bannedIP As String
        Dim tUser    As Integer
        Dim Reason   As String
        Dim i        As Long
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            If tUser > 0 Then bannedIP = UserList(tUser).IP
        End If
        Reason = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(Userindex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneo la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).IP = bannedIP Then
                                Call BanCharacter(Userindex, UserList(i).Name, "IP POR " & Reason)
                            End If
                        End If
                    Next i
                End If
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El personaje no esta online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleUnbanIP(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim bannedIP As String
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleCreateItem(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim tObj    As Integer: tObj = .incomingData.ReadInteger()
        Dim Cuantos As Integer: Cuantos = .incomingData.ReadInteger()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        #If ProteccionGM = 1 Then
            Call WriteConsoleMsg(Userindex, "El comando /CI se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        If Cuantos > 10000 Then Call WriteConsoleMsg(Userindex, "Estas tratando de crear demasiado, como mucho podes crear 10.000 unidades.", FontTypeNames.FONTTYPE_TALK): Exit Sub
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        Dim Objeto As obj
        With Objeto
            .Amount = Cuantos
            .ObjIndex = tObj
        End With
        If ObjData(tObj).Agarrable = 0 Then
            If MeterItemEnInventario(Userindex, Objeto) Then
                Call WriteConsoleMsg(Userindex, "Has creado " & Objeto.Amount & " unidades de " & ObjData(tObj).Name & ".", FontTypeNames.FONTTYPE_INFO)
            Else
                Call TirarItemAlPiso(.Pos, Objeto)
                Call WriteConsoleMsg(Userindex, "No tenes espacio en tu inventario para crear el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            Call TirarItemAlPiso(.Pos, Objeto)
            Call WriteConsoleMsg(Userindex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        End If
        Call LogGM(.Name, "/CI: " & tObj & " - [Nombre del Objeto: " & ObjData(tObj).Name & "] - [Cantidad : " & Cuantos & "]")
    End With
ErrorHandler:
    If Err.Number <> 0 Then
        Call LogError("Error en HandleCreateItem " & Err.Number & " " & Err.description)
    End If
End Sub
Private Sub HandleDestroyItems(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Dim Mapa As Integer
        Dim X    As Byte
        Dim Y    As Byte
        Mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        Dim ObjIndex As Integer
        ObjIndex = MapData(Mapa, X, Y).ObjInfo.ObjIndex
        If ObjIndex = 0 Then Exit Sub
        Call LogGM(.Name, "/DEST " & ObjIndex & " en mapa " & Mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(Mapa, X, Y).ObjInfo.Amount)
        If ObjData(ObjIndex).OBJType = eOBJType.otTeleport And MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(Userindex, "No puede destruir teleports asi. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call EraseObj(10000, Mapa, X, Y)
    End With
End Sub

Private Sub HandleChaosLegionKick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                If PersonajeExiste(UserName) Then
                    Call KickUserChaosLegion(UserName, .Name)
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleRoyalArmyKick(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "ECHO DE LA REAL A: " & UserName)
            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                If PersonajeExiste(UserName) Then
                    Call KickUserRoyalArmy(UserName, .Name)
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleForceMP3All(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim Mp3Id As Byte
        Mp3Id = .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica MP3: " & Mp3Id, FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMp3(Mp3Id))
    End With
End Sub

Private Sub HandleForceMIDIAll(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica MIDI: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

Private Sub HandleForceWAVEAll(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

Private Sub HandleRemovePunishment(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName   As String
        Dim punishment As Byte
        Dim NewText    As String
        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")
                End If
                If PersonajeExiste(UserName) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & " de " & UserName & " y la cambio por: " & NewText)
                    Call AlterUserPunishment(UserName, punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)
                    Call WriteConsoleMsg(Userindex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleTileBlockedToggle(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        Call LogGM(.Name, "/BLOQ")
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

Private Sub HandleKillNPCNoRespawn(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        #If ProteccionGM = 1 Then
            Call WriteConsoleMsg(Userindex, "El comando /MATA se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

Private Sub HandleKillAllNearbyNPCs(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        #If ProteccionGM = 1 Then
            Call WriteConsoleMsg(Userindex, "El comando /MASSKILL se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        Dim X As Long
        Dim Y As Long
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
End Sub

Private Sub HandleLastIP(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName   As String
        Dim lista      As String
        Dim LoopC      As Byte
        Dim priv       As Integer
        Dim validCheck As Boolean
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)
                If PersonajeExiste(UserName) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conecto son:" & vbCrLf & GetUserLastIps(UserName)
                    Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, UserName & " es de mayor jerarquia que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleChatColor(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim color As Long
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

Public Sub HandleIgnored(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

Public Sub HandleCheckSlot(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName         As String
        Dim Slot             As Byte
        Dim tIndex           As Integer
        Dim UserIsAdmin      As Boolean
        Dim OtherUserIsAdmin As Boolean
        UserName = buffer.ReadASCIIString()
        Slot = buffer.ReadByte()
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.Name, .Name & " Checkeo el slot " & Slot & " de " & UserName)
            tIndex = NameIndex(UserName)  'Que user index?
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            If tIndex > 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                        If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                            Call WriteConsoleMsg(Userindex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, "No hay ningUn objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "Slot Invalido.", FontTypeNames.FONTTYPE_TALK)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleResetAutoUpdate(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        Call WriteConsoleMsg(Userindex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleRestart(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        Call LogGM(.Name, .Name & " reinicio el mundo.")
        Call ReiniciarServidor(True)
    End With
End Sub

Public Sub HandleReloadObjects(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        Call LoadOBJData
    End With
End Sub

Public Sub HandleReloadSpells(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        Call CargarHechizos
    End With
End Sub

Public Sub HandleReloadServerIni(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        Call LoadSini
        Call WriteConsoleMsg(Userindex, "Server.ini actualizado correctamente", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleReloadNPCs(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
        Call CargaNpcsDat
        Call WriteConsoleMsg(Userindex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleKickAllChars(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        Call EcharPjsNoPrivilegiados
    End With
End Sub

Public Sub HandleNight(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        DeNoche = Not DeNoche
        Dim i As Long
        For i = 1 To NumUsers
            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)
            End If
        Next i
    End With
End Sub

Public Sub HandleShowServerForm(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

Public Sub HandleCleanSOS(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        Call Ayuda.Reset
    End With
End Sub

Public Sub HandleSaveChars(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
End Sub

Public Sub HandleChangeMapInfoBackup(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim doTheBackUp As Boolean
        doTheBackUp = .incomingData.ReadBoolean()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre el BackUp.")
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0
        End If
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleChangeMapInfoPK(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim isMapPk As Boolean
        isMapPk = .incomingData.ReadBoolean()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es PK el mapa.")
        MapInfo(.Pos.Map).Pk = isMapPk
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))
        Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleChangeMapInfoRestricted(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim tStr As String
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        tStr = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es restringido el mapa.")
                MapInfo(UserList(Userindex).Pos.Map).Restringir = RestrictStringToByte(tStr)
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleChangeMapInfoNoMagic(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim nomagic As Boolean
    With UserList(Userindex)
        Call .incomingData.ReadByte
        nomagic = .incomingData.ReadBoolean
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la magia el mapa.")
            MapInfo(UserList(Userindex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub HandleChangeMapInfoNoInvi(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim noinvi As Boolean
    With UserList(Userindex)
        Call .incomingData.ReadByte
        noinvi = .incomingData.ReadBoolean()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(Userindex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
Public Sub HandleChangeMapInfoNoResu(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim noresu As Boolean
    With UserList(Userindex)
        Call .incomingData.ReadByte
        noresu = .incomingData.ReadBoolean()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar el resucitar en el mapa.")
            MapInfo(UserList(Userindex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub HandleChangeMapInfoLand(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim tStr As String
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        tStr = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion del terreno del mapa.")
                MapInfo(UserList(Userindex).Pos.Map).Terreno = TerrainStringToByte(tStr)
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.Map).Terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "Igualmente, el Unico Util es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleChangeMapInfoZone(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim tStr As String
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        tStr = buffer.ReadASCIIString()
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion de la zona del mapa.")
                MapInfo(UserList(Userindex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, " 'DUNGEON', NO se sentira el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, " 'NIEVE', Les agarra frio y saca salud hasta morir sin ropa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub
            
Public Sub HandleChangeMapInfoStealNpc(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim RoboNpc As Byte
    With UserList(Userindex)
        Call .incomingData.ReadByte
        RoboNpc = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido robar npcs en el mapa.")
            MapInfo(UserList(Userindex).Pos.Map).RoboNpcsPermitido = RoboNpc
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapInfo(.Pos.Map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
Public Sub HandleChangeMapInfoNoOcultar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim NoOcultar As Byte
    Dim Mapa      As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        NoOcultar = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Mapa = .Pos.Map
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido ocultarse en el mapa " & Mapa & ".")
            MapInfo(Mapa).OcultarSinEfecto = NoOcultar
            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(Userindex, "Mapa " & Mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
           
Public Sub HandleChangeMapInfoNoInvocar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim NoInvocar As Byte
    Dim Mapa      As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        NoInvocar = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Mapa = .Pos.Map
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido invocar en el mapa " & Mapa & ".")
            MapInfo(Mapa).InvocarSinEfecto = NoInvocar
            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(Userindex, "Mapa " & Mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub HandleSaveMap(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        Call WriteConsoleMsg(Userindex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub HandleShowGuildMessages(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Guild As String
        Guild = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(Userindex, Guild)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleDoBackUp(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, .Name & " ha hecho un backup.")
        Call ES.DoBackUp
    End With
End Sub

Public Sub HandleToggleCentinelActivated(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If Not EsAdmin(.Name) Or Not EsDios(.Name) Then Exit Sub
        Call modCentinela.CambiarEstado(Userindex)
    End With
End Sub

Public Sub HandleAlterName(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName     As String
        Dim newName      As String
        Dim changeNameUI As Integer
        Dim GuildIndex   As Integer
        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(Userindex, "El Pj esta online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not PersonajeExiste(UserName) Then
                        Call WriteConsoleMsg(Userindex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If GetUserGuildIndex(UserName) > 0 Then
                            Call WriteConsoleMsg(Userindex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not PersonajeExiste(newName) Then
                                Call CopyUser(UserName, newName)
                                If Not Database_Enabled Then
                                    Call SaveBan(UserName, "BAN POR Cambio de nick a " & UCase$(newName), .Name)
                                End If
                                Call WriteConsoleMsg(Userindex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(Userindex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub
Public Sub HandleAlterMail(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim newMail  As String
        UserName = buffer.ReadASCIIString()
        newMail = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not PersonajeExiste(UserName) Then
                    Call WriteConsoleMsg(Userindex, "No existe el charfile de" & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SaveUserEmail(UserName, newMail)
                    Call WriteConsoleMsg(Userindex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If
                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleAlterPassword(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        Dim Salt     As String
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contrasena de " & UserName)
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not PersonajeExiste(UserName) Or Not PersonajeExiste(copyFrom) Then
                    Call WriteConsoleMsg(Userindex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetUserPassword(copyFrom)
                    Salt = GetUserSalt(copyFrom)
                    Call StorePasswordSalt(UserName, Password, Salt)
                    Call WriteConsoleMsg(Userindex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleCreateNPC(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim NpcIndex As Integer: NpcIndex = .incomingData.ReadInteger()
        Dim Respawn As Boolean: Respawn = .incomingData.ReadBoolean()
        If Not EsGm(Userindex) Then Exit Sub
        If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
            Call WriteConsoleMsg(Userindex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CREARPRETORIANOS MAPA X Y.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        If NpcIndex <> 0 Then
            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, Respawn)
            Call LogGM(.Name, "Invoco " & IIf(Respawn, "con respawn", vbNullString) & " a " & Npclist(NpcIndex).Name & " [Indice: " & NpcIndex & "] en el mapa " & .Pos.Map)
        End If
    End With
End Sub

Public Sub HandleImperialArmour(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim index    As Byte
        Dim ObjIndex As Integer
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Select Case index
            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex
        End Select
    End With
End Sub
Public Sub HandleChaosArmour(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim index    As Byte
        Dim ObjIndex As Integer
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Select Case index
            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex
        End Select
    End With
End Sub

Public Sub HandleNavigateToggle(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        Call WriteNavigateToggle(Userindex)
    End With
End Sub

Public Sub HandleServerOpenToUsersToggle(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(Userindex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            frmMain.chkServerHabilitado.Value = vbUnchecked
        Else
            Call WriteConsoleMsg(Userindex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
            frmMain.chkServerHabilitado.Value = vbChecked
        End If
    End With
End Sub

Public Sub HandleTurnOffServer(ByVal Userindex As Integer)
    Dim handle As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "
        Close #handle
        Unload frmMain
    End With
End Sub

Public Sub HandleTurnCriminal(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)
            tUser = NameIndex(UserName)
            If tUser > 0 Then Call VolverCriminal(tUser)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleResetFactions(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser    As Integer
        Dim Char     As String
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                If PersonajeExiste(UserName) Then
                    Call ResetUserFacciones(UserName)
                Else
                    Call WriteConsoleMsg(Userindex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleRemoveCharFromGuild(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName   As String
        Dim GuildIndex As Integer
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)
            GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(Userindex, "No pertenece a ningUn clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleRequestCharMail(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim mail     As String
        UserName = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If PersonajeExiste(UserName) Then
                mail = GetUserEmail(UserName)
                Call WriteConsoleMsg(Userindex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleSystemMessage(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Message As String
        Message = buffer.ReadASCIIString()
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & Message)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleSetMOTD(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim newMOTD           As String
        Dim auxiliaryString() As String
        Dim LoopC             As Long
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            MaxLines = UBound(auxiliaryString()) + 1
            ReDim MOTD(1 To MaxLines)
            Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            For LoopC = 1 To MaxLines
                Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            Call WriteConsoleMsg(Userindex, "Se ha cambiado el MOTD con exito.", FontTypeNames.FONTTYPE_INFO)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleChangeMOTD(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        Dim auxiliaryString As String
        Dim LoopC           As Long
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        Call WriteShowMOTDEditionForm(Userindex, auxiliaryString)
    End With
End Sub

Public Sub HandlePing(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Call WritePong(Userindex)
    End With
End Sub

Public Sub HandleSetIniVar(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String
        sLlave = buffer.ReadASCIIString()
        sClave = buffer.ReadASCIIString()
        sValor = buffer.ReadASCIIString()
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(Userindex, "No puedes modificar esa informacion desde aqui!", FontTypeNames.FONTTYPE_INFO)
            Else
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modifico en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(Userindex, "Modifico " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleCreatePretorianClan(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim Map   As Integer
    Dim X     As Byte
    Dim Y     As Byte
    Dim index As Long
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Map = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(Userindex, "Posicion invalida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Map = MAPA_PRETORIANO Then
            index = ePretorianType.Default
        Else
            index = ePretorianType.Custom
        End If
        If Not ClanPretoriano(index).Active Then
            If Not ClanPretoriano(index).SpawnClan(Map, X, Y, index) Then
                Call WriteConsoleMsg(Userindex, "La posicion no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub HandleDeletePretorianClan(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim Map   As Integer
    Dim index As Long
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Map = .incomingData.ReadInteger()
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(Userindex, "Mapa invalido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        For index = 1 To UBound(ClanPretoriano)
            If ClanPretoriano(index).ClanMap = Map Then
                ClanPretoriano(index).DeleteClan
                Exit For
            End If
        Next index
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub WriteLoggedMessage(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.Logged)
        #If AntiExternos Then
            UserList(Userindex).Redundance = RandomNumber(15, 250)
            Call UserList(Userindex).outgoingData.WriteByte(UserList(Userindex).Redundance)
        #End If
        Call .outgoingData.WriteByte(.Clase)
        Call .outgoingData.WriteLong(IntervaloInvisible)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteRemoveAllDialogs(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteRemoveCharDialog(ByVal Userindex As Integer, ByVal CharIndex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteNavigateToggle(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteDisconnect(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Disconnect)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserOfferConfirm(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCommerceEnd(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBankEnd(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCommerceInit(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBankInit(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(Userindex).outgoingData.WriteLong(UserList(Userindex).Stats.Banco)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserCommerceInit(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(Userindex).outgoingData.WriteASCIIString(UserList(Userindex).ComUsu.DestNick)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserCommerceEnd(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateSta(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateMana(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateHP(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(Userindex).Stats.MinHp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateGold(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(Userindex).Stats.Gld)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateBankGold(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(Userindex).Stats.Banco)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateExp(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(Userindex).Stats.Exp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateStrenghtAndDexterity(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteLong(UserList(Userindex).flags.DuracionEfecto)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateDexterity(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteLong(UserList(Userindex).flags.DuracionEfecto)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateStrenght(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteLong(UserList(Userindex).flags.DuracionEfecto)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeMap(ByVal Userindex As Integer, ByVal Map As Integer, ByVal version As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteASCIIString(MapInfo(Map).Name)
        Call .WriteASCIIString(MapInfo(Map).Zona)
        Call .WriteInteger(version)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePosUpdate(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(Userindex).Pos.X)
        Call .WriteByte(UserList(Userindex).Pos.Y)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChatOverHead(ByVal Userindex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteConsoleMsg(ByVal Userindex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteRenderMsg(ByVal Userindex As Integer, ByVal Chat As String, ByVal FontIndex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareRenderConsoleMsg(Chat, FontIndex))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCommerceChat(ByVal Userindex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
            
Public Sub WriteGuildChat(ByVal Userindex As Integer, ByVal Chat As String)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowMessageBox(ByVal Userindex As Integer, ByVal Message As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Message)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserIndexInServer(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(Userindex)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserCharIndexInServer(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCharacterCreate(ByVal Userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal X As Byte, _
                                ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, ByVal Privileges As Byte)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, Name, NickColor, Privileges))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCharacterRemove(ByVal Userindex As Integer, ByVal CharIndex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCharacterMove(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal Userindex, ByVal Direccion As eHeading)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCharacterChange(ByVal Userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteObjectCreate(ByVal Userindex As Integer, ByVal GrhIndex As Long, ByVal X As Byte, ByVal Y As Byte)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteObjectDelete(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBlockPosition(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePlayMp3(ByVal Userindex As Integer, ByVal mp3 As Integer, Optional ByVal loops As Integer = -1)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMp3(mp3, loops))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePlayMidi(ByVal Userindex As Integer, ByVal midi As Integer, Optional ByVal loops As Integer = -1)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePlayWave(ByVal Userindex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteGuildList(ByVal Userindex As Integer, ByRef guildList() As String)
    On Error GoTo ErrorHandler
    Dim Tmp As String
    Dim i   As Long
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteAreaChanged(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(Userindex).Pos.X)
        Call .WriteByte(UserList(Userindex).Pos.Y)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePauseToggle(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteRainToggle(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCreateFX(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateUserStats(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(Userindex).Stats.MaxHp)
        Call .WriteInteger(UserList(Userindex).Stats.MinHp)
        Call .WriteInteger(UserList(Userindex).Stats.MaxMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MaxSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
        Call .WriteLong(UserList(Userindex).Stats.Gld)
        Call .WriteByte(UserList(Userindex).Stats.ELV)
        Call .WriteLong(UserList(Userindex).Stats.ELU)
        Call .WriteLong(UserList(Userindex).Stats.Exp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeInventorySlot(ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        Dim ObjIndex As Integer
        Dim obData   As ObjData
        ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(Userindex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(Userindex).Invent.Object(Slot).Equipped)
        Call .WriteLong(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteBoolean(ItemIncompatibleConUser(Userindex, ObjIndex))
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteAddSlots(ByVal Userindex As Integer, ByVal Mochila As eMochilas)
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AddSlots)
        Call .WriteByte(Mochila)
    End With
End Sub

Public Sub WriteChangeBankSlot(ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        Dim ObjIndex As Integer
        Dim obData   As ObjData
        ObjIndex = UserList(Userindex).BancoInvent.Object(Slot).ObjIndex
        Call .WriteInteger(ObjIndex)
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(Userindex).BancoInvent.Object(Slot).Amount)
        Call .WriteLong(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteLong(obData.Valor)
        Call .WriteBoolean(ItemIncompatibleConUser(Userindex, ObjIndex))
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeSpellSlot(ByVal Userindex As Integer, ByVal Slot As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(Userindex).Stats.UserHechizos(Slot))
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteAttributes(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithWeapons(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        For i = 1 To UBound(ArmasHerrero())
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).Clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        Call .WriteInteger(Count)
        For i = 1 To Count
            obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.Name)
            Call .WriteLong(obj.GrhIndex)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(obj.Upgrade)
        Next i
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithArmors(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        For i = 1 To UBound(ArmadurasHerrero())
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).Clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        Call .WriteInteger(Count)
        For i = 1 To Count
            obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.Name)
            Call .WriteLong(obj.GrhIndex)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
            Call .WriteInteger(obj.Upgrade)
        Next i
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteInitCarpenting(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.InitCarpenting)
        For i = 1 To UBound(ObjCarpintero())
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(Userindex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        Call .WriteInteger(Count)
        For i = 1 To Count
            obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(obj.Name)
            Call .WriteLong(obj.GrhIndex)
            Call .WriteInteger(obj.Madera)
            Call .WriteInteger(obj.MaderaElfica)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            Call .WriteInteger(obj.Upgrade)
        Next i
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteInitCraftsman(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i              As Long
    Dim j              As Long
    Dim obj            As ObjData
    Dim ObjRequired    As ObjData
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.InitCraftman)
        Call .WriteLong(ArtesaniaCosto)
        Call .WriteInteger(UBound(ObjArtesano))
        For i = 1 To UBound(ObjArtesano)
            obj = ObjData(ObjArtesano(i))
            Call .WriteASCIIString(obj.Name)
            Call .WriteLong(obj.GrhIndex)
            Call .WriteInteger(ObjArtesano(i))
            Call .WriteByte(UBound(obj.ItemCrafteo))
            For j = 1 To UBound(obj.ItemCrafteo)
                ObjRequired = ObjData(obj.ItemCrafteo(j).ObjIndex)
                Call .WriteASCIIString(ObjRequired.Name)
                Call .WriteLong(ObjRequired.GrhIndex)
                Call .WriteInteger(obj.ItemCrafteo(j).ObjIndex)
                Call .WriteInteger(obj.ItemCrafteo(j).Amount)
            Next j
        Next i
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteRestOK(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteErrorMsg(ByVal Userindex As Integer, ByVal Message As String)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(Message))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBlind(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteDumb(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowSignal(ByVal Userindex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteLong(ObjData(ObjIndex).GrhSecundario)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeNPCInventorySlot(ByVal Userindex As Integer, ByVal Slot As Byte, ByRef obj As obj, ByVal price As Single)
    On Error GoTo ErrorHandler
    Dim ObjInfo As ObjData
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)
    End If
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.Name)
        Call .WriteInteger(obj.Amount)
        Call .WriteSingle(price)
        Call .WriteLong(ObjInfo.GrhIndex)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.MaxDef)
        Call .WriteInteger(ObjInfo.MinDef)
        Call .WriteBoolean(ItemIncompatibleConUser(Userindex, obj.ObjIndex))
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUpdateHungerAndThirst(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(Userindex).Stats.MaxAGU)
        Call .WriteByte(UserList(Userindex).Stats.MinAGU)
        Call .WriteByte(UserList(Userindex).Stats.MaxHam)
        Call .WriteByte(UserList(Userindex).Stats.MinHam)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteFame(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)
        Call .WriteLong(UserList(Userindex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(Userindex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(Userindex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(Userindex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(Userindex).Reputacion.NobleRep)
        Call .WriteLong(UserList(Userindex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(Userindex).Reputacion.Promedio)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteMiniStats(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        Call .WriteLong(UserList(Userindex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(Userindex).Faccion.CriminalesMatados)
        Call .WriteLong(UserList(Userindex).Stats.UsuariosMatados)
        Call .WriteInteger(UserList(Userindex).Stats.NPCsMuertos)
        Call .WriteByte(UserList(Userindex).Clase)
        Call .WriteLong(UserList(Userindex).Counters.Pena)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteLevelUp(ByVal Userindex As Integer, ByVal skillPoints As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteAddForumMsg(ByVal Userindex As Integer, ByVal ForumType As eForumType, ByRef Title As String, ByRef Author As String, ByRef Message As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(Message)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowForumForm(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim Visibilidad   As Byte
    Dim CanMakeSticky As Byte
    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        If esCaos(Userindex) Or EsGm(Userindex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
        End If
        If esArmada(Userindex) Or EsGm(Userindex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
        End If
        Call .outgoingData.WriteByte(Visibilidad)
        If EsGm(Userindex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1
        End If
        Call .outgoingData.WriteByte(CanMakeSticky)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteSetInvisible(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteDiceRoll(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteMeditateToggle(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBlindNoMore(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteDumbNoMore(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteSendSkills(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i As Long
    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
        Call .outgoingData.WriteByte(.Clase)
        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(UserList(Userindex).Stats.UserSkills(i))
            If .Stats.UserSkills(i) < MAXSKILLPOINTS Then
                Call .outgoingData.WriteByte(Int(.Stats.ExpSkills(i) * 100 / .Stats.EluSkills(i)))
            Else
                Call .outgoingData.WriteByte(0)
            End If
        Next i
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteTrainerCreatureList(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Str As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            Str = Str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        If LenB(Str) > 0 Then Str = Left$(Str, Len(Str) - 1)
        Call .WriteASCIIString(Str)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteGuildNews(ByVal Userindex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        Call .WriteASCIIString(guildNews)
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
        Tmp = vbNullString
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteOfferDetails(ByVal Userindex As Integer, ByVal details As String)
    On Error GoTo ErrorHandler
    Dim i As Long
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        Call .WriteASCIIString(details)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteAlianceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePeaceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCharacterInfo(ByVal Userindex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal reputation As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        Call .WriteByte(level)
        Call .WriteLong(Gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteGuildLeaderInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String, ByVal guildNews As String, ByRef joinRequests() As String)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
        Call .WriteASCIIString(guildNews)
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteGuildMemberInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteGuildDetails(ByVal Userindex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, ByVal leader As String, ByVal URL As String, _
                             ByVal memberCount As Integer, ByVal electionsOpen As Boolean, ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)
    On Error GoTo ErrorHandler
    Dim i    As Long
    Dim temp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        Call .WriteASCIIString(alignment)
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        Call .WriteASCIIString(antifactionPoints)
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        If Len(temp) > 1 Then temp = Left$(temp, Len(temp) - 1)
        Call .WriteASCIIString(temp)
        Call .WriteASCIIString(guildDesc)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowGuildAlign(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowGuildFundationForm(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteParalizeOK(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ParalizeOK)
        Call .WriteInteger(IIf(UserList(Userindex).flags.Paralizado, UserList(Userindex).Counters.Paralisis, 0))
    End With
    Call WritePosUpdate(Userindex)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowUserRequest(ByVal Userindex As Integer, ByVal details As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        Call .WriteASCIIString(details)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeUserTradeSlot(ByVal Userindex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        Call .WriteByte(OfferSlot)
        Call .WriteInteger(ObjIndex)
        Call .WriteLong(Amount)
        If ObjIndex > 0 Then
            Call .WriteLong(ObjData(ObjIndex).GrhIndex)
            Call .WriteByte(ObjData(ObjIndex).OBJType)
            Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
            Call .WriteInteger(ObjData(ObjIndex).MinHIT)
            Call .WriteInteger(ObjData(ObjIndex).MaxDef)
            Call .WriteInteger(ObjData(ObjIndex).MinDef)
            Call .WriteLong(SalePrice(ObjIndex))
            Call .WriteASCIIString(ObjData(ObjIndex).Name)
            Call .WriteBoolean(ItemIncompatibleConUser(Userindex, ObjIndex))
        Else
            Call .WriteLong(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString("")
            Call .WriteBoolean(False)
        End If
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteSendNight(ByVal Userindex As Integer, ByVal night As Boolean)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteSpawnList(ByVal Userindex As Integer, ByRef npcNames() As String)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowSOSForm(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowDenounces(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim DenounceIndex As Long
    Dim DenounceList  As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowDenounces)
        For DenounceIndex = 1 To Denuncias.Longitud
            DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
        Next DenounceIndex
        If LenB(DenounceList) <> 0 Then DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
        Call .WriteASCIIString(DenounceList)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowPartyForm(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i                         As Long
    Dim Tmp                       As String
    Dim PI                        As Integer
    Dim members(PARTY_MAXMEMBERS) As Integer
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)
        PI = UserList(Userindex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(Userindex)))
        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())
            For i = 1 To PARTY_MAXMEMBERS
                If members(i) > 0 Then
                    Tmp = Tmp & UserList(members(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR
                End If
            Next i
        End If
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
        Call .WriteLong(Parties(PI).ObtenerExperienciaTotal)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowMOTDEditionForm(ByVal Userindex As Integer, ByVal currentMOTD As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        Call .WriteASCIIString(currentMOTD)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowGMPanelForm(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserNameList(ByVal Userindex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
    On Error GoTo ErrorHandler
    Dim i   As Long
    Dim Tmp As String
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePong(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Pong)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub FlushBuffer(ByVal Userindex As Integer)
    Dim sndData As String
    With UserList(Userindex).outgoingData
        If .Length = 0 Then Exit Sub
        sndData = .ReadASCIIStringFixed(.Length)
    End With
    Dim Ret As Long: Ret = WsApiEnviar(Userindex, sndData)
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(Userindex)
        Call Cerrar_Usuario(Userindex)
    End If
End Sub

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)
        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareRenderConsoleMsg(ByVal Chat As String, ByVal FontIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RenderMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(FontIndex)
        PrepareRenderConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, ByVal FontIndex As FontTypeNames) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessagePlayMp3(ByVal mp3 As Integer, Optional ByVal loops As Integer = -1) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMp3)
        Call .WriteInteger(mp3)
        Call .WriteInteger(loops)
        PrepareMessagePlayMp3 = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessagePlayMidi(ByVal midi As Integer, Optional ByVal loops As Integer = -1) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteInteger(midi)
        Call .WriteInteger(loops)
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessagePauseToggle() As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageRainToggle() As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Long, ByVal X As Byte, ByVal Y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteLong(GrhIndex)
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, _
                                              ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, ByVal Privileges As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, _
                                              ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareMessageHeadingChange(ByVal heading As eHeading, ByVal CharIndex As Integer)
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.HeadingChange)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(heading)
        PrepareMessageHeadingChange = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageUpdateTagAndStatus(ByVal Userindex As Integer, ByVal NickColor As Byte, ByRef Tag As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.errorMsg)
        Call .WriteASCIIString(Message)
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Sub WriteStopWorking(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.StopWorking)
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteCancelOfferItem(ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub HandleSetDialog(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim NewDialog As String
        NewDialog = buffer.ReadASCIIString
        Call .incomingData.CopyBuffer(buffer)
        If .flags.TargetNPC > 0 Then
            If Not ((.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster)) Then
                Npclist(.flags.TargetNPC).Desc = NewDialog
            End If
        End If
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleImpersonate(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNPC
        If NpcIndex = 0 Then Exit Sub
        Call ImitateNpc(Userindex, NpcIndex)
        Call WarpUserChar(Userindex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, False, True)
        Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        Call QuitarNPC(NpcIndex)
    End With
End Sub

Private Sub HandleImitate(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.RoleMaster)) <> (PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNPC
        If NpcIndex = 0 Then Exit Sub
        Call ImitateNpc(Userindex, NpcIndex)
        Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
    End With
End Sub
          
Public Sub HandleRecordAdd(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim Reason   As String
        UserName = buffer.ReadASCIIString
        Reason = buffer.ReadASCIIString
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
            If Not PersonajeExiste(UserName) Then
                Call WriteShowMessageBox(Userindex, "El personaje no existe")
            Else
                Call AddRecord(Userindex, UserName, Reason)
                Call WriteRecordList(Userindex)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleRecordAddObs(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim RecordIndex As Byte
        Dim Obs         As String
        RecordIndex = buffer.ReadByte
        Obs = buffer.ReadASCIIString
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
            Call AddObs(Userindex, RecordIndex, Obs)
            Call WriteRecordDetails(Userindex, RecordIndex)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleRecordRemove(ByVal Userindex As Integer)
    Dim RecordIndex As Integer
    With UserList(Userindex)
        Call .incomingData.ReadByte
        RecordIndex = .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(Userindex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(Userindex)
        Else
            Call WriteShowMessageBox(Userindex, "Solo los dioses pueden eliminar seguimientos.")
        End If
    End With
End Sub
        
Public Sub HandleRecordListRequest(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        Call WriteRecordList(Userindex)
    End With
End Sub

Public Sub WriteRecordDetails(ByVal Userindex As Integer, ByVal RecordIndex As Integer)
    On Error GoTo ErrorHandler
    Dim i        As Long
    Dim tIndex   As Integer
    Dim tmpStr   As String
    Dim TempDate As Date
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.RecordDetails)
        Call .WriteASCIIString(Records(RecordIndex).Creador)
        Call .WriteASCIIString(Records(RecordIndex).Motivo)
        tIndex = NameIndex(Records(RecordIndex).Usuario)
        Call .WriteBoolean(tIndex > 0)
        If tIndex > 0 Then
            tmpStr = UserList(tIndex).IP
        Else
            tmpStr = vbNullString
        End If
        Call .WriteASCIIString(tmpStr)
        If tIndex > 0 Then
            TempDate = Now - UserList(tIndex).LogOnTime
            tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
        Else
            tmpStr = vbNullString
        End If
        Call .WriteASCIIString(tmpStr)
        tmpStr = vbNullString
        If Records(RecordIndex).NumObs Then
            For i = 1 To Records(RecordIndex).NumObs
                tmpStr = tmpStr & Records(RecordIndex).Obs(i).Creador & "> " & Records(RecordIndex).Obs(i).Detalles & vbCrLf
            Next i
            tmpStr = Left$(tmpStr, Len(tmpStr) - 1)
        End If
        Call .WriteASCIIString(tmpStr)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteRecordList(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i As Long
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.RecordList)
        Call .WriteByte(NumRecords)
        For i = 1 To NumRecords
            Call .WriteASCIIString(Records(i).Usuario)
        Next i
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub HandleRecordDetailsRequest(ByVal Userindex As Integer)
    Dim RecordIndex As Byte
    With UserList(Userindex)
        Call .incomingData.ReadByte
        RecordIndex = .incomingData.ReadByte
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        Call WriteRecordDetails(Userindex, RecordIndex)
    End With
End Sub

Public Sub HandleMoveItem(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim originalSlot As Byte
        Dim newSlot      As Byte
        Call .incomingData.ReadByte
        originalSlot = .incomingData.ReadByte
        newSlot = .incomingData.ReadByte
        Call .incomingData.ReadByte
        Call InvUsuario.moveItem(Userindex, originalSlot, newSlot)
    End With
End Sub

Private Sub HandleLoginExistingAccount(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Call Err.Raise(UserList(Userindex).incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte
    Dim UserName As String
    Dim Password As String
    Dim version  As String
    UserName = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    If Not CuentaExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "La cuenta no existe.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call ConnectAccount(Userindex, UserName, Password)
    End If
    Exit Sub
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleLoginNewAccount(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    Call buffer.ReadByte
    Dim UserName As String
    Dim Password As String
    Dim version  As String
    UserName = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    If CuentaExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "La cuenta ya existe.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    If Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call CreateNewAccount(Userindex, UserName, Password)
    End If
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub WriteUserAccountLogged(ByVal Userindex As Integer, ByVal UserName As String, ByVal AccountHash As String, ByVal NumberOfCharacters As Byte, ByRef Characters() As AccountUser)
    On Error GoTo ErrorHandler
    Dim i As Long
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AccountLogged)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(AccountHash)
        Call .WriteByte(NumberOfCharacters)
        If NumberOfCharacters > 0 Then
            For i = 1 To NumberOfCharacters
                Call .WriteASCIIString(Characters(i).Name)
                Call .WriteInteger(Characters(i).body)
                Call .WriteInteger(Characters(i).Head)
                Call .WriteInteger(Characters(i).weapon)
                Call .WriteInteger(Characters(i).shield)
                Call .WriteInteger(Characters(i).helmet)
                Call .WriteByte(Characters(i).Class)
                Call .WriteByte(Characters(i).race)
                Call .WriteInteger(Characters(i).Map)
                Call .WriteByte(Characters(i).level)
                Call .WriteLong(Characters(i).Gold)
                Call .WriteBoolean(Characters(i).criminal)
                Call .WriteBoolean(Characters(i).dead)
                Call .WriteBoolean(Characters(i).gameMaster)
            Next i
        End If
        Call SaveLastIpsAccountCharfile(UserName, UserList(Userindex).IP)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Function PrepareMessagePalabrasMagicas(ByVal SpellIndex As Byte, ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PalabrasMagicas)
        Call .WriteByte(SpellIndex)
        Call .WriteInteger(CharIndex)
        PrepareMessagePalabrasMagicas = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterAttackAnim(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayAttackAnim)
        Call .WriteInteger(CharIndex)
        PrepareMessageCharacterAttackAnim = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageFXtoMap(ByVal FxIndex As Integer, ByVal loops As Byte, ByVal X As Integer, ByVal Y As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.FXtoMap)
        Call .WriteByte(loops)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(FxIndex)
        PrepareMessageFXtoMap = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function WriteSearchList(ByVal Userindex As Integer, ByVal Num As Integer, ByVal Datos As String, ByVal obj As Boolean) As String
    On Error GoTo ErrorHandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SearchList)
        Call .WriteInteger(Num)
        Call .WriteBoolean(obj)
        Call .WriteASCIIString(Datos)
    End With
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Function
 
Public Sub HandleSearchNpc(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim i       As Long
        Dim n       As Integer
        Dim Name    As String
        Dim UserNpc As String
        Dim tStr    As String
        UserNpc = buffer.ReadASCIIString()
        tStr = Tilde(UserNpc)
        For i = 1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))
            Name = LeerNPCs.GetValue("NPC" & i, "Name")
            If InStr(1, Tilde(Name), tStr) Then
                Call WriteSearchList(Userindex, i, CStr(i & " - " & Name), False)
                n = n + 1
            End If
        Next i
        If n = 0 Then
            Call WriteSearchList(Userindex, 0, "No hubo resultados de la busqueda.", False)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub
 
Private Sub HandleSearchObj(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserObj As String
        Dim tUser   As Integer
        Dim n       As Integer
        Dim i       As Long
        Dim tStr    As String
        UserObj = buffer.ReadASCIIString()
        tStr = Tilde(UserObj)
        For i = 1 To UBound(ObjData)
            If InStr(1, Tilde(ObjData(i).Name), tStr) Then
                Call WriteSearchList(Userindex, i, CStr(i & " - " & ObjData(i).Name), True)
                n = n + 1
            End If
        Next
        If n = 0 Then
            Call WriteSearchList(Userindex, 0, "No hubo resultados de la busqueda.", False)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleEnviaCvc(ByVal Userindex As Integer)
    With UserList(Userindex)
        .incomingData.ReadByte
        If .flags.TargetUser = 0 Then Exit Sub
        Call Mod_ClanvsClan.Enviar(Userindex, .flags.TargetUser)
    End With
End Sub

Private Sub HandleAceptarCvc(ByVal Userindex As Integer)
    With UserList(Userindex)
        .incomingData.ReadByte
        If .flags.TargetUser = 0 Then Exit Sub
        Call Mod_ClanvsClan.Aceptar(Userindex, .flags.TargetUser)
    End With
End Sub

Private Sub HandleIrCvc(ByVal Userindex As Integer)
    With UserList(Userindex)
        .incomingData.ReadByte
        Call Mod_ClanvsClan.ConectarCVC(Userindex, True)
    End With
End Sub

Public Sub HandleDragAndDropHechizos(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim AnteriorPosicion As Integer: AnteriorPosicion = .incomingData.ReadInteger
        Dim NuevaPosicion As Integer: NuevaPosicion = .incomingData.ReadInteger
        Dim Hechizo As Integer: Hechizo = .Stats.UserHechizos(NuevaPosicion)
        .Stats.UserHechizos(NuevaPosicion) = .Stats.UserHechizos(AnteriorPosicion)
        .Stats.UserHechizos(AnteriorPosicion) = Hechizo
    End With
End Sub

Public Sub WriteQuestDetails(ByVal Userindex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)
    On Error GoTo ErrorHandler
    Dim i As Integer
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.QuestDetails)
        Call .WriteByte(IIf(QuestSlot, 1, 0))
        Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        Call .WriteASCIIString(QuestList(QuestIndex).Desc)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)
        If QuestList(QuestIndex).RequiredNPCs Then
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NpcIndex, "Name"))
                If QuestSlot Then
                    Call .WriteInteger(UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                End If
            Next i
        End If
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex).Name)
            Next i
        End If
        Call .WriteLong(QuestList(QuestIndex).RewardGLD)
        Call .WriteLong(QuestList(QuestIndex).RewardEXP)
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)
        If QuestList(QuestIndex).RewardOBJs Then
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name)
            Next i
        End If
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
 
Public Sub WriteQuestListSend(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i       As Integer
    Dim tmpStr  As String
    Dim tmpByte As Byte
    With UserList(Userindex)
        .outgoingData.WriteByte ServerPacketID.QuestListSend
        For i = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).Nombre & "-"
            End If
        Next i
        Call .outgoingData.WriteByte(tmpByte)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))
        End If
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Function PrepareMessageCreateDamage(ByVal X As Byte, ByVal Y As Byte, ByVal DamageValue As Long, ByVal DamageType As Byte)
    With auxiliarBuffer
         .WriteByte ServerPacketID.CreateDamage
         .WriteByte X
         .WriteByte Y
         .WriteLong DamageValue
         .WriteByte DamageType
         PrepareMessageCreateDamage = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Sub HandleCambiarContrasena(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Correo As String
    Dim NuevaContrasena As String
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Correo = buffer.ReadASCIIString()
        NuevaContrasena = buffer.ReadASCIIString()
        If ConexionAPI Then
            If Not CuentaExiste(Correo) Then
                Call WriteErrorMsg(Userindex, "La cuenta no existe.")
                Call CloseSocket(Userindex)
                Exit Sub
            End If
            Call ApiEndpointSendResetPasswordAccountEmail(Correo, NuevaContrasena)
            Call WriteErrorMsg(Userindex, "Se ha enviado un correo electronico a: " & Correo & " donde debera confirmar el cambio de la password de su cuenta.")
        Else
            Call WriteErrorMsg(Userindex, "Esta funcion se encuentra deshabilitada actualmente, si sos el administrador del servidor necesitas habilitar la API hecha en Node.js (https://github.com/ao-libre/ao-api-server).")
        End If
        Call .incomingData.CopyBuffer(buffer)
        Call CloseSocket(Userindex)
    End With
ErrorHandler:
    Dim Error As Long: Error = Err.Number
    Call CloseSocket(Userindex)
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub WriteUserInEvent(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserInEvent)
    Exit Sub
ErrorHandler:
        If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(Userindex)
            Resume
        End If
End Sub

Private Sub HandleFightSend(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim ListUsers As String
        Dim GldRequired As Long
        Dim Users() As String
        ListUsers = buffer.ReadASCIIString & "-" & .Name
        GldRequired = buffer.ReadLong
        If Len(ListUsers) >= 1 Then
            Users = Split(ListUsers, "-")
            Call Retos.SendFight(Userindex, GldRequired, Users)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleFightAccept(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        UserName = buffer.ReadASCIIString
        If Len(UserName) >= 1 Then
            Call Retos.AcceptFight(Userindex, UserName)
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCloseGuild(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        Dim i As Long
        Dim PreviousGuildIndex  As Integer
        If Not .GuildIndex >= 1 Then
            Call WriteConsoleMsg(Userindex, "No perteneces a ningun clan.", FONTTYPE_GUILD)
            Exit Sub
        End If
        If guilds(.GuildIndex).Fundador <> .Name Then
            Call WriteConsoleMsg(Userindex, "No eres lider del clan.", FONTTYPE_GUILD)
            Exit Sub
        End If
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "GuildName", "CLAN CERRADO")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Founder", "NADIE")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Leader", "NADIE")
        PreviousGuildIndex = .GuildIndex
        Dim GuildMembers() As String
            GuildMembers = guilds(PreviousGuildIndex).GetMemberList()
        For i = 0 To UBound(GuildMembers)
            Call SaveUserGuildIndex(GuildMembers(i), 0)
            Call SaveUserGuildAspirant(GuildMembers(i), 0)
        Next i
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-members.mem")
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-solicitudes.sol")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Clan " & guilds(.GuildIndex).GuildName & " ha cerrado sus puertas.", FontTypeNames.FONTTYPE_GUILD))
    End With
    Call modGuilds.LoadGuildsDB
    Exit Sub
End Sub

Private Sub HandleDiscord(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Chat As String
        Chat = buffer.ReadASCIIString()
        If LenB(Chat) <> 0 Then
            Call Statistics.ParseChat(Chat)
            If ConexionAPI Then
                Call ApiEndpointSendCustomCharacterMessageDiscord(Chat, .Name, .Desc)
                Call WriteConsoleMsg(Userindex, "Link Discord: https://discord.gg/xbAuHcf - El bot de Discord recibio y envio lo siguiente: " & Chat, FontTypeNames.FONTTYPE_INFOBOLD)
            Else
                Call WriteConsoleMsg(Userindex, "(api - node.js)  El modulo para usar esta funcion no esta instalado en este servidor. http://www.github.com/ao-libre/ao-api-server para mas informacion / more info.", FontTypeNames.FONTTYPE_INFOBOLD)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub HandleLimpiarMundo(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    If Not EsGm(Userindex) Then Exit Sub
    Call LogGM(UserList(Userindex).Name, "forzo la limpieza del mundo.")
    tickLimpieza = 16
End Sub

Public Sub WriteEquitandoToggle(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.EquitandoToggle)
        Call .outgoingData.WriteLong(.Counters.MonturaCounter)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Private Sub HandleObtenerDatosServer(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Call WriteEnviarDatosServer(Userindex)
    End With
ErrorHandler:
    Dim Error As Long: Error = Err.Number
    Call CloseSocket(Userindex)
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCraftsmanCreate(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(Userindex).incomingData
        Call .ReadByte
        Dim Item As Integer
        Item = .ReadInteger()
        If Item < LBound(ObjArtesano) Or Item > UBound(ObjArtesano) Then Exit Sub
        If UserList(Userindex).flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Sub
        End If
        If UserList(Userindex).flags.TargetNPC = 0 Then Exit Sub
        If Npclist(UserList(Userindex).flags.TargetNPC).NPCtype <> eNPCType.Artesano Then Exit Sub
        If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 3 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del artesano.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call ArtesanoConstruirItem(Userindex, Item)
    End With
End Sub

Private Sub WriteEnviarDatosServer(ByVal Userindex As Integer)
    Dim MundoSeleccionadoWithoutPath As String
    MundoSeleccionadoWithoutPath = Replace(MundoSeleccionado, "\Mundos\", "")
    MundoSeleccionadoWithoutPath = Replace(MundoSeleccionadoWithoutPath, "\", "")
    With UserList(Userindex)
        Call .outgoingData.WriteByte(ServerPacketID.EnviarDatosServer)
        Call .outgoingData.WriteASCIIString(MundoSeleccionadoWithoutPath)
        Call .outgoingData.WriteASCIIString(NombreServidor)
        Call .outgoingData.WriteASCIIString(DescripcionServidor)
        Call .outgoingData.WriteInteger(STAT_MAXELV)
        Call .outgoingData.WriteInteger(MaxUsers)
        Call .outgoingData.WriteInteger(LastUser - 1)
        Call .outgoingData.WriteInteger(ExpMultiplier)
        Call .outgoingData.WriteInteger(OroMultiplier)
        Call .outgoingData.WriteInteger(OficioMultiplier)
        Call CloseSocket(Userindex)
    End With
End Sub


Public Sub WriteCargarListaDeAmigos(ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    Dim i As Integer
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.EnviarListDeAmigos)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(UserList(Userindex).Amigos(Slot).Nombre)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteSeeInProcess(ByVal Userindex As Integer)
On Error GoTo ErrorHandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SeeInProcess)
Exit Sub
ErrorHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
     
Private Sub HandleSendProcessList(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 5 Then
       Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
       Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Captions As String, Process As String
        Captions = buffer.ReadASCIIString
        Process = buffer.ReadASCIIString
        If .flags.GMRequested > 0 Then
            If UserList(.flags.GMRequested).ConnIDValida Then
                Call WriteShowProcess(.flags.GMRequested, Captions, Process)
                .flags.GMRequested = 0
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub
            
Private Sub HandleLookProcess(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim tName As String
        Dim tIndex As Integer
        tName = buffer.ReadASCIIString
        If EsGm(Userindex) Then
            tIndex = NameIndex(tName)
            If tIndex > 0 Then
                UserList(tIndex).flags.GMRequested = Userindex
                Call WriteSeeInProcess(tIndex)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
    Exit Sub
ErrorHandler:
    Dim Error As Long
    Error = Err.Number
    On Error GoTo 0
    Set buffer = Nothing
    If Error <> 0 Then Err.Raise Error
    LogError ("Error en HandleLookProcess. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub WriteShowProcess(ByVal gmIndex As Integer, ByVal strCaptions As String, ByVal strProcess As String)
    On Error GoTo ErrorHandler
    With UserList(gmIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowProcess)
        Call .WriteASCIIString(strCaptions)
        Call .WriteASCIIString(strProcess)
    End With
    Exit Sub
ErrorHandler:
    If Err.Number = UserList(gmIndex).outgoingData.NotEnoughSpaceErrCode Then Call FlushBuffer(gmIndex): Resume
End Sub

Public Function PrepareMessageProyectil(ByVal Userindex As Integer, ByVal CharSending As Integer, ByVal CharRecieved As Integer, ByVal GrhIndex As Integer) As String
    With auxiliarBuffer
        .WriteByte (ServerPacketID.proyectil)
        .WriteInteger (CharSending)
        .WriteInteger (CharRecieved)
        .WriteInteger (GrhIndex)
        PrepareMessageProyectil = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterIsInChatMode(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayIsInChatMode)
        Call .WriteInteger(CharIndex)
        PrepareMessageCharacterIsInChatMode = .ReadASCIIStringFixed(.Length)
    End With
End Function

Private Sub HandleSendIfCharIsInChatMode(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        Call .incomingData.ReadByte
        .Char.Escribiendo = IIf(.Char.Escribiendo = 1, 0, 1)
        Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageSetTypingFlagToCharIndex(.Char.CharIndex, .Char.Escribiendo))
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure HandleSendIfCharIsInChatMode of Modulo Protocol " & Erl & ".")
End Sub

Private Function PrepareMessageSetTypingFlagToCharIndex(ByVal CharIndex As Integer, ByVal Escribiendo As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayIsInChatMode)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Escribiendo)
        PrepareMessageSetTypingFlagToCharIndex = .ReadASCIIStringFixed(.Length)
    End With
End Function
