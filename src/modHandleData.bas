Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CSetName) = GetAddress(AddressOf HandleSetName)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    
    HandleDataSub(CRequestRank) = GetAddress(AddressOf HandleRequestRank)
    
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CPlayerVisibility) = GetAddress(AddressOf HandlePlayerVisibility)
    HandleDataSub(CHealPlayer) = GetAddress(AddressOf HandleHealPlayer)
    HandleDataSub(CKillPlayer) = GetAddress(AddressOf HandleKillPlayer)
    HandleDataSub(CSayGuild) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CGuildCommand) = GetAddress(AddressOf HandleGuildCommands)
    HandleDataSub(CSaveGuild) = GetAddress(AddressOf HandleGuildSave)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    HandleDataSub(CNewClass) = GetAddress(AddressOf HandleNewClass)
    
HandleDataSub(CClientOpenShop) = GetAddress(AddressOf HandleClientOpenShop)

End Sub

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar index
                Else
                    ' send new char shit
                    If Not IsPlaying(index) Then
                        Call SendNewCharClasses(index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            Else
                Call AlertMsg(index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(index, Name)

            If LenB(Trim$(Player(index).Name)) > 0 Then
                Call DeleteName(Player(index).Name)
            End If

            Call ClearPlayer(index)
            ' Everything went ok
            Call Kill(App.path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Trim$(Buffer.ReadString)
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(index, Name)
            ClearBank index
            LoadBank index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(Player(index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar index
            Else
                ' send new char shit
                If Not IsPlaying(index) Then
                    Call SendNewCharClasses(index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(index) Then
            Call AlertMsg(index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Sex, Class, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(index), index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Msg, White)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(index) & ": " & Msg
    Call SayMsg_Global(index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(index).StunDuration > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Prevent player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If

    Call PlayerMove(index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim InvNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem index, InvNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    Dim Shoot As Boolean
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index
    
    Shoot = False
    
    If TempPlayer(index).Target > 0 Then
        If GetPlayerEquipment(index, weapon) > 0 Then
            If Item(GetPlayerEquipment(index, weapon)).Projectile > 0 Then
                If Item(GetPlayerEquipment(index, weapon)).Ammo > 0 Then
                    If HasItem(index, Item(GetPlayerEquipment(index, weapon)).Ammo) Then
                        TakeInvItem index, Item(GetPlayerEquipment(index, weapon)).Ammo, 1
                        Shoot = True
                    Else
                        PlayerMsg index, "Out of ammo!", BrightRed
                    End If
                Else
                    Shoot = True
                End If
            End If
        End If
    End If
    
    If Shoot = True Then
        Select Case TempPlayer(index).targetType
            Case TARGET_TYPE_NPC: TryPlayerShootNpc index, TempPlayer(index).Target
            Case TARGET_TYPE_PLAYER: TryPlayerShootPlayer index, TempPlayer(index).Target
        End Select
        Exit Sub
    End If

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> index Then
            TryPlayerAttackPlayer index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, i
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    CheckResource index, x, y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(index, PointType) >= 30000 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString 'Parse(1)
    Set Buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
    Exit Sub
    End If
    
    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    i = FindPlayer(Buffer.ReadString)
    Set Buffer = Nothing
    
    Call SetPlayerSprite(i, n)
    Call SendPlayerData(i)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim mapnum As Long
    Dim x As Long
    Dim y As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(index)
    i = Map(mapnum).Revision + 1
    Call ClearMap(mapnum)
    
    Map(mapnum).Name = Buffer.ReadString
    Map(mapnum).Music = Buffer.ReadString
    Map(mapnum).BGS = Buffer.ReadString
    Map(mapnum).Revision = i
    Map(mapnum).Moral = Buffer.ReadByte
    Map(mapnum).Up = Buffer.ReadLong
    Map(mapnum).Down = Buffer.ReadLong
    Map(mapnum).Left = Buffer.ReadLong
    Map(mapnum).Right = Buffer.ReadLong
    Map(mapnum).BootMap = Buffer.ReadLong
    Map(mapnum).BootX = Buffer.ReadByte
    Map(mapnum).BootY = Buffer.ReadByte
    
    Map(mapnum).Weather = Buffer.ReadLong
    Map(mapnum).WeatherIntensity = Buffer.ReadLong
    
    Map(mapnum).Fog = Buffer.ReadLong
    Map(mapnum).FogSpeed = Buffer.ReadLong
    Map(mapnum).FogOpacity = Buffer.ReadLong
    
    Map(mapnum).Red = Buffer.ReadLong
    Map(mapnum).Green = Buffer.ReadLong
    Map(mapnum).Blue = Buffer.ReadLong
    Map(mapnum).Alpha = Buffer.ReadLong
    
    Map(mapnum).MaxX = Buffer.ReadByte
    Map(mapnum).MaxY = Buffer.ReadByte
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Layer(i).x = Buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(i).y = Buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(i).Tileset = Buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Autotile(z) = Buffer.ReadLong
            Next
            Map(mapnum).Tile(x, y).Type = Buffer.ReadByte
            Map(mapnum).Tile(x, y).Data1 = Buffer.ReadLong
            Map(mapnum).Tile(x, y).Data2 = Buffer.ReadLong
            Map(mapnum).Tile(x, y).Data3 = Buffer.ReadLong
            Map(mapnum).Tile(x, y).Data4 = Buffer.ReadString
            Map(mapnum).Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(mapnum).NPC(x) = Buffer.ReadLong
        Map(mapnum).NpcSpawnType(x) = Buffer.ReadLong
        Call ClearMapNpc(x, mapnum)
    Next
    
    'Event Data!
    Map(mapnum).EventCount = Buffer.ReadLong
        
    If Map(mapnum).EventCount > 0 Then
        ReDim Map(mapnum).Events(0 To Map(mapnum).EventCount)
        For i = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(i)
                .Name = Buffer.ReadString
                .Global = Buffer.ReadLong
                .x = Buffer.ReadLong
                .y = Buffer.ReadLong
                .PageCount = Buffer.ReadLong
            End With
            If Map(mapnum).Events(i).PageCount > 0 Then
                ReDim Map(mapnum).Events(i).Pages(0 To Map(mapnum).Events(i).PageCount)
                For x = 1 To Map(mapnum).Events(i).PageCount
                    With Map(mapnum).Events(i).Pages(x)
                        .chkVariable = Buffer.ReadLong
                        .VariableIndex = Buffer.ReadLong
                        .VariableCondition = Buffer.ReadLong
                        .VariableCompare = Buffer.ReadLong
                            
                        .chkSwitch = Buffer.ReadLong
                        .SwitchIndex = Buffer.ReadLong
                        .SwitchCompare = Buffer.ReadLong
                            
                        .chkHasItem = Buffer.ReadLong
                        .HasItemIndex = Buffer.ReadLong
                            
                        .chkSelfSwitch = Buffer.ReadLong
                        .SelfSwitchIndex = Buffer.ReadLong
                        .SelfSwitchCompare = Buffer.ReadLong
                            
                        .GraphicType = Buffer.ReadLong
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .GraphicX2 = Buffer.ReadLong
                        .GraphicY2 = Buffer.ReadLong
                            
                        .MoveType = Buffer.ReadLong
                        .MoveSpeed = Buffer.ReadLong
                        .MoveFreq = Buffer.ReadLong
                            
                        .MoveRouteCount = Buffer.ReadLong
                        
                        .IgnoreMoveRoute = Buffer.ReadLong
                        .RepeatMoveRoute = Buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(mapnum).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                            For y = 1 To .MoveRouteCount
                                .MoveRoute(y).index = Buffer.ReadLong
                                .MoveRoute(y).Data1 = Buffer.ReadLong
                                .MoveRoute(y).Data2 = Buffer.ReadLong
                                .MoveRoute(y).Data3 = Buffer.ReadLong
                                .MoveRoute(y).Data4 = Buffer.ReadLong
                                .MoveRoute(y).data5 = Buffer.ReadLong
                                .MoveRoute(y).data6 = Buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = Buffer.ReadLong
                        .DirFix = Buffer.ReadLong
                        .WalkThrough = Buffer.ReadLong
                        .ShowName = Buffer.ReadLong
                        .Trigger = Buffer.ReadLong
                        .CommandListCount = Buffer.ReadLong
                            
                        .Position = Buffer.ReadLong
                    End With
                        
                    If Map(mapnum).Events(i).Pages(x).CommandListCount > 0 Then
                        ReDim Map(mapnum).Events(i).Pages(x).CommandList(0 To Map(mapnum).Events(i).Pages(x).CommandListCount)
                        For y = 1 To Map(mapnum).Events(i).Pages(x).CommandListCount
                            Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount = Buffer.ReadLong
                            Map(mapnum).Events(i).Pages(x).CommandList(y).ParentList = Buffer.ReadLong
                            If Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                ReDim Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount)
                                For z = 1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        .index = Buffer.ReadLong
                                        .Text1 = Buffer.ReadString
                                        .Text2 = Buffer.ReadString
                                        .Text3 = Buffer.ReadString
                                        .Text4 = Buffer.ReadString
                                        .Text5 = Buffer.ReadString
                                        .Data1 = Buffer.ReadLong
                                        .Data2 = Buffer.ReadLong
                                        .Data3 = Buffer.ReadLong
                                        .Data4 = Buffer.ReadLong
                                        .data5 = Buffer.ReadLong
                                        .data6 = Buffer.ReadLong
                                        .ConditionalBranch.CommandList = Buffer.ReadLong
                                        .ConditionalBranch.Condition = Buffer.ReadLong
                                        .ConditionalBranch.Data1 = Buffer.ReadLong
                                        .ConditionalBranch.Data2 = Buffer.ReadLong
                                        .ConditionalBranch.Data3 = Buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = Buffer.ReadLong
                                        .MoveRouteCount = Buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).index = Buffer.ReadLong
                                                .MoveRoute(w).Data1 = Buffer.ReadLong
                                                .MoveRoute(w).Data2 = Buffer.ReadLong
                                                .MoveRoute(w).Data3 = Buffer.ReadLong
                                                .MoveRoute(w).Data4 = Buffer.ReadLong
                                                .MoveRoute(w).data5 = Buffer.ReadLong
                                                .MoveRoute(w).data6 = Buffer.ReadLong
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    'End Event Data
    
    Map(mapnum).IsMember = Buffer.ReadByte
    Map(mapnum).DayNight = Buffer.ReadByte
    Map(mapnum).Panorama = Buffer.ReadByte

    Call SendMapNpcsToMap(mapnum)
    Call SpawnMapNpcs(mapnum)
    Call SpawnGlobalEvents(mapnum)
    
    For i = 1 To Player_HighIndex
        If Player(i).Map = mapnum Then
            SpawnMapEventsFor i, mapnum
        End If
    Next

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))
    ' Save the map
    Call SaveMap(mapnum)
    Call MapCache_Create(mapnum)
    Call ClearTempTile(mapnum)
    Call CacheResources(mapnum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            Call PlayerWarp(i, mapnum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    Call CacheMapBlocks(mapnum)

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SpawnMapEventsFor(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next

    TempPlayer(index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, InvNum) < 1 Or GetPlayerInvItemNum(index, InvNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(index, InvNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    Call PlayerMsg(index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapReport
    
    For i = 1 To MAX_MAPS
        Buffer.WriteString Trim$(Map(i).Name)
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(index) & "!", White)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "That is a higher or same access admin than you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim f As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    f = FreeFile
    Open App.path & "\data\banlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s
        Input #f, Name
        Call PlayerMsg(index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #f
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim f As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    Kill filename
    Call PlayerMsg(index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call BanIndex(n, index)
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    SendMapEventData (index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NPCNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    NPCNum = Buffer.ReadLong

    ' Prevent hacking
    If NPCNum < 0 Or NPCNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(NPCNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(NPCNum)
    Call SaveNpc(NPCNum)
    Call AddLog(GetPlayerName(index) & " saved Npc #" & NPCNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = Buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SpellNum = Buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(index) Then
                Call PlayerMsg(index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Set name packet ::
' :::::::::::::::::::::::
Sub HandleSetName(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The name
    i = Buffer.ReadString 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check if player is on
    If n > 0 Then

        'check to see if same level access is trying to change another access of the very same level and boot them if they are.
        If GetPlayerAccess(n) = GetPlayerAccess(index) Then
            Call PlayerMsg(index, "Invalid access level.", Red)
            Exit Sub
        End If
            
        Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s name too " & i & ".", ADMIN_LOG)
        Call SetPlayerName(n, i)
        Call SendPlayerData(n)
            
        If GetPlayerAccess(n) <= 0 Then
            Call PlayerMsg(n, "Your Name has been changed.", White)
        End If
    Else
        Call PlayerMsg(index, "Player is not online.", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong 'CLng(Parse(1))
    y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(index) = GetPlayerMap(i) Then
                If Not GetPlayerVisible(i) = 1 Then
                    If GetPlayerX(i) = x Then
                        If GetPlayerY(i) = y Then
                        ' Change target
                            If TempPlayer(index).targetType = TARGET_TYPE_PLAYER And TempPlayer(index).Target = i Then
                                TempPlayer(index).Target = 0
                                TempPlayer(index).targetType = TARGET_TYPE_NONE
                                ' send target to player
                                SendTarget index
                            Else
                                TempPlayer(index).Target = i
                                TempPlayer(index).targetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                SendTarget index
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).NPC(i).num > 0 Then
            If MapNpc(GetPlayerMap(index)).NPC(i).x = x Then
                If MapNpc(GetPlayerMap(index)).NPC(i).y = y Then
                    If TempPlayer(index).Target = i And TempPlayer(index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).Target = 0
                        TempPlayer(index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget index
                    Else
                        ' Change target
                        TempPlayer(index).Target = i
                        TempPlayer(index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    i = Buffer.ReadLong
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, n, i)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(n) > GetTickCount Then
            PlayerMsg index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem index, Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    Set Buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim thePlr As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    thePlr = FindPlayer(Buffer.ReadString)

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    SetPlayerExp thePlr, GetPlayerNextLevel(thePlr)
    CheckPlayerLevelUp thePlr
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellSlot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.Spell = spellSlot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(index).Spell(spellSlot) = 0
    Player(index).SpellUses(spellSlot) = 0
    SendPlayerSpells index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemamount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemamount = HasItem(index, .costitem)
        If itemamount = 0 Or itemamount < .costvalue Then
            PlayerMsg index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem index, .costitem, .costvalue
        GiveInvItem index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim itemnum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemnum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    price = Item(itemnum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg index, "The shop doesn't want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, itemnum, 1
    GiveInvItem index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    TakeBankItem index, BankSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    GiveBankItem index, invSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank index
    SavePlayer index
    
    TempPlayer(index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
    If GetPlayerAccess(index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, x
        SetPlayerY index, y
        SendPlayerXYToMap index
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).Target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    If Player(index).Transformed = True Then
        PlayerMsg index, "Kamu tidak bisa trade dalam keadaan mount .", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(index).x
    sY = Player(index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    If TempPlayer(index).InTrade > 0 Then
        TempPlayer(index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(index).TradeRequest
        ' let them know they're trading
        PlayerMsg index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
        PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has accepted your trade request.", BrightGreen
        ' clear the tradeRequest server-side
        TempPlayer(index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they're trading with each other
        TempPlayer(index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = index
        ' clear out their trade offers
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade index, tradeTarget
        SendTrade tradeTarget, index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate index, 0
        SendTradeUpdate index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " has declined your trade request.", BrightRed
    PlayerMsg index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemnum As Long
    
    TempPlayer(index).AcceptTrade = True
    
    tradeTarget = TempPlayer(index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(index).TradeOffer(i).num > 0 Then
                itemnum = Player(index).Inv(TempPlayer(index).TradeOffer(i).num).num
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem(i).num = itemnum
                    tmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot index, TempPlayer(index).TradeOffer(i).num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).num > 0 Then
                itemnum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).num = itemnum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num, tmpTradeItem2(i).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).num > 0 Then
                ' give away!
                GiveInvItem index, tmpTradeItem2(i).num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).num, tmpTradeItem(i).Value, False
            End If
        Next
    
        SendInventory index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg index, "Trade completed.", BrightGreen
        PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
            
    End If
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg index, "You declined the trade.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim itemnum As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(index, invSlot)
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).num = invSlot Then
                PlayerMsg index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(index).Hotbar(hotbarNum).Slot = 0
            Player(index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(index).Inv(Slot).num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Inv(Slot).num
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(index).Spell(Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Spell(Slot)
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case Player(index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(index).Inv(i).num > 0 Then
                    If Player(index).Inv(i).num = Player(index).Hotbar(Slot).Slot Then
                        'If Item(Player(index).Inv(i).num).Type = ITEM_TYPE_CONSUME Then
                         '   Player(index).Hotbar(Slot).Slot = 0
                          '  Player(index).Hotbar(Slot).sType = 0
                           ' SendHotbar index
                        'End If
                        UseItem index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(index).Spell(i) > 0 Then
                    If Player(index).Spell(i) = Player(index).Hotbar(Slot).Slot Then
                        BufferSpell index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).Target = index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(index).Target) Or Not IsPlaying(TempPlayer(index).Target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).Target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).Target
End Sub
Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If TempPlayer(index).inParty Then
        PlayerMsg index, "You are already in a party!", BrightRed
        Exit Sub
    End If
    Party_InviteAccept TempPlayer(index).partyInvite, index
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Sub HandleEventChatReply(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim eventID As Long, pageID As Long, reply As Long, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    eventID = Buffer.ReadLong
    pageID = Buffer.ReadLong
    reply = Buffer.ReadLong
    
    If TempPlayer(index).EventProcessingCount > 0 Then
        For i = 1 To TempPlayer(index).EventProcessingCount
            If TempPlayer(index).EventProcessing(i).eventID = eventID And TempPlayer(index).EventProcessing(i).pageID = pageID Then
                If TempPlayer(index).EventProcessing(i).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).index = EventType.evShowText Then
                            TempPlayer(index).EventProcessing(i).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data1
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 2
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data2
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 3
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data3
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 4
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data4
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(index).EventProcessing(i).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    
    
    
    Set Buffer = Nothing
End Sub

Sub HandleEvent(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long, begineventprocessing As Boolean, z As Long, Buffer As clsBuffer

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    i = Buffer.ReadLong
    Set Buffer = Nothing
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
            If TempPlayer(index).EventMap.EventPages(z).eventID = i Then
                i = z
                begineventprocessing = True
                Exit For
            End If
        Next
    End If
    
    If begineventprocessing = True Then
        If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
            'Process this event, it is action button and everything checks out.
            TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
            ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
            With TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                .ActionTimer = GetTickCount
                .CurList = 1
                .CurSlot = 1
                .eventID = TempPlayer(index).EventMap.EventPages(i).eventID
                .pageID = TempPlayer(index).EventMap.EventPages(i).pageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount)
            End With
        End If
        begineventprocessing = False
    End If
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (index)
End Sub

Sub HandleSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set Buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

Sub HandlePlayerVisibility(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    If Not Player(index).Visible = 0 Then
        Player(index).Visible = 0
    Else
        Player(index).Visible = 1
    End If
    
    Call SendPlayerData(index)
End Sub

' ::::::::::::::::::::::::
' :: Heal Player packet ::
' ::::::::::::::::::::::::
Sub HandleHealPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MONITOR Then Exit Sub

    ' The index
    n = FindPlayer(Buffer.ReadString)
    Set Buffer = Nothing

    ' Check if player is on
    If n > 0 Then
        Call SetPlayerVital(n, Vitals.HP, GetPlayerMaxVital(n, Vitals.HP))
        Call SetPlayerVital(n, Vitals.MP, GetPlayerMaxVital(n, Vitals.MP))
        Call SendVital(n, Vitals.HP)
        Call SendVital(n, Vitals.MP)
        Call PlayerMsg(n, "You have been healed by " & GetPlayerName(index) & "!", BrightBlue)
        Call AddLog(GetPlayerName(index) & " healed" & GetPlayerName(n) & ".", ADMIN_LOG)
    Else
        Call PlayerMsg(index, "Player is not online.", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Kill Player packet ::
' ::::::::::::::::::::::::
Sub HandleKillPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MONITOR Then Exit Sub

    ' The index
    n = FindPlayer(Buffer.ReadString)
    Set Buffer = Nothing

    ' Check if player is on
    If n > 0 Then
        Call SetPlayerVital(n, Vitals.HP, 0)
        Call SendVital(n, Vitals.HP)
        Call OnDeath(n)
        Call PlayerMsg(n, "You have been killed by " & GetPlayerName(index) & "!", BrightRed)
        Call AddLog(GetPlayerName(index) & " killed" & GetPlayerName(n) & ".", ADMIN_LOG)
    Else
        Call PlayerMsg(index, "Player is not online.", White)
    End If
End Sub

Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    
    n = Buffer.ReadLong 'CLng(Parse(1))
    
    If n < 0 Or n > MAX_QUESTS Then
    Exit Sub
    End If
    
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
SendQuests index
End Sub

Sub HandlePlayerHandleQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim QuestNum As Long, Order As Long, i As Long, n As Long
Dim RemoveStartItems As Boolean

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong
    Order = Buffer.ReadLong '1 = accept quest, 2 = cancel quest

    If Order = 1 Then
        RemoveStartItems = False
        'Alatar v1.2
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If FindOpenInvSlot(index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                    PlayerMsg index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                    RemoveStartItems = True
                    Exit For
                Else
                    If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        GiveInvItem index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            If FindOpenInvSlot(index, Quest(QuestNum).GiveItem(i).Item) = 0 Then
                                PlayerMsg index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                                RemoveStartItems = True
                                Exit For
                            Else
                                GiveInvItem index, Quest(QuestNum).GiveItem(i).Item, 1
                            End If
                        Next
                    End If
                End If
            End If
        Next

        If RemoveStartItems = False Then 'this means everything went ok
            Player(index).PlayerQuest(QuestNum).Status = QUEST_STARTED '1
            Player(index).PlayerQuest(QuestNum).ActualTask = 1
            Player(index).PlayerQuest(QuestNum).CurrentCount = 0
            PlayerMsg index, "New quest accepted: " & Trim$(Quest(QuestNum).Name) & "!", BrightGreen
        End If
    
    ElseIf Order = 2 Then
        Player(index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED '2
        Player(index).PlayerQuest(QuestNum).ActualTask = 1
        Player(index).PlayerQuest(QuestNum).CurrentCount = 0
        RemoveStartItems = True 'avoid exploits
        PlayerMsg index, Trim$(Quest(QuestNum).Name) & " has been canceled!", BrightGreen
    End If

    If RemoveStartItems = True Then
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If HasItem(index, Quest(QuestNum).GiveItem(i).Item) > 0 Then
                    If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            TakeInvItem index, Quest(QuestNum).GiveItem(i).Item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If


    SavePlayer index
    SendPlayerData index
    SendPlayerQuests index
    
    Set Buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests index
End Sub

Private Sub HandleNewClass(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Bid As Long
Dim num As Long
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    num = Buffer.ReadByte

    Set Buffer = Nothing
    If Class(num).startSpellCount > 0 Then
        For i = 1 To Class(num).startSpellCount
            If Class(num).StartSpell(i) > 0 Then
                If Len(Trim$(Spell(Class(num).StartSpell(i)).Name)) > 0 Then
                    If FindOpenSpellSlot(index) > 0 Then
                        If HasSpell(index, Class(num).StartSpell(i)) = False Then
                            SetPlayerSpell i, FindOpenSpellSlot(index), Class(num).StartSpell(i)
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    For i = 1 To Stats.Stat_Count - 1
        Player(index).stat(i) = Player(index).stat(i) + Class(num).stat(i)
    Next
    
    Player(index).Class = num
    
    
    Call SendPlayerSpells(index)
    Call SendPlayerData(index)
End Sub

Sub HandleRequestRank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

SendRankUpdate index

End Sub

Private Sub HandleClientOpenShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

Set Buffer = New clsBuffer
Buffer.WriteBytes Data()
i = Buffer.ReadLong
Call SendOpenShop(index, i)
TempPlayer(index).InShop = i
Set Buffer = Nothing

End Sub
