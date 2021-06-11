Attribute VB_Name = "modPlayer"
Option Explicit
Dim HealingAmount As Long

Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    Dim MyDate As String
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lbPlayers.List(index - 1) = GetPlayerIP(index) + " - " + GetPlayerLogin(index) + " - " + GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    
    If Player(index).Transformed = True Then
    SetPlayerSprite index, Player(index).Spriteold
          SetPlayerEquipment index, Player(index).EquipmentOld(1), Enchant
          SetPlayerEquipment index, Player(index).EquipmentOld(2), Helmet
          SetPlayerEquipment index, Player(index).EquipmentOld(3), Ring
          SetPlayerEquipment index, Player(index).EquipmentOld(4), weapon
          SetPlayerEquipment index, Player(index).EquipmentOld(5), Armor
          SetPlayerEquipment index, Player(index).EquipmentOld(6), Shield
          SetPlayerEquipment index, Player(index).EquipmentOld(7), Charm
          SetPlayerEquipment index, Player(index).EquipmentOld(8), Boots
          SetPlayerEquipment index, Player(index).EquipmentOld(9), Whetstone
    Player(index).Transformed = False
    End If
    
    Call SendQuests(index)
    Call SendClientTimeTo(index)
    
    Player(index).Visible = 0
    
    
    Call SetPlayerColorA(index, 255)
    Call SetPlayerColorR(index, 255)
    Call SetPlayerColorG(index, 255)
    Call SetPlayerColorB(index, 255)
    
    
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)
    
    'Do all the guild start up checks
    Call GuildLoginCheck(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next
    
    MyDate = Format(Date, "m/d/yyyy")

If Player(index).IsMember = 1 Then
        If DateDiff("d", Player(index).DateCount, MyDate) >= 31 Then
            PlayerMsg index, "Your membership has expired.", BrightRed
            MemberUnEquipItem index
            If Map(GetPlayerMap(index)).IsMember > 0 Then
                PlayerWarp index, Map(GetPlayerMap(index)).BootMap, Map(GetPlayerMap(index)).BootX, Map(GetPlayerMap(index)).BootY
            End If
            Player(index).IsMember = 0
            SavePlayer index
        Else
            PlayerMsg index, "You have " & (31 - DateDiff("d", Player(index).DateCount, MyDate)) & " days remaining of your membership!", Yellow
        End If
    End If

    
    ' Check Rank

For i = 1 To MAX_RANK

If Trim$(Rank(i).Name) = GetPlayerName(index) Then

Exit For

End If

If GetPlayerLevel(index) > Rank(i).Level Then

Rank(i).Name = GetPlayerName(index)

Rank(i).Level = GetPlayerLevel(index)

SaveRank

Exit For

End If

Next i
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
    
      Call SendNews(index)
   Call SendPatchNotes(index)
    
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False
        
        If Player(index).Transformed = True Then
          SetPlayerSprite index, Player(index).Spriteold
          SetPlayerEquipment index, Player(index).EquipmentOld(1), Enchant
          SetPlayerEquipment index, Player(index).EquipmentOld(2), Helmet
          SetPlayerEquipment index, Player(index).EquipmentOld(3), Ring
          SetPlayerEquipment index, Player(index).EquipmentOld(4), weapon
          SetPlayerEquipment index, Player(index).EquipmentOld(5), Armor
          SetPlayerEquipment index, Player(index).EquipmentOld(6), Shield
          SetPlayerEquipment index, Player(index).EquipmentOld(7), Charm
          SetPlayerEquipment index, Player(index).EquipmentOld(8), Boots
          SetPlayerEquipment index, Player(index).EquipmentOld(9), Whetstone
    End If

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave index
        
        If Player(index).GuildFileId > 0 Then
            'Set player online flag off
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Online = False
            Call CheckUnloadGuild(TempPlayer(index).tmpGuildSlot)
        End If

        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
Dim Armor As Long
Dim Helm As Long
Dim Whetstone As Long ' New
Dim Boots As Long ' New
Dim Charm As Long ' New
Dim Ring As Long ' New
Dim Enchant As Long ' New

GetPlayerProtection = 0

' Check for subscript out of range
If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
Exit Function
End If

Armor = GetPlayerEquipment(index, Armor)
Helm = GetPlayerEquipment(index, Helmet)
Whetstone = GetPlayerEquipment(index, Whetstone) ' New
Boots = GetPlayerEquipment(index, Boots) ' New
Charm = GetPlayerEquipment(index, Charm) ' New
Ring = GetPlayerEquipment(index, Ring) ' New
Enchant = GetPlayerEquipment(index, Enchant) ' New
GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

If Armor > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
End If

If Helm > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
End If
' New
If Whetstone > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Whetstone).Data2
End If

If Boots > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Boots).Data2
End If

If Charm > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Charm).Data2
End If

If Ring > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Ring).Data2
End If

If Enchant > 0 Then
GetPlayerProtection = GetPlayerProtection + Item(Enchant).Data2
End If
' /New
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If
    
    If Map(mapnum).IsMember > Player(index).IsMember Then
        PlayerMsg index, "This is a members only map and can not be visited without membership.", BrightRed
        Set Buffer = New clsBuffer
        Buffer.WriteLong SMapDone
        SendDataTo index, Buffer.ToArray()
        Set Buffer = Nothing
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
    If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
        Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
    End If
    
    TempPlayer(index).EventProcessingCount = 0
    TempPlayer(index).EventMap.CurrentEvents = 0
    
    ' clear target
    TempPlayer(index).Target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    
    ' Check to see if its a Party Dungeon
    If Map(mapnum).Moral = MAP_MORAL_PARTY_MAP Then
    
    ' Check to make sure the player is in a party. If not exit the sub so they dont change maps
    If TempPlayer(index).inParty < 1 Then
    Call PlayerMsg(index, "This is a party map. You must be in a party to enter it.", Red)
    Exit Sub
    End If
    
    End If

    If OldMap <> mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If
    
    UpdateMapBlock OldMap, GetPlayerX(index), GetPlayerY(index), False
    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    UpdateMapBlock mapnum, x, y, True
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).NPC(i).num > 0 Then
                MapNpc(OldMap).NPC(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).NPC(i).num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(index).GettingMap = YES
    Call CheckTasks(index, QUEST_TYPE_GOREACH, mapnum)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong mapnum
    Buffer.WriteLong Map(mapnum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False, Optional ByVal Sliding As Boolean = False, Optional ByVal SlideTime As Long)
    Dim Buffer As clsBuffer, mapnum As Long, i As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long, begineventprocessing As Boolean
    Dim FirstSlideTime As Long
    
    TempPlayer(index).isHealing = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = NO
    mapnum = GetPlayerMap(index)
    FirstSlideTime = GetTickCount + 300
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                                Call SetPlayerY(index, GetPlayerY(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(mapnum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                                Call SetPlayerY(index, GetPlayerY(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                                Call SetPlayerX(index, GetPlayerX(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(mapnum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                                Call SetPlayerX(index, GetPlayerX(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, mapnum, x, y)
            Moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
            Call PlayerWarp(index, mapnum, x, y)
            Moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And temptile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                temptile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                temptile(GetPlayerMap(index)).DoorTimer = GetTickCount
                SendMapKey index, x, y, 1
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            'VitalType = .Data1
            HealingAmount = .Data2
            
            Moved = YES
            
            TempPlayer(index).isHealing = True
        End If
        
        If .Type <> TILE_TYPE_HEAL Then
        TempPlayer(index).isHealing = False
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - Amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - Amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
        If Sliding = True Then
Slide:
            If GetTickCount > SlideTime Then
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Else
            GoTo Slide:
            End If
            'ForcePlayerMove index, MOVING_WALKING, .data1
        Else
Slide2:
            If GetTickCount > FirstSlideTime Then
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Else
            GoTo Slide2:
            End If
        End If
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If
    
    x = GetPlayerX(index)
    y = GetPlayerY(index)
    
    If Moved = YES Then
        If TempPlayer(index).EventMap.CurrentEvents > 0 Then
            For i = 1 To TempPlayer(index).EventMap.CurrentEvents
                If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Global = 1 Then
                    If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).x = x And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).y = y And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).Trigger = 1 And TempPlayer(index).EventMap.EventPages(i).Visible = 1 Then begineventprocessing = True
                Else
                    If TempPlayer(index).EventMap.EventPages(i).x = x And TempPlayer(index).EventMap.EventPages(i).y = y And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).Trigger = 1 And TempPlayer(index).EventMap.EventPages(i).Visible = 1 Then begineventprocessing = True
                End If
                If begineventprocessing = True Then
                    'Process this event, it is on-touch and everything checks out.
                    If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
                        TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
                        ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).ActionTimer = GetTickCount
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).CurList = 1
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).CurSlot = 1
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).eventID = TempPlayer(index).EventMap.EventPages(i).eventID
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).pageID = TempPlayer(index).EventMap.EventPages(i).pageID
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).WaitingForResponse = 0
                        ReDim TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount)
                    End If
                    begineventprocessing = False
                End If
            Next
        End If
    End If

End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 7 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    If Direction = DIR_UP Then
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_SLIDE Then
    PlayerMove index, Direction, movement, True, True, GetTickCount + 300
    Else
    PlayerMove index, Direction, movement, True
    End If
    End If
    
    If Direction = DIR_DOWN Then
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_SLIDE Then
    PlayerMove index, Direction, movement, True, True, GetTickCount + 300
    Else
    PlayerMove index, Direction, movement, True
    End If
    End If
    
    If Direction = DIR_LEFT Then
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_SLIDE Then
    PlayerMove index, Direction, movement, True, True, GetTickCount + 300
    Else
    PlayerMove index, Direction, movement, True
    End If
    End If
    
    If Direction = DIR_RIGHT Then
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_SLIDE Then
    PlayerMove index, Direction, movement, True, True, GetTickCount + 300
    Else
    PlayerMove index, Direction, movement, True
    End If
    End If
    
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(index, i)

        If itemnum > 0 Then

            Select Case i
                Case Equipment.weapon
                
                If Item(itemnum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor
                
                If Item(itemnum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                'NEW
                Case Equipment.Helmet
                
                If Item(itemnum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Whetstone
                
                If Item(itemnum).Type <> ITEM_TYPE_WHETSTONE Then SetPlayerEquipment index, 0, i
                
                Case Equipment.Boots
                
                If Item(itemnum).Type <> ITEM_TYPE_BOOTS Then SetPlayerEquipment index, 0, i
                Case Equipment.Charm
                
                If Item(itemnum).Type <> ITEM_TYPE_CHARM Then SetPlayerEquipment index, 0, i
                Case Equipment.Ring
                
                If Item(itemnum).Type <> ITEM_TYPE_RING Then SetPlayerEquipment index, 0, i
                Case Equipment.Enchant
                
                If Item(itemnum).Type <> ITEM_TYPE_ENCHANT Then SetPlayerEquipment index, 0, i
                ' /New
                Case Equipment.Shield
                
                If Item(itemnum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
                End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function FindItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            FindItem = i
            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemnum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemnum = GetPlayerInvItemNum(index, invSlot)

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendupdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, itemnum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemnum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        If sendupdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)

    For i = MAX_MAP_ITEMS To 1 Step -1
        ' See if theres even an item here
        If (MapItem(mapnum, i).num > 0) And (MapItem(mapnum, i).num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, i).num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, n)).Stackable > 0 Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                Msg = MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Call CheckTasks(index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(Item(GetPlayerInvItemNum(index, n)).Name)))
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long, Optional ByVal Dead As Boolean)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).num = GetPlayerInvItemNum(index, InvNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable > 0 Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                        If Dead = False Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Else
                        Call SendActionMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32))
                        End If
                        Call SetPlayerInvItemNum(index, InvNum, 0)
                        Call SetPlayerInvItemValue(index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = Amount
                        If Dead = True Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Else
                        Call SendActionMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32))
                        End If
                        Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    If Dead = True Then
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)) & ".", Yellow)
                    Else
                    Call SendActionMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)) & ".", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32))
                    End If
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).num, Amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    Dim RankPos As Byte
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 1)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
   ' LoadClasses
   ' For i = 1 To Max_Classes
    If Class(GetPlayerClass(index)).Next = 0 And Player(index).Level >= 25 Then
        SendNewClass index, GetPlayerClass(index), Class(GetPlayerClass(index)).Advancement(1), Class(GetPlayerClass(index)).Advancement(2), Class(GetPlayerClass(index)).Advancement(3)
        
    End If
    'Next
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP index
        SendPlayerData index
        
        ' check rank

RankPos = CheckRank(index)

If RankPos > 0 Then

ChangeRank index, RankPos

End If

        
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Sub SetPlayerSpriteOld(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Spriteold = Sprite
End Sub


Function GetPlayerLevel(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function
Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    Player(index).exp = exp
    If GetPlayerLevel(index) = MAX_LEVELS And Player(index).exp > GetPlayerNextLevel(index) Then
        Player(index).exp = GetPlayerNextLevel(index)
    End If
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVisible(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVisible = Player(index).Visible
End Function

Sub SetPlayerVisible(ByVal index As Long, ByVal Visible As Long)
    Player(index).Visible = Visible
End Sub

Function GetPlayerColorA(ByVal index As Long) As Long
Dim A2 As Byte

    If index > MAX_PLAYERS Then Exit Function
    
    
    GetPlayerColorA = Player(index).A
End Function

Sub SetPlayerColorA(ByVal index As Long, ByVal A As Long)
    Player(index).A = A
End Sub

Function GetPlayerColorR(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerColorR = Player(index).R
End Function

Sub SetPlayerColorR(ByVal index As Long, ByVal R As Long)
    Player(index).R = R
End Sub

Function GetPlayerColorG(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerColorG = Player(index).G
End Function

Sub SetPlayerColorG(ByVal index As Long, ByVal G As Long)
    Player(index).G = G
End Sub

Function GetPlayerColorB(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerColorB = Player(index).b
End Function

Sub SetPlayerColorB(ByVal index As Long, ByVal b As Long)
    Player(index).b = b
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
   
    x = Player(index).stat(stat)
   
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(stat) > 0 Then
                x = x + Item(Player(index).Equipment(i)).Add_Stat(stat)
            End If
        End If
    Next
   
    Select Case stat
        Case Stats.Strength
            For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_STR Then
                    x = x + TempPlayer(index).BuffValue(i)
                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_STR Then
                    x = x - TempPlayer(index).BuffValue(i)
                End If
            Next
        Case Stats.Endurance
            For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_END Then
                    x = x + TempPlayer(index).BuffValue(i)
                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_END Then
                    x = x - TempPlayer(index).BuffValue(i)
                End If
            Next
        Case Stats.Agility
            For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_AGI Then
                    x = x + TempPlayer(index).BuffValue(i)
                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_AGI Then
                    x = x - TempPlayer(index).BuffValue(i)
                End If
            Next
        Case Stats.Intelligence
            For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_INT Then
                    x = x + TempPlayer(index).BuffValue(i)
                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_INT Then
                    x = x - TempPlayer(index).BuffValue(i)
                End If
            Next
        Case Stats.Willpower
            For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_WILL Then
                    x = x + TempPlayer(index).BuffValue(i)
                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_WILL Then
                    x = x - TempPlayer(index).BuffValue(i)
                End If
            Next
    End Select
   
    GetPlayerStat = x
End Function


Public Function GetPlayerRawStat(ByVal index As Long, ByVal stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).stat(stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal stat As Stats, ByVal Value As Long)
    Player(index).stat(stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    Player(index).Inv(invSlot).num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellSlot As Long, ByVal SpellNum As Long)
    Player(index).Spell(spellSlot) = SpellNum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = InvNum
End Sub

Sub SetPlayerEquipmentOld(ByVal index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).EquipmentOld(EquipmentSlot) = InvNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.Target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    Dim ToolpowerReq As Long
    Dim Toolpower As Long
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(index, weapon) > 0 Then
                If Item(GetPlayerEquipment(index, weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                        
                        Damage = Item(GetPlayerEquipment(index, weapon)).Data2
                        
                        ' check if tool power is strong enough
                        If (Resource(Resource_index).ToolpowerReq <= Item(GetPlayerEquipment(index, weapon)).Toolpower) Then
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                                ' carry on
                                GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                GivePlayerEXP index, Resource(Resource_index).exp
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                            Call CheckTasks(index, QUEST_TYPE_GOTRAIN, Resource_index)
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                        
                        Else
                            'Tool power too low
                            SendActionMsg GetPlayerMap(index), "Tool too weak!", BrightRed, 1, (rX * 32), (rY * 32)
                            Call PlayerMsg(index, "Your " & Trim$(Item(GetPlayerEquipment(index, weapon)).Name) & " isn't strong enough to gather this.", White)
                            End If
                            
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                        End If
                    End If

                Else
                    PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Bank(index).Item(BankSlot).num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal Amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, invSlot)).Stackable > 0 Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(index, BankSlot)).Stackable > 0 Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        SendEXP index
        Call PlayerMsg(index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal InvNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, itemnum As Long, b As Long, j As Long
Dim TotalPoints As Long
Dim Result As Long
Dim MaxItems As Long

For j = 1 To MAX_INV
    Next
    
    b = FindOpenInvSlot(index, j)

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, InvNum)).Data2
        itemnum = GetPlayerInvItemNum(index, InvNum)
        
        If Player(index).Transformed = True Then
            If Item(itemnum).Type >= ITEM_TYPE_WEAPON And Item(itemnum).Type <= ITEM_TYPE_SHIELD Then
                PlayerMsg index, "Cant equip items while mount.", BrightRed
                Exit Sub
            End If
        End If
        
    If Item(itemnum).IsMember > 0 And Player(index).IsMember = 0 Then
        PlayerMsg index, "This is a members only item and can not be used without membership.", BrightRed
        Exit Sub
    End If
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemnum, Armor
                'PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If Item(itemnum).istwohander = True Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If GetPlayerEquipment(index, weapon) > 0 Then
                            If b < 1 Then
                                Call PlayerMsg(index, "You have no room in your inventory.", BrightRed)
                                Exit Sub
                                End If
                            End If
                        End If
                    End If

                If GetPlayerEquipment(index, weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, weapon)
                End If

                SetPlayerEquipment index, itemnum, weapon
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                If Item(itemnum).istwohander = True Then
                If GetPlayerEquipment(index, Shield) > 0 Then
                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                SetPlayerEquipment index, 0, Shield
                End If
                End If
                
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemnum, Helmet
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
                Case ITEM_TYPE_WHETSTONE
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                Exit Sub
                End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
               ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                If GetPlayerEquipment(index, Whetstone) > 0 Then
                tempItem = GetPlayerEquipment(index, Whetstone)
                End If
                
                SetPlayerEquipment index, itemnum, Whetstone
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1
                
                If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
                Case ITEM_TYPE_BOOTS
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                Exit Sub
                End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                If GetPlayerEquipment(index, Boots) > 0 Then
                tempItem = GetPlayerEquipment(index, Boots)
                End If
                
                SetPlayerEquipment index, itemnum, Boots
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1
                
                If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
                Case ITEM_TYPE_CHARM
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                Exit Sub
                End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                If GetPlayerEquipment(index, Charm) > 0 Then
                tempItem = GetPlayerEquipment(index, Charm)
                End If
                
                SetPlayerEquipment index, itemnum, Charm
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1
                
                If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
                Case ITEM_TYPE_RING
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                Exit Sub
                End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                If GetPlayerEquipment(index, Ring) > 0 Then
                tempItem = GetPlayerEquipment(index, Ring)
                End If
                
                SetPlayerEquipment index, itemnum, Ring
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1
                
                If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
                Case ITEM_TYPE_ENCHANT
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                Exit Sub
                End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                Exit Sub
                End If
                
                If GetPlayerEquipment(index, Enchant) > 0 Then
                tempItem = GetPlayerEquipment(index, Enchant)
                End If
                
                SetPlayerEquipment index, itemnum, Enchant
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1
                
                If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, itemnum, Shield
                SendActionMsg GetPlayerMap(index), "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                If GetPlayerEquipment(index, weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, weapon)).istwohander = True Then
                GiveInvItem index, GetPlayerEquipment(index, weapon), 0
                SetPlayerEquipment index, 0, weapon
                End If
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                Call SendStats(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                
                If Item(itemnum).addExpMultiplierTime > 0 Then
                Player(index).ExpMultiplierTime = Item(itemnum).addExpMultiplierTime
                Player(index).ExpMultiplier = Item(itemnum).addExpMultiplier
                SendActionMsg GetPlayerMap(index), "x" & Item(itemnum).addExpMultiplier & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
               End If
                
                Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(InvNum).num, 1)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If itemnum = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                        temptile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        temptile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        SendMapKey index, x, y, 1
                        Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, itemnum, 1)
                            Call PlayerMsg(index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            
            Case ITEM_TYPE_GIFT

                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                            TakeInvItem index, itemnum, 1
                            If Item(itemnum).multigift = True Then
                                For i = 1 To 10
                                    If Item(itemnum).gift(i) = 0 Then Exit For
                                    If Rnd * 100 <= Item(itemnum).gchance(i) Then
                                        If Item(itemnum).gift(i) > 0 Then
                                            If FindOpenInvSlot(index, Item(itemnum).gift(i)) = 0 Then
                                                Exit Sub
                                            End If
                                        End If
                                        GiveInvItem index, Item(itemnum).gift(i), Item(itemnum).gvalue(i)
                                    End If
                                Next
                            ElseIf Item(itemnum).multigift = False Then
                                For i = 1 To 10
                                    If Item(itemnum).gift(i) = 0 Then Exit For
                                    If Rnd * 100 <= Item(itemnum).gchance(i) Then
                                        If Item(itemnum).gift(i) > 0 Then
                                            If FindOpenInvSlot(index, Item(itemnum).gift(i)) = 0 Then
                                                Exit Sub
                                            End If
                                        End If
                                        GiveInvItem index, Item(itemnum).gift(i), Item(itemnum).gvalue(i)
                                        Exit For
                                    End If
                                Next
                            End If
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
            Case ITEM_TYPE_SPELL
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    'If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Call SetPlayerSpell(index, i, n)
                                    Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemnum, 1)
                                    Call PlayerMsg(index, "You feel imbued with knowledge. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    'Else
                       ' Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    'End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                
                Case ITEM_TYPE_STAT_RESET
                    TotalPoints = GetPlayerPOINTS(index)
                    
                    For i = 1 To Stats.Stat_Count - 1
                        TotalPoints = TotalPoints + GetPlayerRawStat(index, i)
                        Call SetPlayerStat(index, i, 1)
                    Next
                    Call SetPlayerPOINTS(index, TotalPoints - 5)
                    Call TakeInvItem(index, Player(index).Inv(InvNum).num, 1)
                    Call SendStats(index)
                    Call SendPlayerData(index)
                    Call PlayerMsg(index, "Your stats have been reset!", Yellow)
                    
                    Case ITEM_TYPE_RECIPE
                    
                    ' Check if on crafting tile
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_CRAFT Then
                    PlayerMsg index, "You cannot craft this item here. Travel to a more suitable location.", BrightRed
                    Exit Sub
                    End If
                    
                    ' Get the recipe information
                    Result = Item(GetPlayerInvItemNum(index, InvNum)).Data3
                    
                    ' Perform Recipe checks
                    If Result <= 0 Then Exit Sub
                    
                    If GetPlayerEquipment(index, weapon) <= 0 Then
                    Call PlayerMsg(index, "You must equip a tool!", White)
                    Exit Sub
                    End If
                    
                    ' Make sure proper weapon is equipped
                    If Item(GetPlayerInvItemNum(index, InvNum)).ToolReq <> 0 Then
                    If Item(GetPlayerEquipment(index, weapon)).Tool <> Item(GetPlayerInvItemNum(index, InvNum)).ToolReq Then
                    PlayerMsg index, "This isn't the tool used in the recipe.", BrightRed
                    Exit Sub
                    End If
                    End If
                    
                    If GetPlayerEquipment(index, weapon) <= 0 Then
                    Call PlayerMsg(index, "You must equip a tool!", White)
                    Exit Sub
                    End If
                    
                    
                    For i = 1 To MAX_RECIPE_ITEMS
                    If Item(itemnum).Recipe(i) > 0 Then
                    MaxItems = i
                    End If
                    Next
                    
                    ' Make sure there are at least two ingredients, and prevent subscript out of range
                    If MaxItems < 2 Or MaxItems > MAX_RECIPE_ITEMS Then
                    PlayerMsg index, "This is an incomplete recipe.", BrightRed
                    Exit Sub
                    End If
                    
                    ' Proceed to craft the item
                    CraftItem index, itemnum, MaxItems, Result
                    
                    Case ITEM_TYPE_TRANSFORM

               ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq(GetPlayerClass(index)) > 0 Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If Player(index).Transformed = True Then
                
                 SetPlayerSprite index, Player(index).Spriteold
                            SetPlayerEquipment index, Player(index).EquipmentOld(1), Enchant
                            SetPlayerEquipment index, Player(index).EquipmentOld(2), Helmet
                            SetPlayerEquipment index, Player(index).EquipmentOld(3), Ring
                            SetPlayerEquipment index, Player(index).EquipmentOld(4), weapon
                            SetPlayerEquipment index, Player(index).EquipmentOld(5), Armor
                            SetPlayerEquipment index, Player(index).EquipmentOld(6), Shield
                            SetPlayerEquipment index, Player(index).EquipmentOld(7), Charm
                            SetPlayerEquipment index, Player(index).EquipmentOld(8), Boots
                            SetPlayerEquipment index, Player(index).EquipmentOld(9), Whetstone
    
                            Player(index).Transformed = False
                            
                            SendActionMsg GetPlayerMap(index), "Melepas Mount!", BrightBlue, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

                            Call SendPlayerData(index)
                            Call SendWornEquipment(index)
                            Call SendMapEquipment(index)
                            Call SendVital(index, Vitals.HP)
                            Call SendVital(index, Vitals.MP)
                            ' send vitals to party if in one
                            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                            Exit Sub
                            End If

                
                If Player(index).Transformed = False Then
                            If Item(itemnum).Sprite > 0 Then
                                SetPlayerSpriteOld index, Player(index).Sprite
                                SetPlayerSprite index, Player(index).Spriteold
                                SetPlayerSprite index, Item(itemnum).Sprite
                                SetPlayerEquipmentOld index, Player(index).Equipment(1), Enchant
                                SetPlayerEquipmentOld index, Player(index).Equipment(2), Helmet
                                SetPlayerEquipmentOld index, Player(index).Equipment(3), Ring
                                SetPlayerEquipmentOld index, Player(index).Equipment(4), weapon
                                SetPlayerEquipmentOld index, Player(index).Equipment(5), Armor
                                SetPlayerEquipmentOld index, Player(index).Equipment(6), Shield
                                SetPlayerEquipmentOld index, Player(index).Equipment(7), Charm
                                SetPlayerEquipmentOld index, Player(index).Equipment(8), Boots
                                SetPlayerEquipmentOld index, Player(index).Equipment(9), Whetstone
                       
                                SetPlayerEquipment index, Player(index).EquipmentOld(1), Enchant
                                SetPlayerEquipment index, Player(index).EquipmentOld(2), Helmet
                                SetPlayerEquipment index, Player(index).EquipmentOld(3), Ring
                                SetPlayerEquipment index, Player(index).EquipmentOld(4), weapon
                                SetPlayerEquipment index, Player(index).EquipmentOld(5), Armor
                                SetPlayerEquipment index, Player(index).EquipmentOld(6), Shield
                                SetPlayerEquipment index, Player(index).EquipmentOld(7), Charm
                                SetPlayerEquipment index, Player(index).EquipmentOld(8), Boots
                                SetPlayerEquipment index, Player(index).EquipmentOld(9), Whetstone
    
                                SetPlayerEquipment index, 0, Enchant
                                SetPlayerEquipment index, 0, Helmet
                                SetPlayerEquipment index, 0, Ring
                                SetPlayerEquipment index, 0, weapon
                                SetPlayerEquipment index, 0, Armor
                                SetPlayerEquipment index, 0, Shield
                                SetPlayerEquipment index, 0, Charm
                                SetPlayerEquipment index, 0, Boots
                                SetPlayerEquipment index, 0, Whetstone
                                
                       
                                Player(index).Transformed = True
                                
                SendActionMsg GetPlayerMap(index), "Menaiki Mount!", BrightBlue, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
      
                                Call SendWornEquipment(index)
                                Call SendMapEquipment(index)
                                Call SendVital(index, Vitals.HP)
                                ' send vitals to party if in one
                                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                                Call SendPlayerData(index)
                                
                                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
                                
                            End If
                            Else
                            PlayerMsg index, "You can not mount this player!", BrightRed
                            End If
                    
        End Select
    End If
End Sub

Function GetPlayerSpellUses(ByVal index As Long, ByVal spellSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpellUses = Player(index).SpellUses(spellSlot)
End Function

Sub SetPlayerSpellUses(ByVal index As Long, ByVal spellSlot As Long, ByVal SpellNum As Long)
    Player(index).SpellUses(spellSlot) = SpellNum
End Sub

Sub CraftItem(ByVal index As Long, ByVal itemnum As Long, ByVal MaxItems As Long, ByVal Result As Long)
Dim i As Long
Dim CrftItemAmount As Long
' See what items are needed, and how many
' If requirements are met, craft the item, if not, show failure message
Select Case MaxItems

Case 1
    PlayerMsg index, "This is an incomplete recipe.", BrightRed
Exit Sub

Case 2
If HasItem(index, Item(itemnum).Recipe(1)) And HasItem(index, Item(itemnum).Recipe(2)) Then

    If Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) Then

        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 2 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        End If

    Else

    TakeInvItem index, Item(itemnum).Recipe(1), 1
    TakeInvItem index, Item(itemnum).Recipe(2), 1
    End If

GoTo MakeItem

Else

PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
Exit Sub
End If

Case 3
If HasItem(index, Item(itemnum).Recipe(1)) And HasItem(index, Item(itemnum).Recipe(2)) And HasItem(index, Item(itemnum).Recipe(3)) Then

'All 3 are the same
If Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(3) Then


        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 3 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        End If

ElseIf Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) Then

'1 different

CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 2 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        End If
Else

TakeInvItem index, Item(itemnum).Recipe(1), 1
TakeInvItem index, Item(itemnum).Recipe(2), 1
TakeInvItem index, Item(itemnum).Recipe(3), 1

End If


GoTo MakeItem
Else
PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
Exit Sub
End If

Case 4
If HasItem(index, Item(itemnum).Recipe(1)) And HasItem(index, Item(itemnum).Recipe(2)) And HasItem(index, Item(itemnum).Recipe(3)) And HasItem(index, Item(itemnum).Recipe(4)) Then

'All 4 are the same
If Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(3) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(4) Then


        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 4 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        End If

ElseIf Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(3) Then

'3 are the same


        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 3 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        End If

ElseIf Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) Then

'2 are the same

CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 2 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        End If
Else


TakeInvItem index, Item(itemnum).Recipe(1), 1
TakeInvItem index, Item(itemnum).Recipe(2), 1
TakeInvItem index, Item(itemnum).Recipe(3), 1
TakeInvItem index, Item(itemnum).Recipe(4), 1
End If

GoTo MakeItem
Else
PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
Exit Sub
End If


Case 5
If HasItem(index, Item(itemnum).Recipe(1)) And HasItem(index, Item(itemnum).Recipe(2)) And HasItem(index, Item(itemnum).Recipe(3)) And HasItem(index, Item(itemnum).Recipe(4)) And HasItem(index, Item(itemnum).Recipe(5)) Then

'All 5 are the same
If Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(3) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(4) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(5) Then


        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 5 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        TakeInvItem index, Item(itemnum).Recipe(5), 1
        End If

ElseIf Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(3) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(4) Then


'4 are the same

        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 4 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        TakeInvItem index, Item(itemnum).Recipe(5), 1
        End If

ElseIf Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) And Item(itemnum).Recipe(1) = Item(itemnum).Recipe(3) Then

'3 are the same


        CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 3 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        TakeInvItem index, Item(itemnum).Recipe(5), 1
        End If

ElseIf Item(itemnum).Recipe(1) = Item(itemnum).Recipe(2) Then

'2 are the same

CrftItemAmount = 0
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Item(itemnum).Recipe(1) Then
                    CrftItemAmount = CrftItemAmount + 1
            End If
        Next

        If CrftItemAmount = 0 Or CrftItemAmount < 2 Then
                    PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
        Exit Sub
        Else
        TakeInvItem index, Item(itemnum).Recipe(1), 1
        TakeInvItem index, Item(itemnum).Recipe(2), 1
        TakeInvItem index, Item(itemnum).Recipe(3), 1
        TakeInvItem index, Item(itemnum).Recipe(4), 1
        TakeInvItem index, Item(itemnum).Recipe(5), 1
        End If
Else


TakeInvItem index, Item(itemnum).Recipe(1), 1
TakeInvItem index, Item(itemnum).Recipe(2), 1
TakeInvItem index, Item(itemnum).Recipe(3), 1
TakeInvItem index, Item(itemnum).Recipe(4), 1
TakeInvItem index, Item(itemnum).Recipe(5), 1

End If
PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
Exit Sub
Else
PlayerMsg index, "You do not have all of the ingredients needed in this recipe.", BrightRed
Exit Sub
End If
End Select

MakeItem:

GiveInvItem index, Result, 1
If itemnum = Item(itemnum).Recipe(1) Then

Else
TakeInvItem index, itemnum, 1
End If
PlayerMsg index, "You have successfully created a " & Trim(Item(Result).Name) & ".", White
End Sub

Public Function isHealing(ByVal index As Long)

Dim mapnum As Long

mapnum = GetPlayerMap(index)

                If Not GetPlayerVital(index, Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP) Then
                SendActionMsg mapnum, "+" & HealingAmount + ((GetPlayerLevel(index) * 5)), BrightGreen, 1, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SetPlayerVital index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + HealingAmount + ((GetPlayerLevel(index) * 5))
                End If
                
                If Not GetPlayerVital(index, Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP) Then
                SendActionMsg mapnum, "+" & HealingAmount + ((GetPlayerLevel(index) * 3)), BrightBlue, 1, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 16
                SetPlayerVital index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + HealingAmount + ((GetPlayerLevel(index) * 3))
                End If
                
                'PlayerMsg index, "You feel replenished.", BrightGreen
                
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
End Function

Public Sub ApplyBuff(ByVal index As Long, ByVal BuffType As Long, ByVal Duration As Long, ByVal Amount As Long)
    Dim i As Long
   
    For i = 1 To 10
        If TempPlayer(index).Buffs(i) = 0 Then
            TempPlayer(index).Buffs(i) = BuffType
            TempPlayer(index).BuffTimer(i) = Duration
            TempPlayer(index).BuffValue(i) = Amount
            Exit For
        End If
    Next
   
    If BuffType = BUFF_ADD_HP Then
        Call SetPlayerVital(index, HP, GetPlayerVital(index, Vitals.HP) + Amount)
    End If
    If BuffType = BUFF_ADD_MP Then
        Call SetPlayerVital(index, MP, GetPlayerVital(index, Vitals.MP) + Amount)
    End If
   
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
   
End Sub

Private Function CheckRank(ByVal index As Long) As Byte

Dim i As Byte

For i = 1 To MAX_RANK

If GetPlayerLevel(index) > Rank(i).Level Then

CheckRank = i

Exit Function

End If

Next i

End Function



Private Sub ChangeRank(ByVal index As Long, RankPos As Byte)

Dim i As Long, ClearPos As Byte



' if not change position in rank

If GetPlayerName(index) = Trim$(Rank(RankPos).Name) Then

Rank(RankPos).Level = GetPlayerLevel(index)

SaveRank

Exit Sub

End If



' search player in rank

For i = 1 To MAX_RANK

If GetPlayerName(index) = Trim$(Rank(i).Name) Then

Rank(i).Name = vbNullString

Rank(i).Level = 0

ClearPos = i

Exit For

End If

Next i



' down clear position

If ClearPos > 0 Then

For i = ClearPos To MAX_RANK

If i = MAX_RANK Then

Rank(i).Name = vbNullString

Rank(i).Level = 0

Else

Rank(i).Name = Rank(i + 1).Name

Rank(i).Level = Rank(i + 1).Level

End If

Next i

End If



' open space in rank to player

For i = MAX_RANK To RankPos Step -1

If i > RankPos Then

Rank(i).Name = Rank(i - 1).Name

Rank(i).Level = Rank(i - 1).Level

End If

Next i



' put player in rank

Rank(RankPos).Name = GetPlayerName(index)

Rank(RankPos).Level = GetPlayerLevel(index)



SaveRank

End Sub

Public Sub MemberUnEquipItem(ByVal index As Long)
Dim recordbankslot As Long
Dim i As Long

    PlayerMsg index, "All 'Members only' items that are currently equipped will either go in your inventory or bank, if your inventory and bank is full, you'll lose the items. If you were in an 'Members only' area then you have been auto-warped to The Garden.", BrightRed

    For i = 1 To Equipment.Equipment_Count - 1

                If GetPlayerEquipment(index, i) > 0 Then
                        If Item(GetPlayerEquipment(index, i)).IsMember > 0 Then
                                If FindOpenInvSlot(index, GetPlayerEquipment(index, i)) > 0 Then
                                PlayerUnequipItem index, i
                                    ElseIf FindOpenBankSlot(index, GetPlayerEquipment(index, i)) > 0 Then
                                    recordbankslot = FindOpenBankSlot(index, GetPlayerEquipment(index, i))
                                    SetPlayerBankItemNum index, recordbankslot, GetPlayerEquipment(index, i)
                                    SetPlayerBankItemValue index, recordbankslot, GetPlayerBankItemValue(index, i) + 1
                                    SetPlayerEquipment index, 0, i
                                    Else
                                    SetPlayerEquipment index, 0, i
                                End If
                        End If
                End If
    Next i

        SendWornEquipment index
        SendMapEquipment index
        SavePlayer index
        SendPlayerData index
End Sub

