Attribute VB_Name = "modCombat"
Option Explicit
Dim i As Long

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(index)
                Case 1 ' Page ' 2087 at 255 Endurance
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 100
                Case 2 ' Apprentice ' 1015 at 255 Endurance
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 7 + 88
                Case 3 ' Thief ' 1417 at 255 Endurance
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 10 + 92
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
            End Select
            
             For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_HP Then
                    GetPlayerMaxVital = GetPlayerMaxVital + TempPlayer(index).BuffValue(i)

                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_HP Then
                    GetPlayerMaxVital = GetPlayerMaxVital - TempPlayer(index).BuffValue(i)
                End If
            Next
            
        Case MP
            Select Case GetPlayerClass(index)
                Case 1 ' Page ' 700 at 255 Intelligence
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 37
                Case 2 ' Apprentice ' 1425 at 255 Intelligence
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 10 + 100
                Case 3 ' Thief ' 835  at 255 Intelligence
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 6 + 40
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
            End Select
            
            For i = 1 To 10
                If TempPlayer(index).Buffs(i) = BUFF_ADD_MP Then
                    GetPlayerMaxVital = GetPlayerMaxVital + TempPlayer(index).BuffValue(i)
                End If
                If TempPlayer(index).Buffs(i) = BUFF_SUB_MP Then
                    GetPlayerMaxVital = GetPlayerMaxVital - TempPlayer(index).BuffValue(i)
                End If
            Next
            
    End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    Dim i As Long
    
    For i = 1 To 10
        If TempPlayer(index).Buffs(i) = BUFF_ADD_ATK Then
            GetPlayerDamage = GetPlayerDamage + TempPlayer(index).BuffValue(i)
        End If
        If TempPlayer(index).Buffs(i) = BUFF_SUB_ATK Then
            GetPlayerDamage = GetPlayerDamage - TempPlayer(index).BuffValue(i)
        End If
    Next
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) + (GetPlayerLevel(index) / 5)
    End If

End Function

Function GetPlayerDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim Def As Long
    Dim i As Long
    
    For i = 1 To 10
        If TempPlayer(index).Buffs(i) = BUFF_ADD_DEF Then
            GetPlayerDef = GetPlayerDef + TempPlayer(index).BuffValue(i)
        End If
        If TempPlayer(index).Buffs(i) = BUFF_SUB_DEF Then
            GetPlayerDef = GetPlayerDef - TempPlayer(index).BuffValue(i)
        End If
    Next
    
    GetPlayerDef = 0
    Def = 0
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, Helmet)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Whetstone) > 0 Then
        DefNum = GetPlayerEquipment(index, Whetstone)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Enchant) > 0 Then
        DefNum = GetPlayerEquipment(index, Enchant)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Ring) > 0 Then
        DefNum = GetPlayerEquipment(index, Ring)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Boots) > 0 Then
        DefNum = GetPlayerEquipment(index, Boots)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Charm) > 0 Then
        DefNum = GetPlayerEquipment(index, Charm)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        Def = Def + Item(DefNum).Data2
    End If
    
  If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, Helmet) > 0 And Not GetPlayerEquipment(index, Shield) > 0 Then
        GetPlayerDef = 0.085 * GetPlayerStat(index, Endurance) + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDef = 0.085 * GetPlayerStat(index, Endurance) * Def + (GetPlayerLevel(index) / 5)
    End If
    

End Function

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(NPCNum).HP
        Case MP
            GetNpcMaxVital = 30 + (NPC(NPCNum).stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (NPC(NPCNum).stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (NPC(NPCNum).stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal NPCNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * NPC(NPCNum).stat(Stats.Strength) * NPC(NPCNum).Damage + (NPC(NPCNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(index, Agility) / 52.08
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(index, Agility) / 83.3
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False
    
    If GetPlayerStat(index, Strength) >= 200 Then
    rate = GetPlayerStat(index, Strength) * 0.001 + 50
    Else
    rate = GetPlayerStat(index, Strength) * 0.25
    End If
    rndNum = rand(1, 100)

    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim stat As Long
Dim rndNum As Long

    CanNpcBlock = False
    
    stat = NPC(NPCNum).stat(Stats.Agility) / 5  'guessed shield agility
    rate = stat / 12.08
    
    rndNum = rand(1, 100)
    
    If rndNum <= rate Then
        CanNpcBlock = True
    End If
    
End Function

Public Function CanNpcCrit(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = NPC(NPCNum).stat(Stats.Agility) / 52.08
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = NPC(NPCNum).stat(Stats.Agility) / 83.3
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False
    
    If NPC(NPCNum).stat(Stats.Strength) >= 200 Then
    rate = NPC(NPCNum).stat(Stats.Strength) * 0.001 + 50
    Else
    rate = NPC(NPCNum).stat(Stats.Strength) * 0.25
    End If
    'If NPC(NPCNum).Stat(Stats.strength) >= 1000 Then
    'rndNum = rand(1, 1000)
    'Else
    rndNum = rand(1, 100)
    'End If
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNpcNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, MapNpcNum) Then
    
        mapnum = GetPlayerMap(index)
        NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        

If GetPlayerEquipment(index, weapon) <> 0 Then
If Item(Player(index).Equipment(Equipment.weapon)).Element <> 0 Then
        
        
        'Weapon Element Fire
If Item(Player(index).Equipment(Equipment.weapon)).Element = 1 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage / 2
        SendActionMsg mapnum, "Not very Effective...", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 2
        SendActionMsg mapnum, "Very Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End If
        
'Weapon Element Water
If Item(Player(index).Equipment(Equipment.weapon)).Element = 2 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage * 2
        SendActionMsg mapnum, "Very Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If

'Weapon Element Wind
'Give Wind element weapons more attack speed
If Item(Player(index).Equipment(Equipment.weapon)).Element = 3 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage / 2
        SendActionMsg mapnum, "Not Very Effective...", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If

'Weapon Element Earth
If Item(Player(index).Equipment(Equipment.weapon)).Element = 4 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If

'Weapon Element Light
If Item(Player(index).Equipment(Equipment.weapon)).Element = 5 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage * 2
        SendActionMsg mapnum, "Very Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End If

'Weapon Element Dark
If Item(Player(index).Equipment(Equipment.weapon)).Element = 6 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage / 2
        SendActionMsg mapnum, "Not very Effective...", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If
End If

End If
        'End If
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (NPC(NPCNum).stat(Stats.Endurance) * 2))
        
        ' randomise from 1 to max hit IF it's two-handed then give it higher odds
        If GetPlayerEquipment(index, weapon) > 0 Then
        If Item(Player(index).Equipment(Equipment.weapon)).istwohander = True Then
        Damage = rand(Damage / 2.3, Damage)
        Else
        Damage = rand(1, Damage)
        End If
        End If
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
        
        'Stealth Bonus!
        If TempPlayer(index).StealthDuration > 0 Then
        Damage = Damage * 3
        SendActionMsg mapnum, "Stealth hit!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        TempPlayer(index).StealthDuration = 0
        TempPlayer(index).StealthTimer = 0
        Player(index).Visible = 0
        Call SetPlayerColorA(index, 255)
        
        SendStealthed index
        End If
        
        If Damage > 0 Then
            Call PlayerAttackNpc(index, MapNpcNum, Damage)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim AttackSpeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).NPC(MapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(attacker)
    NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If NPCNum > 0 Then
                If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(attacker).Target = MapNpcNum
                    SendTarget attacker
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, weapon) > 0 Then
            AttackSpeed = Item(GetPlayerEquipment(attacker, weapon)).speed
        Else
            AttackSpeed = 1000
        End If

        If NPCNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + AttackSpeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x + 1
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x - 1
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y
            End Select

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(attacker).Target = MapNpcNum
                    SendTarget attacker
                        CanPlayerAttackNpc = True
                    Else
                    If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                    Call CheckTasks(attacker, QUEST_TYPE_GOTALK, NPCNum)
                    Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, NPCNum)
                    Call CheckTasks(attacker, QUEST_TYPE_GOGET, NPCNum)
                    
                    
                    
                    If NPC(NPCNum).Quest = YES Then
                    
                    With MapNpc(mapnum).NPC(MapNpcNum)
                                    'MapNpc(mapnum).NPC(npcNum).inChatWith = attacker
                                    'MapNpc(mapnum).NPC(npcNum).lastDir = .Dir
                                    If GetPlayerY(attacker) = .y - 1 Then
                                        .Dir = DIR_UP
                                    ElseIf GetPlayerY(attacker) = .y + 1 Then
                                        .Dir = DIR_DOWN
                                    ElseIf GetPlayerX(attacker) = .x - 1 Then
                                        .Dir = DIR_LEFT
                                    ElseIf GetPlayerX(attacker) = .x + 1 Then
                                        .Dir = DIR_RIGHT
                                    End If
                                    ' send NPC's dir to the map
                                    NpcDir mapnum, NPCNum, .Dir
                                End With
                                
                    If Player(attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED Then
                    If Quest(NPC(NPCNum).Quest).Repeat = YES Then
                    Player(attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED_BUT
                    Exit Function
                    End If
                    End If
                    If CanStartQuest(attacker, NPC(NPCNum).QuestNum) Then
                    'if can start show the request message (speech1)
                    QuestMessage attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(1)), NPC(NPCNum).QuestNum
                    Exit Function
                    End If
                    If QuestInProgress(attacker, NPC(NPCNum).QuestNum) Then
                    'if the quest is in progress show the meanwhile message (speech2)
                    QuestMessage attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(2)), 0
                    Exit Function
                    End If
                    End If
                    End If
                        If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                            Call SendChatBubble(mapnum, MapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim Def As Long
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
    Name = Trim$(NPC(NPCNum).Name)
    
     'Check for blind :D
        If TempPlayer(attacker).BlindDuration > 0 Then
            Damage = 0
            SendActionMsg mapnum, "Miss!", BrightRed, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        End If
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, weapon) > 0 Then
        n = GetPlayerEquipment(attacker, weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
    
    If Damage > 0 Then

    ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y
    
    ' send animation
    If n > 0 Then
        If Not overTime Then
            If SpellNum = 0 Then
                Call SendAnimation(mapnum, Item(n).Animation, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y)
                SendMapSound attacker, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, SoundEntity.seItem, n
            End If
        End If
    End If
    
    If SpellNum > 0 Then
        Call SendAnimation(mapnum, Spell(SpellNum).SpellAnim, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y)
        SendMapSound attacker, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, SoundEntity.seSpell, SpellNum
    End If
    
    Call SendFlash(MapNpcNum, mapnum, True)
    
    If Damage >= MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) Then
        
         
        
        ' make sure exp mod is not 0.
        If (frmServer.scrlExp.Value <> 0) Then
            ' Calculate exp to give attacker
            exp = NPC(NPCNum).exp * frmServer.scrlExp.Value
                        
                        
                        ' Check if npc level is bigger then 0

     If NPC(NPCNum).Level > 0 Then

         ' exp deduction

         If NPC(NPCNum).Level <= GetPlayerLevel(attacker) - 10 Then

             ' 10 levels lower, exp 0

             Call PlayerMsg(attacker, "Musuhmu 10 level di bawahmu, kamu tidak akan mendapatkan exp.", BrightRed) ' IF you want you can delete this msg cuz its only informating player

             exp = 0

         ElseIf NPC(NPCNum).Level <= GetPlayerLevel(attacker) - 5 Then

             ' half exp if enemy is 3 levels lower

             Call PlayerMsg(attacker, "Musuhmu 5 level di bawahmu, kamu mendapatkan setengah exp.", BrightRed) ' IF you want you can delete this msg cuz its only informating player

             exp = exp / 2

         ElseIf NPC(NPCNum).Level >= GetPlayerLevel(attacker) + 10 Then

         End If

     End If
                
                        
                        
                            ' Double exp?
        If DoubleExp Then exp = exp * 2
    
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
    
            ' in party?
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party sharing function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerMap(attacker)
            Else
                ' no party - keep exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        
        'Drop the goods if they get it
            For n = 1 To MAX_NPC_DROPS
            If NPC(NPCNum).DropItem(n) = 0 Then Exit For
            
            If Rnd <= NPC(NPCNum).DropChance(n) Then
            Call SpawnItem(NPC(NPCNum).DropItem(n), NPC(NPCNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y)
            End If
            Next
            
            If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).isBoss = YES Then
            SendBossMsg Trim$(NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
            GlobalMsg Trim$(NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(MapNpcNum).FirstCast = False
        MapNpc(mapnum).NPC(MapNpcNum).num = 0
        MapNpc(mapnum).NPC(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = 0
        UpdateMapBlock mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, False
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(MapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(MapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, NPCNum)
        
        ' send death to the map
        Set buffer = New clsBuffer
        buffer.WriteLong SNpcDead
        buffer.WriteLong MapNpcNum
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = MapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) - Damage

        ' Set the NPC target to the player
        MapNpc(mapnum).NPC(MapNpcNum).targetType = 1 ' player
        MapNpc(mapnum).NPC(MapNpcNum).Target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(MapNpcNum).num Then
                    MapNpc(mapnum).NPC(i).Target = attacker
                    MapNpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).NPC(MapNpcNum).stopRegen = True
        MapNpc(mapnum).NPC(MapNpcNum).stopRegenTimer = GetTickCount
        
        End If
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNpcNum, mapnum, SpellNum
            'Blind
            If Spell(SpellNum).BlindDuration > 0 Then BlindNPC MapNpcNum, mapnum, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc mapnum, MapNpcNum, SpellNum, attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, MapNpcNum
End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
Dim mapnum As Long, NPCNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNpcNum, index) Then
        mapnum = GetPlayerMap(index)
        NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armor
        Damage = Damage - rand(1, (GetPlayerStat(index, Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        End If
        
        ' Calculate Defense
        Damage = Damage - GetPlayerDef(index)
        
        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
    NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).AttackSpeed > 0 Then
    If GetTickCount < MapNpc(mapnum).NPC(MapNpcNum).AttackTimer + NPC(MapNpc(mapnum).NPC(MapNpcNum).num).AttackSpeed Then
    Exit Function
    End If
    Else
    If GetTickCount < MapNpc(mapnum).NPC(MapNpcNum).AttackTimer + 1000 Then
    Exit Function
    End If
    End If
    
    'Check if the NPC isn't stunned.
    If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).StunDuration > 0 Then
        CanNpcAttackPlayer = False
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapnum).NPC(MapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NPCNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim exp As Long
    Dim mapnum As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).NPC(MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong MapNpcNum
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    'Check for blind :D
        If MapNpc(mapnum).NPC(MapNpcNum).BlindDuration > 0 Then
            Damage = 0
            SendActionMsg mapnum, "Miss!", BrightRed, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        End If


    If Damage > 0 Then
    
    ' set the regen timer
    MapNpc(mapnum).NPC(MapNpcNum).stopRegen = True
    MapNpc(mapnum).NPC(MapNpcNum).stopRegenTimer = GetTickCount
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(victim)).NPC(MapNpcNum).num).Animation, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim)
    ' send the sound
    SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(MapNpcNum).num
    
    Call SendFlash(victim, mapnum, False)
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(MapNpcNum).Target = 0
        MapNpc(mapnum).NPC(MapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If
End If
End Sub

Sub NpcSpellPlayer(ByVal MapNpcNum As Long, ByVal victim As Long, SpellSlotNum As Long, FCast As Boolean)
Dim mapnum As Long
Dim i As Long
Dim n As Long
Dim SpellNum As Long
Dim buffer As clsBuffer
Dim InitDamage As Long
Dim Damage As Long
Dim MaxHeals As Long

' Check for subscript out of range
If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
Exit Sub
End If

' Check for subscript out of range
If MapNpc(GetPlayerMap(victim)).NPC(MapNpcNum).num <= 0 Then
Exit Sub
End If

If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub

' The Variables
mapnum = GetPlayerMap(victim)
SpellNum = NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Spell(SpellSlotNum)

' Send this packet so they can see the person attacking
Set buffer = New clsBuffer
buffer.WriteLong SNpcAttack
buffer.WriteLong MapNpcNum
SendDataToMap mapnum, buffer.ToArray()
Set buffer = Nothing

If FCast = False Then
MapNpc(mapnum).NPC(MapNpcNum).SpellTimer = GetTickCount + 3000
MapNpc(mapnum).NPC(MapNpcNum).FirstCast = True
End If

' CoolDown Time
If MapNpc(mapnum).NPC(MapNpcNum).SpellTimer > GetTickCount Then Exit Sub

' Spell Types
Select Case Spell(SpellNum).Type
' AOE Healing Spells
Case SPELL_TYPE_HEALHP
' Make sure an npc waits for the spell to cooldown
MaxHeals = 1 + NPC(MapNpc(mapnum).NPC(MapNpcNum).num).stat(Stats.Intelligence) \ 25
If MapNpc(mapnum).NPC(MapNpcNum).Heals >= MaxHeals Then Exit Sub
If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) <= NPC(MapNpc(mapnum).NPC(MapNpcNum).num).HP * 0.3 Then
If Spell(SpellNum).IsAoE Then
For i = 1 To MAX_MAP_NPCS
If MapNpc(mapnum).NPC(i).num > 0 Then
If MapNpc(mapnum).NPC(i).Vital(Vitals.HP) > 0 Then
If isInRange(Spell(SpellNum).AoE, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
InitDamage = Spell(SpellNum).Vital + (NPC(MapNpc(mapnum).NPC(MapNpcNum).num).stat(Stats.Intelligence) * 2)

MapNpc(mapnum).NPC(i).Vital(Vitals.HP) = MapNpc(mapnum).NPC(i).Vital(Vitals.HP) + InitDamage
SendActionMsg mapnum, "+" & InitDamage, BrightGreen, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32)
Call SendAnimation(mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, MapNpcNum)

If MapNpc(mapnum).NPC(i).Vital(Vitals.HP) > NPC(MapNpc(mapnum).NPC(i).num).HP Then
MapNpc(mapnum).NPC(i).Vital(Vitals.HP) = NPC(MapNpc(mapnum).NPC(i).num).HP
End If

MapNpc(mapnum).NPC(MapNpcNum).Heals = MapNpc(mapnum).NPC(MapNpcNum).Heals + 1

MapNpc(mapnum).NPC(MapNpcNum).SpellTimer = GetTickCount + Spell(SpellNum).CDTime * 1000
Exit Sub
End If
End If
End If
Next
Else
' Non AOE Healing Spells
InitDamage = Spell(SpellNum).Vital + (NPC(MapNpc(mapnum).NPC(MapNpcNum).num).stat(Stats.Intelligence) * 2)

MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) + InitDamage
SendActionMsg mapnum, "+" & InitDamage, BrightGreen, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
Call SendAnimation(mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, MapNpcNum)

If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) > NPC(MapNpc(mapnum).NPC(MapNpcNum).num).HP Then
MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = NPC(MapNpc(mapnum).NPC(MapNpcNum).num).HP
End If

MapNpc(mapnum).NPC(MapNpcNum).Heals = MapNpc(mapnum).NPC(MapNpcNum).Heals + 1

MapNpc(mapnum).NPC(MapNpcNum).SpellTimer = GetTickCount + Spell(SpellNum).CDTime * 1000
Exit Sub
End If
End If

' AOE Damaging Spells
Case SPELL_TYPE_DAMAGEHP
' Make sure an npc waits for the spell to cooldown
If Spell(SpellNum).IsAoE Then
For i = 1 To Player_HighIndex
If IsPlaying(i) Then
If GetPlayerMap(i) = mapnum Then
If isInRange(Spell(SpellNum).AoE, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, GetPlayerX(i), GetPlayerY(i)) Then
InitDamage = Spell(SpellNum).Vital + (NPC(MapNpc(mapnum).NPC(MapNpcNum).num).stat(Stats.Intelligence) * 2)
Damage = InitDamage - Player(i).stat(Stats.Willpower)

If Damage <= 0 Then
SendActionMsg GetPlayerMap(i), "Resist!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
Exit Sub
Else
NpcAttackPlayer MapNpcNum, i, Damage
'if a stunning spell, stun the player
If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
'Blind
If Spell(SpellNum).BlindDuration > 0 Then BlindPlayer victim, SpellNum

SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, MapNpcNum
MapNpc(mapnum).NPC(MapNpcNum).SpellTimer = GetTickCount + Spell(SpellNum).CDTime * 1000
Exit Sub
End If
End If
End If
End If
Next
' Non AoE Damaging Spells
Else
If isInRange(Spell(SpellNum).Range, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, GetPlayerX(victim), GetPlayerY(victim)) Then
InitDamage = Spell(SpellNum).Vital + (NPC(MapNpc(mapnum).NPC(MapNpcNum).num).stat(Stats.Intelligence) * 2)
Damage = InitDamage - Player(victim).stat(Stats.Willpower)

If Damage <= 0 Then
SendActionMsg GetPlayerMap(victim), "Resist!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
Exit Sub
Else
NpcAttackPlayer MapNpcNum, victim, Damage
'if a stunning spell, stun the player
If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum

'Blind
If Spell(SpellNum).BlindDuration > 0 Then BlindPlayer victim, SpellNum

SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, victim
MapNpc(mapnum).NPC(MapNpcNum).SpellTimer = GetTickCount + Spell(SpellNum).CDTime * 1000
Exit Sub
End If
End If
End If
End Select
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        mapnum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(victim, Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If
        
        ' Calculate Defense
        Damage = Damage - GetPlayerDef(victim)

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, weapon)).speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
     If GetPlayerLevel(attacker) > TempPlayer(victim).inParty Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
     End If
        
    TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).Target = victim
    SendTarget attacker

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, weapon) > 0 Then
        n = GetPlayerEquipment(attacker, weapon)
    End If
    
    'Check for blind :D
        If TempPlayer(attacker).BlindDuration > 0 Then
            Damage = 0
            SendActionMsg GetPlayerMap(victim), "Miss!", BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
    
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
    ' send animation
    If n > 0 Then
        If SpellNum = 0 Then
            Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, GetPlayerX(victim), GetPlayerY(victim))
            SendMapSound attacker, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seItem, n
        End If
    End If
    
    If SpellNum > 0 Then
        Call SendAnimation(GetPlayerMap(victim), Spell(SpellNum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim))
        SendMapSound attacker, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
    End If
    
    Call SendFlash(victim, GetPlayerMap(victim), False)
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
    
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)
        
             If GetPlayerLevel(victim) > 0 Then

         ' exp deduction

         If GetPlayerLevel(victim) <= GetPlayerLevel(attacker) - 10 Then

             ' 10 levels lower, exp 0

             exp = 0

         ElseIf GetPlayerLevel(victim) <= GetPlayerLevel(attacker) - 5 Then

             ' half exp if enemy is 5 levels lower

             exp = exp / 2

         ElseIf GetPlayerLevel(victim) >= GetPlayerLevel(attacker) + 10 Then

         End If

     End If


        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerMap(attacker)
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!", BrightRed)
        End If
        
        Call CheckTasks(attacker, QUEST_TYPE_GOKILL, victim)
        
        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
            'Blind
            If Spell(SpellNum).BlindDuration > 0 Then BlindPlayer victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player victim, SpellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellSlot As Long, Optional ByVal Cancel As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(index, spellSlot)
    mapnum = GetPlayerMap(index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough spirit!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    Target = TempPlayer(index).Target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If Cancel = 2 Then
        'SendAnimation mapnum, 0, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg mapnum, "Spell Canceled!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(index).spellBuffer.Spell = 0
        TempPlayer(index).spellBuffer.Timer = 0
        TempPlayer(index).spellBuffer.Target = 0
        TempPlayer(index).spellBuffer.tType = 0
        SendClearSpellBuffer index
        Exit Sub
    End If
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(SpellNum).CastAnim, GetPlayerX(index), GetPlayerY(index), TARGET_TYPE_PLAYER, index
        SendActionMsg mapnum, "Casting " & Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(index).spellBuffer.Spell = spellSlot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.Target = TempPlayer(index).Target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellSlot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
   
   Dim Dur As Long
   
    Dim buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(index, spellSlot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough spirit!", BrightRed)
        Exit Sub
    End If
   
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(SpellNum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(SpellNum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
   
    ' set the vital
      If Spell(SpellNum).Type <> SPELL_TYPE_BUFF Then
        Vital = Spell(SpellNum).Vital
        Vital = Round((Vital * 0.6)) * Round((Player(index).Level * 1.14)) * Round((Stats.Intelligence + (Stats.Willpower / 2)))
   
        If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Willpower) * 1.2))
        End If
   
        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEHP Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Intelligence) * 1.2))
        End If
    End If
   
    If Spell(SpellNum).Type = SPELL_TYPE_BUFF Then
        If Round(GetPlayerStat(index, Stats.Willpower) / 5) > 1 Then
            Dur = Spell(SpellNum).Duration * Round(GetPlayerStat(index, Stats.Willpower) / 5)
        Else
            Dur = Spell(SpellNum).Duration
        End If
    End If
    
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(SpellNum).Type
            
                Case SPELL_TYPE_BUFF
                        Call ApplyBuff(index, Spell(SpellNum).BuffType, Dur, Spell(SpellNum).Vital)
                        SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                        ' send the sound
                        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
                        DidCast = True
            
                Case SPELL_TYPE_HEALHP
                
                If Spell(SpellNum).StealthDuration > 0 Then
                        StealthPlayer index, SpellNum
                        Player(index).Visible = 1
                        End If
                        
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                
                If Spell(SpellNum).StealthDuration > 0 Then
                        StealthPlayer index, SpellNum
                        End If
                        
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, SpellNum
                    DidCast = True
                    
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(SpellNum).Map, Spell(SpellNum).x, Spell(SpellNum).y
                    SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
                    
                    Case SPELL_TYPE_BUFF
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).BuffType <= BUFF_ADD_DEF And Map(GetPlayerMap(index)).Moral <> MAP_MORAL_NONE Or Spell(SpellNum).BuffType > BUFF_NONE And Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Then
                            Call ApplyBuff(Target, Spell(SpellNum).BuffType, Dur, Spell(SpellNum).Vital)
                            SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                            ' send the sound
                            SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
                            DidCast = True
                        Else
                            PlayerMsg index, "You can not debuff another player in a safe zone!", BrightRed
                        End If
                    End If

                    
            End Select
        
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(SpellNum).Type
                
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, Vital, SpellNum
                                            
                                            If Spell(SpellNum).StealthDuration > 0 Then
                                            StealthPlayer index, SpellNum
                                            End If
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, Vital, SpellNum
                                        
                                        If Spell(SpellNum).StealthDuration > 0 Then
                                            StealthPlayer index, SpellNum
                                            End If
                                            
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        If Spell(SpellNum).StealthDuration > 0 Then
                        StealthPlayer index, SpellNum
                        End If
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        If Spell(SpellNum).StealthDuration > 0 Then
                        StealthPlayer index, SpellNum
                        End If
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        If Spell(SpellNum).StealthDuration > 0 Then
                        StealthPlayer index, SpellNum
                        End If
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        SpellPlayer_Effect VitalType, increment, i, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    'For i = 1 To MAX_MAP_NPCS
                        'If MapNpc(mapnum).NPC(i).num > 0 Then
                            'If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                'If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    'SpellNpc_Effect VitalType, increment, i, Vital, SpellNum, mapnum
                                'End If
                            'End If
                        'End If
                    'Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, Vital, SpellNum
                                
                                If Spell(SpellNum).StealthDuration > 0 Then
                                StealthPlayer index, SpellNum
                                End If
                                
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, Vital, SpellNum
                                
                                If Spell(SpellNum).StealthDuration > 0 Then
                                StealthPlayer index, SpellNum
                                End If
                                
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, SpellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, SpellNum
                        End If
                    Else
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, SpellNum, mapnum
                            End If
                        Else
                            If Not Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                            SpellNpc_Effect VitalType, increment, Target, Vital, SpellNum, mapnum
                            End If
                        End If
                    End If
                    
                    Case SPELL_TYPE_BUFF
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).BuffType <= BUFF_ADD_DEF And Map(GetPlayerMap(index)).Moral <> MAP_MORAL_NONE Or Spell(SpellNum).BuffType > BUFF_NONE And Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Then
                            Call ApplyBuff(Target, Spell(SpellNum).BuffType, Dur, Spell(SpellNum).Vital)
                            SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                            ' send the sound
                            SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
                            DidCast = True
                        Else
                            PlayerMsg index, "You can not debuff another player in a safe zone!", BrightRed
                        End If
                    End If
                    
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellSlot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
        Call SendCooldown(index, spellSlot)
        SetPlayerSpellUsage index, spellSlot
        SendActionMsg mapnum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seAnimation, Spell(SpellNum).SpellAnim
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player index, SpellNum
            End If
            Call SendVital(index, Vitals.HP)
            Call SendVital(index, Vitals.MP)
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
            Call SendVital(index, Vitals.HP)
            Call SendVital(index, Vitals.MP)
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendActionMsg mapnum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
        
        ' send the sound
        SendMapSound index, MapNpc(mapnum).NPC(index).x, MapNpc(mapnum).NPC(index).y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            MapNpc(mapnum).NPC(index).Vital(Vital) = MapNpc(mapnum).NPC(index).Vital(Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Npc mapnum, index, SpellNum
            End If
        ElseIf Not increment Then
            MapNpc(mapnum).NPC(index).Vital(Vital) = MapNpc(mapnum).NPC(index).Vital(Vital) - Damage
        End If
        ' send update
        SendMapNpcVitals mapnum, index
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
Dim Vital As Long

    With TempPlayer(index).HoT(hotNum)
    
    If .Used And .Spell > 0 Then
    
        Vital = Spell(.Spell).Vital
        Vital = Round((Vital * 0.6)) * Round((Player(index).Level * 1.14)) * (Round((Stats.Intelligence + (Stats.Willpower / 2))) / 1.5)
    
        If Spell(.Spell).WillBase = True Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Willpower) * 0.6))
        End If
    
        If Spell(.Spell).IntBase = True Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Intelligence) * 0.4))
        End If
        
        If Spell(.Spell).StrBase = True Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Strength) * 0.3))
        End If
        
        If Spell(.Spell).AgiBase = True Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Agility) * 0.4))
        End If
        
        If Spell(.Spell).EndBase = True Then
            Vital = Vital + Round((GetPlayerStat(index, Stats.Endurance) * 0.3))
        End If

            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                   SendActionMsg Player(index).Map, "+" & Vital, BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                   SetPlayerVital index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Vital
                   Call SendVital(index, Vitals.HP)
                Else
                   SendActionMsg Player(index).Map, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                   SetPlayerVital index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Vital
                   Call SendVital(index, Vitals.MP)
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long)
Dim Vital As Long
    With MapNpc(mapnum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
        
        Vital = Spell(.Spell).Vital
        Vital = Round((Vital * 0.6)) * Round((Player(index).Level * 1.14)) * (Round((Stats.Intelligence + (Stats.Willpower / 2))) / 1.5)
    
        If Spell(.Spell).WillBase = True Then
            Vital = Vital + Round((GetPlayerStat(.Caster, Stats.Willpower) * 0.6))
        End If
    
        If Spell(.Spell).IntBase = True Then
            Vital = Vital + Round((GetPlayerStat(.Caster, Stats.Intelligence) * 0.4))
        End If
        
        If Spell(.Spell).StrBase = True Then
            Vital = Vital + Round((GetPlayerStat(.Caster, Stats.Strength) * 0.3))
        End If
        
        If Spell(.Spell).AgiBase = True Then
            Vital = Vital + Round((GetPlayerStat(.Caster, Stats.Agility) * 0.4))
        End If
        
        If Spell(.Spell).EndBase = True Then
            Vital = Vital + Round((GetPlayerStat(.Caster, Stats.Endurance) * 0.3))
        End If
        
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpcDOT .Caster, index, Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal hotNum As Long)
    With MapNpc(mapnum).NPC(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                    MapNpc(mapnum).NPC(index).Vital(Vitals.HP) = MapNpc(mapnum).NPC(index).Vital(Vitals.HP) + Spell(.Spell).Vital
                Else
                    SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                    MapNpc(mapnum).NPC(index).Vital(Vitals.MP) = MapNpc(mapnum).NPC(index).Vital(Vitals.MP) + Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "You have been Stunned!", BrightCyan
    End If
End Sub

Public Sub BlindPlayer(ByVal index As Long, ByVal SpellNum As Long)
    ' check if it's a blinding spell
    If Spell(SpellNum).BlindDuration > 0 Then
        ' set the values on index
        TempPlayer(index).BlindDuration = Spell(SpellNum).BlindDuration
        TempPlayer(index).BlindTimer = GetTickCount
        ' send it to the index
        SendBlinded index
        ' tell him he's blinded
    SendActionMsg index, "You have been Blinded!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    End If
End Sub

Public Sub StealthPlayer(ByVal index As Long, ByVal SpellNum As Long)
    ' check if it's a stealth spell
    If Spell(SpellNum).StealthDuration > 0 Then
        
        ' set the values on index
        TempPlayer(index).StealthDuration = Spell(SpellNum).StealthDuration
        TempPlayer(index).StealthTimer = GetTickCount
        Player(index).Visible = 1
        Call SetPlayerColorA(index, 100)
        ' send it to the index
        SendStealthed index
        ' tell him he's stealthed
    PlayerMsg index, "You have been Stealthed!", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal mapnum As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(mapnum).NPC(index).StunTimer = GetTickCount
    End If
End Sub

Public Sub BlindNPC(ByVal index As Long, ByVal mapnum As Long, ByVal SpellNum As Long)
    ' check if it's a blinding spell
    If Spell(SpellNum).BlindDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(index).BlindDuration = Spell(SpellNum).BlindDuration
        MapNpc(mapnum).NPC(index).BlindTimer = GetTickCount
    End If
End Sub

Public Sub TryPlayerShootNpc(ByVal index As Long, ByVal MapNpcNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim n As Long
Dim stat As Stats

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootNpc(index, MapNpcNum) Then
    
        mapnum = GetPlayerMap(index)
        NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
        Call CreateProjectile(mapnum, index, TARGET_TYPE_PLAYER, MapNpcNum, TARGET_TYPE_NPC, Item(GetPlayerEquipment(index, weapon)).Projectile, Item(GetPlayerEquipment(index, weapon)).Rotation)
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        'Weapon Element Fire
If Item(Player(index).Equipment(Equipment.weapon)).Element = 1 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage / 2
        SendActionMsg mapnum, "Not very Effective...", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 2
        SendActionMsg mapnum, "Very Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End If
        
'Weapon Element Water
If Item(Player(index).Equipment(Equipment.weapon)).Element = 2 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage * 2
        SendActionMsg mapnum, "Very Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If

'Weapon Element Wind
'Give Wind element weapons more attack speed
If Item(Player(index).Equipment(Equipment.weapon)).Element = 3 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage / 2
        SendActionMsg mapnum, "Not Very Effective...", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If

'Weapon Element Earth
If Item(Player(index).Equipment(Equipment.weapon)).Element = 4 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If

'Weapon Element Light
If Item(Player(index).Equipment(Equipment.weapon)).Element = 5 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage * 2
        SendActionMsg mapnum, "Very Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End If

'Weapon Element Dark
If Item(Player(index).Equipment(Equipment.weapon)).Element = 6 Then
        
        If NPC(NPCNum).Element <> 0 Then 'Calculate Damage

        If NPC(NPCNum).Element = 1 Then
        Damage = Damage / 2
        SendActionMsg mapnum, "Not very Effective...", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 2 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 3 Then
        Damage = Damage * 1.5
        SendActionMsg mapnum, "Effective!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If

        If NPC(NPCNum).Element = 4 Then
        Damage = Damage * 0.8
        End If

        If NPC(NPCNum).Element = 5 Then
        Damage = Damage
        End If

        If NPC(NPCNum).Element = 6 Then
        Damage = Damage
        End If
    End If
End If
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand((NPC(NPCNum).stat(Stats.Endurance) * 1.1), (NPC(NPCNum).stat(Stats.Endurance) * 3))
        
        
        'Add Scaling based on agility
        'Damage = Damage + Round((Player(index).Level * 1.14)) + Round((GetPlayerStat(index, Stats.Agility) * 0.6))
        
        If Item(Player(index).Equipment(Equipment.weapon)).MageProjectile = True Then
                Damage = Damage + Round((Player(index).Level * 1.14)) + Round(GetPlayerStat(index, Intelligence) * 0.6)
            Else
                Damage = Damage + Round((Player(index).Level * 1.14)) + Round((GetPlayerStat(index, Stats.Agility) * 0.6))
        End If
        
        ' randomise from 1 to max hit
        Damage = rand(Damage / 3, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
        
        If TempPlayer(index).StealthDuration > 0 Then
        Damage = Damage * 3
        SendActionMsg mapnum, "Stealth hit!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        TempPlayer(index).StealthDuration = 0
        TempPlayer(index).StealthTimer = 0
        Player(index).Visible = 0
        Call SetPlayerColorA(index, 255)
        
        SendStealthed index
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, MapNpcNum, Damage, -1)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Public Function CanPlayerShootNpc(ByVal attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim AttackSpeed As Long
    Dim R As Long
    Dim NPCXDif As Long
    Dim NPCYDif As Long
    

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).NPC(MapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(attacker)
    NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
        ' attack speed from weapon
        If GetPlayerEquipment(attacker, weapon) > 0 Then
            If Not isInRange(Item(GetPlayerEquipment(attacker, weapon)).Range, GetPlayerX(attacker), GetPlayerY(attacker), MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y) Then Exit Function
            AttackSpeed = Item(GetPlayerEquipment(attacker, weapon)).speed
        Else
            AttackSpeed = 1000
        End If

        If NPCNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + AttackSpeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x + 1
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x - 1
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y
            End Select
            
            'get the difference between the points
            
            If MapNpc(mapnum).NPC(MapNpcNum).x > Player(attacker).x Then
            
           
            NPCXDif = (MapNpc(mapnum).NPC(MapNpcNum).x - Player(attacker).x)
            
            For R = 1 To NPCXDif
                If Map(mapnum).Tile((Player(attacker).x + R), Player(attacker).y).Type = 1 Or Map(mapnum).Tile((Player(attacker).x + R), Player(attacker).y).Type = 7 Then
                    CanPlayerShootNpc = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            Else
            
            NPCXDif = (Player(attacker).x - (MapNpc(mapnum).NPC(MapNpcNum).x))
            
            For R = 1 To NPCXDif
                If Map(mapnum).Tile((Player(attacker).x - R), Player(attacker).y).Type = 1 Or Map(mapnum).Tile((Player(attacker).x - R), Player(attacker).y).Type = 7 Then
                    CanPlayerShootNpc = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            End If
            
            'y axis
            
            If MapNpc(mapnum).NPC(MapNpcNum).y > Player(attacker).y Then
            
           
            NPCYDif = (MapNpc(mapnum).NPC(MapNpcNum).y - Player(attacker).y)
            
            For R = 1 To NPCYDif
                If Map(mapnum).Tile((Player(attacker).x), Player(attacker).y + R).Type = 1 Or Map(mapnum).Tile((Player(attacker).x), Player(attacker).y + R).Type = 7 Then
                    CanPlayerShootNpc = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            Else
            
            NPCYDif = (Player(attacker).y - (MapNpc(mapnum).NPC(MapNpcNum).y))
            
            For R = 1 To NPCYDif
                If Map(mapnum).Tile((Player(attacker).x), Player(attacker).y - R).Type = 1 Or Map(mapnum).Tile((Player(attacker).x), Player(attacker).y - R).Type = 7 Then
                    CanPlayerShootNpc = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            End If
            
            
            'NPCYDif = (MapNpc(mapnum).NPC(MapNpcNum).y - Player(attacker).y)
            'For R = 1 To Range
             '   If Map(mapnum).Tile((Player(attacker).x), Player(attacker).y + NPCYDif).Type = TILE_TYPE_BLOCKED Or Map(mapnum).Tile((Player(attacker).x), Player(attacker).y + NPCYDif).Type = TILE_TYPE_RESOURCE Then
              '      CanPlayerShootNpc = False
               '     Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                '    Exit Function
                'End If
            'Next
                
            If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(attacker).Target = MapNpcNum
                SendTarget attacker
                CanPlayerShootNpc = True
            Else
                If NpcX = GetPlayerX(attacker) Then
                    If NpcY = GetPlayerY(attacker) Then
                         If NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NPCNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            Call CheckTasks(attacker, QUEST_TYPE_GOTALK, NPCNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, NPCNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGET, NPCNum)
                            
                            
                            If NPC(NPCNum).Quest = YES Then
                                If Player(attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED Then
                                    If Quest(NPC(NPCNum).Quest).Repeat = YES Then
                                        Player(attacker).PlayerQuest(NPC(NPCNum).Quest).Status = QUEST_COMPLETED_BUT
                                        Exit Function
                                    End If
                                End If
                                If CanStartQuest(attacker, NPC(NPCNum).QuestNum) Then
                                    'if can start show the request message (speech1)
                                    QuestMessage attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(1)), NPC(NPCNum).QuestNum
                                    Exit Function
                                End If
                                If QuestInProgress(attacker, NPC(NPCNum).QuestNum) Then
                                    'if the quest is in progress show the meanwhile message (speech2)
                                    QuestMessage attacker, NPC(NPCNum).QuestNum, Trim$(Quest(NPC(NPCNum).QuestNum).Speech(2)), 0
                                    Exit Function
                                End If
                            End If
                        End If
                        If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                            Call SendChatBubble(mapnum, MapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Sub CreateProjectile(ByVal mapnum As Long, ByVal attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal RotateSpeed As Long)
Dim Rotate As Long
Dim buffer As clsBuffer
    
    If AttackerType = TARGET_TYPE_PLAYER Then
        ' ****** Initial Rotation Value ******
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(GetPlayerX(attacker), GetPlayerY(attacker), GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(GetPlayerX(attacker), GetPlayerY(attacker), MapNpc(mapnum).NPC(victim).x, MapNpc(mapnum).NPC(victim).y)
        End Select
    
        ' ****** Set Player Direction Based On Angle ******
       ' If Rotate >= 315 And Rotate <= 360 Then
           ' Call SetPlayerDir(attacker, DIR_UP)
        'ElseIf Rotate >= 0 And Rotate <= 45 Then
          '  Call SetPlayerDir(attacker, DIR_UP)
       ' ElseIf Rotate >= 225 And Rotate <= 315 Then
            'Call SetPlayerDir(attacker, DIR_LEFT)
        'ElseIf Rotate >= 135 And Rotate <= 225 Then
          '  Call SetPlayerDir(attacker, DIR_DOWN)
       ' ElseIf Rotate >= 45 And Rotate <= 135 Then
           ' Call SetPlayerDir(attacker, DIR_RIGHT)
        'End If
        
       ' Set buffer = New clsBuffer
       ' buffer.WriteLong SPlayerDir
        'buffer.WriteLong attacker
        'buffer.WriteLong GetPlayerDir(attacker)
        'Call SendDataToMap(mapnum, buffer.ToArray())
        'Set buffer = Nothing
    ElseIf AttackerType = TARGET_TYPE_NPC Then
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(MapNpc(mapnum).NPC(attacker).x, MapNpc(mapnum).NPC(attacker).y, GetPlayerX(victim), GetPlayerY(victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(MapNpc(mapnum).NPC(attacker).x, MapNpc(mapnum).NPC(attacker).y, MapNpc(mapnum).NPC(victim).x, MapNpc(mapnum).NPC(victim).y)
        End Select
    End If

    Call SendProjectile(mapnum, attacker, AttackerType, victim, targetType, Graphic, Rotate, RotateSpeed)
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal targetX As Integer, ByVal targetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = targetY Then
        'Check for going right (90 degrees)
        If CenterX < targetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
        
        'Exit the function
        Exit Function
    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = targetX Then
        'Check for going up (360 degrees)
        If CenterY > targetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function
    End If

    'Calculate Side C
    SideC = Sqr(Abs(targetX - CenterX) ^ 2 + Abs(targetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(targetX - CenterX) ^ 2 + targetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If targetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function
    Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function
End Function

Public Sub TryNpcShootPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
Dim mapnum As Long, NPCNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcShootPlayer(MapNpcNum, index) Then
        mapnum = GetPlayerMap(index)
        NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
        Call CreateProjectile(mapnum, MapNpcNum, TARGET_TYPE_NPC, index, TARGET_TYPE_PLAYER, NPC(NPCNum).Projectile, NPC(NPCNum).Rotation)
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, ((GetPlayerStat(index, Endurance) * 2) / rand(1, 3)))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, index, Damage)
        End If
    End If
End Sub

Function CanNpcShootPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim PlrXDif As Long
    Dim PlrYDif As Long
    Dim R As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
    NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).AttackSpeed > 0 Then
    If GetTickCount < MapNpc(mapnum).NPC(MapNpcNum).AttackTimer + NPC(MapNpc(mapnum).NPC(MapNpcNum).num).AttackSpeed Then
    'CanNpcShootPlayer = False
    Exit Function
    End If
    Else
    If GetTickCount < MapNpc(mapnum).NPC(MapNpcNum).AttackTimer + 1000 Then
    Exit Function
    End If
    End If
    
    MapNpc(mapnum).NPC(MapNpcNum).AttackTimer = GetTickCount
    
      'get the difference between the points
            
            If Player(index).x > MapNpc(mapnum).NPC(MapNpcNum).x Then
            
           
            PlrXDif = Player(index).x - MapNpc(mapnum).NPC(MapNpcNum).x
            
            For R = 1 To PlrXDif
                If Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x + R), MapNpc(mapnum).NPC(MapNpcNum).y).Type = 1 Or Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x + R), MapNpc(mapnum).NPC(MapNpcNum).y).Type = 7 Then
                    CanNpcShootPlayer = False
                    'Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            Else
            
            PlrXDif = MapNpc(mapnum).NPC(MapNpcNum).x - (Player(index).x)
            
            For R = 1 To PlrXDif
                 If Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x - R), MapNpc(mapnum).NPC(MapNpcNum).y).Type = 1 Or Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x - R), MapNpc(mapnum).NPC(MapNpcNum).y).Type = 7 Then
                    CanNpcShootPlayer = False
                    'Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            End If
            
            'y axis
            
            If Player(index).y > MapNpc(mapnum).NPC(MapNpcNum).y Then
            
           
            PlrYDif = Player(index).y - MapNpc(mapnum).NPC(MapNpcNum).y
            
            For R = 1 To PlrYDif
                 If Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x), MapNpc(mapnum).NPC(MapNpcNum).y + R).Type = 1 Or Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x), MapNpc(mapnum).NPC(MapNpcNum).y + R).Type = 7 Then
                    CanNpcShootPlayer = False
                    'Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            Else
            
            PlrYDif = MapNpc(mapnum).NPC(MapNpcNum).y - Player(index).y
            
            For R = 1 To PlrYDif
                 If Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x), MapNpc(mapnum).NPC(MapNpcNum).y - R).Type = 1 Or Map(mapnum).Tile((MapNpc(mapnum).NPC(MapNpcNum).x), MapNpc(mapnum).NPC(MapNpcNum).y - R).Type = 7 Then
                    CanNpcShootPlayer = False
                    'Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            End If

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NPCNum > 0 Then
            If isInRange(NPC(NPCNum).ProjectileRange, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, GetPlayerX(index), GetPlayerY(index)) Then
                CanNpcShootPlayer = True
            End If
        End If
    End If
End Function


Function CanPlayerShootPlayer(ByVal attacker As Long, ByVal victim As Long) As Boolean
Dim PlrXDif As Long
Dim PlrYDif As Long
Dim R As Long
Dim mapnum As Long

mapnum = GetPlayerMap(attacker)


    ' Check attack timer
    If GetPlayerEquipment(attacker, weapon) > 0 Then
        If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, weapon)).speed Then Exit Function
    Else
        If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If
    
      'get the difference between the points
            
            If Player(victim).x > Player(attacker).x Then
            
           
            PlrXDif = Player(victim).x - Player(attacker).x
            
            For R = 1 To PlrXDif
                If Map(mapnum).Tile((Player(attacker).x + R), Player(attacker).y).Type = 1 Or Map(mapnum).Tile((Player(attacker).x + R), Player(attacker).y).Type = 7 Then
                    CanPlayerShootPlayer = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            Else
            
            PlrXDif = Player(attacker).x - Player(victim).x
            
            For R = 1 To PlrXDif
                If Map(mapnum).Tile((Player(attacker).x - R), Player(attacker).y).Type = 1 Or Map(mapnum).Tile((Player(attacker).x - R), Player(attacker).y).Type = 7 Then
                    CanPlayerShootPlayer = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            End If
            
            'y axis
            
            If Player(victim).y > Player(attacker).y Then
            
           
            PlrYDif = (Player(victim).y - Player(attacker).y)
            
            For R = 1 To PlrYDif
                If Map(mapnum).Tile((Player(attacker).x), Player(attacker).y + R).Type = 1 Or Map(mapnum).Tile((Player(attacker).x), Player(attacker).y + R).Type = 7 Then
                    CanPlayerShootPlayer = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            Else
            
            PlrYDif = (Player(attacker).y - Player(victim).y)
            
            For R = 1 To PlrYDif
                If Map(mapnum).Tile((Player(attacker).x), Player(attacker).y - R).Type = 1 Or Map(mapnum).Tile((Player(attacker).x), Player(attacker).y - R).Type = 7 Then
                    CanPlayerShootPlayer = False
                    Call PlayerMsg(attacker, "Can't get a clear shot!", BrightRed)
                    Exit Function
                End If
            Next
            
            End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function
    TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).Target = victim
    SendTarget attacker
    CanPlayerShootPlayer = True
End Function
Public Sub TryPlayerShootPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim n As Long
Dim stat As Stats

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootPlayer(attacker, victim) Then
    
        mapnum = GetPlayerMap(attacker)
        Call CreateProjectile(mapnum, attacker, TARGET_TYPE_PLAYER, victim, TARGET_TYPE_PLAYER, Item(GetPlayerEquipment(attacker, weapon)).Projectile, Item(GetPlayerEquipment(attacker, weapon)).Rotation)
        ' check if player can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        If Item(Player(attacker).Equipment(Equipment.weapon)).MageProjectile = True Then
                Damage = Damage + Round((Player(attacker).Level * 1.14)) + Round(GetPlayerStat(attacker, Intelligence) * 0.6)
            Else
                Damage = Damage + Round((Player(attacker).Level * 1.14)) + Round((GetPlayerStat(attacker, Stats.Agility) * 0.6))
        End If
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(victim, Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage, -1)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Sub FireProjectile(ByVal index As Long, ByVal Dir As Long, ByVal Damage As Long, ByVal Range As Long, Optional ByVal SpellNum As Long = 0, Optional ByVal itemnum As Long = 0)
Dim i As Long
Dim R As Long
Dim mapnum As Long
Dim canShoot As Boolean
Dim DidCast As Boolean

    mapnum = GetPlayerMap(index)

    canShoot = False
    DidCast = False

    For R = 1 To Range
        Select Case Dir
            Case DIR_UP
                If GetPlayerY(index) - R < 0 Then Exit Sub
                If Map(mapnum).Tile(GetPlayerX(index), GetPlayerY(index) - R).Type = TILE_TYPE_BLOCKED Or Map(mapnum).Tile(GetPlayerX(index), GetPlayerY(index) - R).Type = TILE_TYPE_RESOURCE Then
                    'If SpellNum > 0 Then
                     '   DidCast = True
                    'Else
                        Exit Sub
                    'End If
                End If
            Case DIR_DOWN
                If GetPlayerY(index) + R > Map(mapnum).MaxY Then Exit Sub
                If Map(mapnum).Tile(GetPlayerX(index), GetPlayerY(index) + R).Type = TILE_TYPE_BLOCKED Or Map(mapnum).Tile(GetPlayerX(index), GetPlayerY(index) + R).Type = TILE_TYPE_RESOURCE Then
                   ' If SpellNum > 0 Then
                       ' DidCast = True
                    'Else
                        Exit Sub
                    'End If
                End If
            Case DIR_LEFT
                If GetPlayerX(index) - R < 0 Then Exit Sub
                If Map(mapnum).Tile(GetPlayerX(index) - R, GetPlayerY(index)).Type = TILE_TYPE_BLOCKED Or Map(mapnum).Tile(GetPlayerX(index) - R, GetPlayerY(index)).Type = TILE_TYPE_RESOURCE Then
                    'If SpellNum > 0 Then
                     '   DidCast = True
                   ' Else
                        Exit Sub
                   ' End If
                End If
            Case DIR_RIGHT
                If GetPlayerX(index) + R > Map(mapnum).MaxX Then Exit Sub
                If Map(mapnum).Tile(GetPlayerX(index) + R, GetPlayerY(index)).Type = TILE_TYPE_BLOCKED Or Map(mapnum).Tile(GetPlayerX(index) + R, GetPlayerY(index)).Type = TILE_TYPE_RESOURCE Then
                    'If SpellNum > 0 Then
                      '  DidCast = True
                    'Else
                        Exit Sub
                    'End If
                End If
        End Select

        If DidCast = True Then
            Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - Spell(Player(index).Spell(SpellNum)).MPCost)
            Call SendVital(index, Vitals.MP)
        
            TempPlayer(index).SpellCD(Player(index).Spell(SpellNum)) = GetTickCount + (Spell(Player(index).Spell(SpellNum)).CDTime * 1000)
            Call SendCooldown(index, SpellNum)
            SendActionMsg mapnum, Trim$(Spell(Player(index).Spell(SpellNum)).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
            Exit Sub
        End If
        
        For i = 1 To MAX_MAP_NPCS
            If Dir = DIR_UP Then
                If GetPlayerX(index) = MapNpc(mapnum).NPC(i).x And GetPlayerY(index) - R = MapNpc(mapnum).NPC(i).y Then
                    canShoot = True
                End If
            End If
            If Dir = DIR_DOWN Then
                If GetPlayerX(index) = MapNpc(mapnum).NPC(i).x And GetPlayerY(index) + R = MapNpc(mapnum).NPC(i).y Then
                    canShoot = True
                End If
            End If
            If Dir = DIR_LEFT Then
                If GetPlayerX(index) - R = MapNpc(mapnum).NPC(i).x And GetPlayerY(index) = MapNpc(mapnum).NPC(i).y Then
                    canShoot = True
                End If
            End If
            If Dir = DIR_RIGHT Then
                If GetPlayerX(index) + R = MapNpc(mapnum).NPC(i).x And GetPlayerY(index) = MapNpc(mapnum).NPC(i).y Then
                    canShoot = True
                End If
            End If
            
            If canShoot = True Then
                If SpellNum > 0 Then
                    CastSpell index, SpellNum, i, TARGET_TYPE_NPC
                    Exit Sub
                Else
                    Call TryPlayerAttackNpc(index, i)
                    Exit Sub
                End If
            End If
        Next
        
        For i = 1 To Player_HighIndex
            If i <> index Then
                If Dir = DIR_UP Then
                    If GetPlayerX(index) = GetPlayerX(i) And GetPlayerY(index) - R = GetPlayerY(i) Then
                        canShoot = True
                    End If
                End If
                If Dir = DIR_DOWN Then
                    If GetPlayerX(index) = GetPlayerX(i) And GetPlayerY(index) + R = GetPlayerY(i) Then
                        canShoot = True
                    End If
                End If
                If Dir = DIR_LEFT Then
                    If GetPlayerX(index) - R = GetPlayerX(i) And GetPlayerY(index) = GetPlayerY(i) Then
                        canShoot = True
                    End If
                End If
                If Dir = DIR_RIGHT Then
                    If GetPlayerX(index) + R = GetPlayerX(i) And GetPlayerY(index) = GetPlayerY(i) Then
                        canShoot = True
                    End If
                End If
                    
                If canShoot = True Then
                    If SpellNum > 0 Then
                        CastSpell index, SpellNum, i, TARGET_TYPE_PLAYER
                        Exit Sub
                    Else
                        Call TryPlayerAttackPlayer(index, i)
                        Exit Sub
                    End If
                End If
            End If
        Next
    Next
    
    Player(index).x = Player(index).x
    Player(index).y = Player(index).y
End Sub

Public Sub SetPlayerSpellUsage(ByVal index As Long, ByVal spellSlot As Long)
Dim SpellNum As Long, i As Long
    SpellNum = Player(index).Spell(spellSlot)
    ' if has a next rank then increment usage
    If Spell(SpellNum).NextRank > 0 Then
        If Player(index).SpellUses(spellSlot) < Spell(SpellNum).NextUses - 1 Then
            Player(index).SpellUses(spellSlot) = Player(index).SpellUses(spellSlot) + 1
        Else
            If GetPlayerLevel(index) >= Spell(Spell(SpellNum).NextRank).LevelReq Then
                Player(index).Spell(spellSlot) = Spell(SpellNum).NextRank
                Player(index).SpellUses(spellSlot) = 0
                PlayerMsg index, "Your spell has ranked up!", Blue
                ' update hotbar
                For i = 1 To MAX_HOTBAR
                    If Player(index).Hotbar(i).Slot > 0 Then
                        If Player(index).Hotbar(i).sType = 2 Then ' spell
                            If Spell(Player(index).Hotbar(i).Slot).NextRank = Spell(SpellNum).NextRank Then
                                Player(index).Hotbar(i).Slot = Spell(SpellNum).NextRank
                                SendHotbar index
                            End If
                        End If
                    End If
                Next
            Else
                Player(index).SpellUses(spellSlot) = Spell(SpellNum).NextUses
            End If
        End If
        SendPlayerSpells index
    End If
End Sub

Public Sub PlayerAttackNpcDOT(ByVal attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim Def As Long
    Dim mapnum As Long
    Dim NPCNum As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    NPCNum = MapNpc(mapnum).NPC(MapNpcNum).num
    Name = Trim$(NPC(NPCNum).Name)
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
    
    If Damage > 0 Then

    ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y
    
    
    If Damage >= MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) Then
        
        ' check if the server mod stops exp flow.
        If (frmServer.scrlExp.Value <> 0) Then
            ' Calculate exp to give attacker
            exp = NPC(NPCNum).exp * frmServer.scrlExp.Value
    
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
    
            ' in party?
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party sharing function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerMap(attacker)
            Else
                ' no party - keep exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        'Drop the goods if they get it
            For n = 1 To MAX_NPC_DROPS
            If NPC(NPCNum).DropItem(n) = 0 Then Exit For
            
            If Rnd <= NPC(NPCNum).DropChance(n) Then
            Call SpawnItem(NPC(NPCNum).DropItem(n), NPC(NPCNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y)
            End If
            Next
            
            If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).isBoss = YES Then
            SendBossMsg Trim$(NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
            GlobalMsg Trim$(NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Name) & " has been slain by " & Trim$(GetPlayerName(attacker)) & " in " & Trim$(Map(GetPlayerMap(attacker)).Name) & ".", Magenta
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(MapNpcNum).num = 0
        MapNpc(mapnum).NPC(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = 0
        UpdateMapBlock mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, False
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(MapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(MapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, NPCNum)
        
        ' send death to the map
        Set buffer = New clsBuffer
        buffer.WriteLong SNpcDead
        buffer.WriteLong MapNpcNum
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = MapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) - Damage

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(MapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(MapNpcNum).num Then
                    MapNpc(mapnum).NPC(i).Target = attacker
                    MapNpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).NPC(MapNpcNum).stopRegen = True
        MapNpc(mapnum).NPC(MapNpcNum).stopRegenTimer = GetTickCount
        
        End If
        
        SendMapNpcVitals mapnum, MapNpcNum
End If
End Sub
