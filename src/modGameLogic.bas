Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).canDespawn = canDespawn
            MapItem(mapnum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).num = itemnum
            MapItem(mapnum, i).Value = ItemVal
            MapItem(mapnum, i).x = x
            MapItem(mapnum, i).y = y
            ' send to map
            SendSpawnItemToMap mapnum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(mapnum).Tile(x, y).Data1).Stackable > 0 And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal mapnum As Long, Optional ForcedSpawn As Boolean = False)
    Dim buffer As clsBuffer
    Dim NPCNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    NPCNum = Map(mapnum).NPC(MapNpcNum)
    If ForcedSpawn = False And Map(mapnum).NpcSpawnType(MapNpcNum) = 1 Then NPCNum = 0
    If NPCNum > 0 Then
    
        MapNpc(mapnum).NPC(MapNpcNum).num = NPCNum
        MapNpc(mapnum).NPC(MapNpcNum).Target = 0
        MapNpc(mapnum).NPC(MapNpcNum).targetType = 0 ' clear
        
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NPCNum, Vitals.HP)
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NPCNum, Vitals.MP)
        
        MapNpc(mapnum).NPC(MapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapnum).MaxX
            For y = 0 To Map(mapnum).MaxY
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapnum).Tile(x, y).Data1 = MapNpcNum Then
                        MapNpc(mapnum).NPC(MapNpcNum).x = x
                        MapNpc(mapnum).NPC(MapNpcNum).y = y
                        MapNpc(mapnum).NPC(MapNpcNum).Dir = Map(mapnum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(mapnum).MaxX)
                y = Random(0, Map(mapnum).MaxY)
    
                If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
                If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).NPC(MapNpcNum).x = x
                    MapNpc(mapnum).NPC(MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapnum).MaxX
                For y = 0 To Map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).NPC(MapNpcNum).x = x
                        MapNpc(mapnum).NPC(MapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set buffer = New clsBuffer
            buffer.WriteLong SSpawnNpc
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).num
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
            UpdateMapBlock mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, True
        End If
        
        SendMapNpcVitals mapnum, MapNpcNum
    Else
        MapNpc(mapnum).NPC(MapNpcNum).num = 0
        MapNpc(mapnum).NPC(MapNpcNum).Target = 0
        MapNpc(mapnum).NPC(MapNpcNum).targetType = 0 ' clear
        ' send death to the map
        Set buffer = New clsBuffer
        buffer.WriteLong SNpcDead
        buffer.WriteLong MapNpcNum
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
    End If

End Sub

Public Sub SpawnMapEventsFor(index As Long, mapnum As Long)
Dim i As Long, x As Long, y As Long, z As Long, spawncurrentevent As Boolean, p As Long
Dim buffer As clsBuffer
    
    TempPlayer(index).EventMap.CurrentEvents = 0
    ReDim TempPlayer(index).EventMap.EventPages(0)
    
    If Map(mapnum).EventCount <= 0 Then Exit Sub
    For i = 1 To Map(mapnum).EventCount
        If Map(mapnum).Events(i).PageCount > 0 Then
            For z = Map(mapnum).Events(i).PageCount To 1 Step -1
                With Map(mapnum).Events(i).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        If Player(index).Variables(.VariableIndex) < .VariableCondition Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSwitch = 1 Then
                        If Player(index).Switches(.SwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(index, .HasItemIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(i).SelfSwitches(.SelfSwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        'spawn the event... send data to player
                        TempPlayer(index).EventMap.CurrentEvents = TempPlayer(index).EventMap.CurrentEvents + 1
                        ReDim Preserve TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                        With TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                            If Map(mapnum).Events(i).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(i).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            .GraphicNum = Map(mapnum).Events(i).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(i).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(i).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(i).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(i).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(i).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(i).Pages(z).MoveSpeed
                                Case 0
                                    .movementspeed = 2
                                Case 1
                                    .movementspeed = 3
                                Case 2
                                    .movementspeed = 4
                                Case 3
                                    .movementspeed = 6
                                Case 4
                                    .movementspeed = 12
                                Case 5
                                    .movementspeed = 24
                            End Select
                            If Map(mapnum).Events(i).Global Then
                                .x = TempEventMap(mapnum).Events(i).x
                                .y = TempEventMap(mapnum).Events(i).y
                                .Dir = TempEventMap(mapnum).Events(i).Dir
                                .MoveRouteStep = TempEventMap(mapnum).Events(i).MoveRouteStep
                            Else
                                .x = Map(mapnum).Events(i).x
                                .y = Map(mapnum).Events(i).y
                                .MoveRouteStep = 0
                            End If
                            .Position = Map(mapnum).Events(i).Pages(z).Position
                            .eventID = i
                            .pageID = z
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(mapnum).Events(i).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount)
                                If Map(mapnum).Events(i).Pages(z).MoveRouteCount > 0 Then
                                    For p = 0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(mapnum).Events(i).Pages(z).MoveRoute(p)
                                    Next
                                End If
                            End If
                            
                            .RepeatMoveRoute = Map(mapnum).Events(i).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(i).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(i).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(mapnum).Events(i).Pages(z).WalkAnim
                            .WalkThrough = Map(mapnum).Events(i).Pages(z).WalkThrough
                            .ShowName = Map(mapnum).Events(i).Pages(z).ShowName
                            .FixedDir = Map(mapnum).Events(i).Pages(z).DirFix
                            
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For i = 1 To TempPlayer(index).EventMap.CurrentEvents
            Set buffer = New clsBuffer
            buffer.WriteLong SSpawnEvent
            buffer.WriteLong i
            With TempPlayer(index).EventMap.EventPages(i)
                buffer.WriteString Map(GetPlayerMap(index)).Events(i).Name
                buffer.WriteLong .Dir
                buffer.WriteLong .GraphicNum
                buffer.WriteLong .GraphicType
                buffer.WriteLong .GraphicX
                buffer.WriteLong .GraphicX2
                buffer.WriteLong .GraphicY
                buffer.WriteLong .GraphicY2
                buffer.WriteLong .movementspeed
                buffer.WriteLong .x
                buffer.WriteLong .y
                buffer.WriteLong .Position
                buffer.WriteLong .Visible
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkAnim
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).DirFix
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkThrough
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).ShowName
            End With
            SendDataTo index, buffer.ToArray
            Set buffer = Nothing
        Next
    End If
End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).NPC(LoopI).num > 0 Then
            If MapNpc(mapnum).NPC(LoopI).x = x Then
                If MapNpc(mapnum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next
    
    For LoopI = 1 To TempEventMap(mapnum).EventCount
        If TempEventMap(mapnum).Events(LoopI).active = 1 Then
            If MapNpc(mapnum).NPC(LoopI).x = TempEventMap(mapnum).Events(LoopI).x Then
                If MapNpc(mapnum).NPC(LoopI).y = TempEventMap(mapnum).Events(LoopI).y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_HEAL Then
                    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_CRAFT Then
                        NpcTileIsOpen = False
                    End If
                End If
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        If Map(mapnum).NPC(i) > 0 And Map(mapnum).NPC(i) <= MAX_NPCS Then
            If Options.DayNight = 1 Then
                If DayTime = True And NPC(Map(mapnum).NPC(i)).SpawnAtDay = 0 Then
                    Call SpawnNpc(i, mapnum)
                ElseIf DayTime = False And NPC(Map(mapnum).NPC(i)).SpawnAtNight = 0 Then
                    Call SpawnNpc(i, mapnum)
                End If
            Else
                Call SpawnNpc(i, mapnum)
            End If
        End If
    Next
    
    CacheMapBlocks mapnum

End Sub


Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Sub SpawnAllMapGlobalEvents()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnGlobalEvents(i)
    Next

End Sub

Sub SpawnGlobalEvents(ByVal mapnum As Long)
    Dim i As Long, z As Long
    
    If Map(mapnum).EventCount > 0 Then
        TempEventMap(mapnum).EventCount = 0
        ReDim TempEventMap(mapnum).Events(0)
        For i = 1 To Map(mapnum).EventCount
            TempEventMap(mapnum).EventCount = TempEventMap(mapnum).EventCount + 1
            ReDim Preserve TempEventMap(mapnum).Events(0 To TempEventMap(mapnum).EventCount)
            If Map(mapnum).Events(i).PageCount > 0 Then
                If Map(mapnum).Events(i).Global = 1 Then
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).x = Map(mapnum).Events(i).x
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).y = Map(mapnum).Events(i).y
                    If Map(mapnum).Events(i).Pages(1).GraphicType = 1 Then
                        Select Case Map(mapnum).Events(i).Pages(1).GraphicY
                            Case 0
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                    End If
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).active = 1
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = Map(mapnum).Events(i).Pages(1).MoveType
                    
                    If TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = 2 Then
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRouteCount = Map(mapnum).Events(i).Pages(1).MoveRouteCount
                        ReDim TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount)
                        For z = 0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount
                            TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(z) = Map(mapnum).Events(i).Pages(1).MoveRoute(z)
                        Next
                    End If
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).RepeatMoveRoute = Map(mapnum).Events(i).Pages(1).RepeatMoveRoute
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveFreq = Map(mapnum).Events(i).Pages(1).MoveFreq
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveSpeed = Map(mapnum).Events(i).Pages(1).MoveSpeed
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkThrough = Map(mapnum).Events(i).Pages(1).WalkThrough
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).FixedDir = Map(mapnum).Events(i).Pages(1).DirFix
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkingAnim = Map(mapnum).Events(i).Pages(1).WalkAnim
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).ShowName = Map(mapnum).Events(i).Pages(1).ShowName
                    
                End If
            End If
        Next
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).NPC(MapNpcNum).x
    y = MapNpc(mapnum).NPC(MapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_HEAL And n <> TILE_TYPE_CRAFT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_HEAL And n <> TILE_TYPE_CRAFT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_HEAL And n <> TILE_TYPE_CRAFT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x - 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_HEAL And n <> TILE_TYPE_CRAFT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x + 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(MapNpcNum).Dir = Dir
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, False

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).NPC(MapNpcNum).y = MapNpc(mapnum).NPC(MapNpcNum).y - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).NPC(MapNpcNum).y = MapNpc(mapnum).NPC(MapNpcNum).y + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).NPC(MapNpcNum).x = MapNpc(mapnum).NPC(MapNpcNum).x - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).NPC(MapNpcNum).x = MapNpc(mapnum).NPC(MapNpcNum).x + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
    End Select
    
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y, True

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    ' Check if have player on map
If PlayersOnMap(mapnum) = NO Then Exit Sub
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(MapNpcNum).Dir = Dir
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDir
    buffer.WriteLong MapNpcNum
    buffer.WriteLong Dir
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    temptile(mapnum).DoorTimer = 0
    ReDim temptile(mapnum).DoorOpen(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            temptile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(Map(mapnum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long
    Dim OldUses As Long, NewUses As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    OldUses = GetPlayerSpellUses(index, oldSlot)
    NewUses = GetPlayerSpellUses(index, newSlot)
    
    
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SetPlayerSpellUses index, oldSlot, NewUses
    SetPlayerSpellUses index, newSlot, OldUses
    
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        If Item(GetPlayerEquipment(index, EqSlot)).Stackable > 0 Then
            GiveInvItem index, GetPlayerEquipment(index, EqSlot), 1
        Else
            GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        End If
        SendActionMsg GetPlayerMap(index), "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If
End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            Party(partyNum).PTEXP = 0
            Party(partyNum).PTlevel = 0
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        Party(partyNum).PTEXP = 0
        Party(partyNum).PTlevel = 0
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal index As Long, ByVal mapnum As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long, LoseMemberCount As Byte


    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp
        Exit Sub
    End If
    
    ' check members in outhers maps
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> mapnum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next i
    
    
    ' Check if npc level is bigger then 0

     If GetPlayerLevel(Party(partyNum).MemberCount) > 0 Then

         ' exp deduction

         If GetPlayerLevel(Party(partyNum).MemberCount) <= GetPlayerLevel(Party(partyNum).Leader) - 10 Then

             ' 10 levels lower, exp 0

             PartyMsg partyNum, "Partymu 10 level di bawahmu, kamu tidak akan mendapatkan exp.", BrightRed
             exp = 0

         ElseIf GetPlayerLevel(Party(partyNum).MemberCount) >= GetPlayerLevel(Party(partyNum).Leader) + 10 Then

         End If

     End If
    
    'Calculate the EXP for the party itself
    If Party(partyNum).MemberCount > 1 Then
        Party(partyNum).PTEXP = Party(partyNum).PTEXP + exp
        Party(partyNum).PTCheckLevel = Party(partyNum).PTEXP / Round(100 * Party(partyNum).MemberCount * GetPlayerLevel(Party(partyNum).Leader))
    End If
      
    'check if part level has increased
    If (Party(partyNum).PTCheckLevel > Party(partyNum).PTlevel) Then
        Party(partyNum).PTlevel = Party(partyNum).PTCheckLevel
        If (Party(partyNum).PTlevel > 50) Then
            Party(partyNum).PTlevel = 50
            PartyMsg partyNum, "Party Level is LV " & Party(partyNum).PTlevel & " and the party will gain a " & Party(partyNum).PTlevel & "% exp boost.", BrightGreen
            PartyMsg partyNum, "Your party has reached the Max Level and will not level up any higher", BrightGreen
        Else
            PartyMsg partyNum, "Party Level Increased to LV " & Party(partyNum).PTlevel & " and the party will gain a " & Party(partyNum).PTlevel & "% exp boost.", BrightGreen
        End If
    End If
        
    exp = exp + Round(((exp / 100) * Party(partyNum).PTlevel))
    
    ' find out the equal share
    expShare = exp \ Party(partyNum).MemberCount
    leftOver = exp Mod Party(partyNum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(rand(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
    ' give the exp
    
    If Player(index).ExpMultiplierTime > 0 Then
    exp = exp * Player(index).ExpMultiplier
    End If
    
    If Player(index).IsMember = 1 Then exp = exp * 2
    
    Call SetPlayerExp(index, GetPlayerExp(index) + exp)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp index
End Sub

Function CanEventMove(index As Long, ByVal mapnum As Long, x As Long, y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional globalevent As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long, z As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    CanEventMove = True
    
    

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y - 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x) And (MapNpc(mapnum).NPC(i).y = y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y + 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x) And (MapNpc(mapnum).NPC(i).y = y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x - 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x - 1) And (MapNpc(mapnum).NPC(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x - 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x + 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x + 1) And (MapNpc(mapnum).NPC(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x + 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

    End Select

End Function

Sub EventDir(playerindex As Long, ByVal mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional globalevent As Boolean = False)
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(playerindex).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(playerindex).EventMap.EventPages(eventID).Dir = Dir
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventDir
    buffer.WriteLong eventID
    If globalevent Then
        buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
    Else
        buffer.WriteLong TempPlayer(playerindex).EventMap.EventPages(eventID).Dir
    End If
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub EventMove(index As Long, mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, movementspeed As Long, Optional globalevent As Boolean = False)
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
        UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, False
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(index).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(index).EventMap.EventPages(eventID).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
            
        Case DIR_DOWN
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
        Case DIR_LEFT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
        Case DIR_RIGHT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
    End Select

End Sub


Public Function KeepTwoDigit(num As Byte)
    If (num < 10) Then
        KeepTwoDigit = "0" & num
    Else
        KeepTwoDigit = num
    End If
End Function

Sub DespawnNPC(ByVal mapnum As Long, ByVal NPCNum As Long)
Dim i As Long, buffer As clsBuffer
                       
    ' Reset the targets..
    MapNpc(mapnum).NPC(NPCNum).Target = 0
    MapNpc(mapnum).NPC(NPCNum).targetType = TARGET_TYPE_NONE
    
    ' Set the NPC data to blank so it despawns.
    MapNpc(mapnum).NPC(NPCNum).num = 0
    MapNpc(mapnum).NPC(NPCNum).SpawnWait = 0
    MapNpc(mapnum).NPC(NPCNum).Vital(Vitals.HP) = 0
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(NPCNum).x, MapNpc(mapnum).NPC(NPCNum).y, False
        
    ' clear DoTs and HoTs
    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(NPCNum).DoT(i)
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
            .Used = False
        End With
            
        With MapNpc(mapnum).NPC(NPCNum).HoT(i)
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
            .Used = False
       End With
    Next
        
    ' send death to the map
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDead
    buffer.WriteLong NPCNum
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
        
    'Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If Player(i).Map = mapnum Then
                If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                    If TempPlayer(i).Target = NPCNum Then
                        TempPlayer(i).Target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
End Sub
