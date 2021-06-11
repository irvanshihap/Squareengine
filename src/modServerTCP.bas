Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If LenB(Trim$(Player(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim f As Long
    filename = App.path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    f = FreeFile
    Open filename For Input As #f

    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData buffer.ToArray()
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal mapnum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapnum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, buffer.ToArray
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToMap mapnum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo index, buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
    Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

   ' If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
    '    If TempPlayer(index).DataBytes > 1100 Then
     '       Exit Sub
      '  End If
    
        ' Check for packet flooding
       ' If TempPlayer(index).DataPackets > 30 Then
        '    Exit Sub
        'End If
    'End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(index).buffer.WriteBytes buffer()
    
    If TempPlayer(index).buffer.Length >= 4 Then
        pLength = TempPlayer(index).buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).buffer.Length - 4
        If pLength <= TempPlayer(index).buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).buffer.ReadLong
            HandleData index, TempPlayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).buffer.Length >= 4 Then
            pLength = TempPlayer(index).buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

End Sub

Public Sub MapCache_Create(ByVal mapnum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong mapnum
    buffer.WriteString Trim$(Map(mapnum).Name)
    buffer.WriteString Trim$(Map(mapnum).Music)
    buffer.WriteString Trim$(Map(mapnum).BGS)
    buffer.WriteLong Map(mapnum).Revision
    buffer.WriteByte Map(mapnum).Moral
    buffer.WriteLong Map(mapnum).Up
    buffer.WriteLong Map(mapnum).Down
    buffer.WriteLong Map(mapnum).Left
    buffer.WriteLong Map(mapnum).Right
    buffer.WriteLong Map(mapnum).BootMap
    buffer.WriteByte Map(mapnum).BootX
    buffer.WriteByte Map(mapnum).BootY
    
    buffer.WriteLong Map(mapnum).Weather
    buffer.WriteLong Map(mapnum).WeatherIntensity
    
    buffer.WriteLong Map(mapnum).Fog
    buffer.WriteLong Map(mapnum).FogSpeed
    buffer.WriteLong Map(mapnum).FogOpacity
    
    buffer.WriteLong Map(mapnum).Red
    buffer.WriteLong Map(mapnum).Green
    buffer.WriteLong Map(mapnum).Blue
    buffer.WriteLong Map(mapnum).Alpha
    
    buffer.WriteByte Map(mapnum).MaxX
    buffer.WriteByte Map(mapnum).MaxY

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            With Map(mapnum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).x
                    buffer.WriteLong .Layer(i).y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
                buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map(mapnum).NPC(x)
        buffer.WriteLong Map(mapnum).NpcSpawnType(x)
    Next

    buffer.WriteByte Map(mapnum).IsMember
    buffer.WriteByte Map(mapnum).DayNight
    buffer.WriteByte Map(mapnum).Panorama
    MapCache(mapnum).Data = buffer.ToArray()
    
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long
    If index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerLevel(index)
    buffer.WriteLong GetPlayerPOINTS(index)
    buffer.WriteLong GetPlayerSprite(index)
    buffer.WriteLong GetPlayerMap(index)
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteLong GetPlayerClass(index)
    buffer.WriteLong GetPlayerVisible(index)
    buffer.WriteLong GetPlayerColorA(index)
    buffer.WriteLong GetPlayerColorR(index)
    buffer.WriteLong GetPlayerColorG(index)
    buffer.WriteLong GetPlayerColorB(index)
    
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    If Player(index).GuildFileId > 0 Then
        If TempPlayer(index).tmpGuildSlot > 0 Then
            buffer.WriteByte 1
            buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
            buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Tag
            buffer.WriteLong GuildData(TempPlayer(index).tmpGuildSlot).Guild_Color
        End If
    Else
        buffer.WriteByte 0
    End If
    
        buffer.WriteByte Player(index).IsMember
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong index
    SendDataToMapBut index, mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal mapnum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(mapnum).Data) - LBound(MapCache(mapnum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(mapnum).Data()
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(mapnum, i).playerName
        buffer.WriteLong MapItem(mapnum, i).num
        buffer.WriteLong MapItem(mapnum, i).Value
        buffer.WriteLong MapItem(mapnum, i).x
        buffer.WriteLong MapItem(mapnum, i).y
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    
    Dim buffer As clsBuffer
    ' Check if have player on map
    If PlayersOnMap(mapnum) = NO Then Exit Sub
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(mapnum, i).playerName
        buffer.WriteLong MapItem(mapnum, i).num
        buffer.WriteLong MapItem(mapnum, i).Value
        buffer.WriteLong MapItem(mapnum, i).x
        buffer.WriteLong MapItem(mapnum, i).y
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapnum As Long, ByVal MapNpcNum As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    ' Check if have player on map
    If PlayersOnMap(mapnum) = NO Then Exit Sub
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcVitals
    buffer.WriteLong MapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Vital(i)
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
     ' Check if have player on map
    If PlayersOnMap(mapnum) = NO Then Exit Sub
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapnum).NPC(i).num
        buffer.WriteLong MapNpc(mapnum).NPC(i).x
        buffer.WriteLong MapNpc(mapnum).NPC(i).y
        buffer.WriteLong MapNpc(mapnum).NPC(i).Dir
        buffer.WriteLong MapNpc(mapnum).NPC(i).Vital(HP)
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    ' Check if have player on map
    If PlayersOnMap(mapnum) = NO Then Exit Sub
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapnum).NPC(i).num
        buffer.WriteLong MapNpc(mapnum).NPC(i).x
        buffer.WriteLong MapNpc(mapnum).NPC(i).y
        buffer.WriteLong MapNpc(mapnum).NPC(i).Dir
        buffer.WriteLong MapNpc(mapnum).NPC(i).Vital(HP)
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(index, i)
        buffer.WriteLong GetPlayerInvItemValue(index, i)
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteLong invSlot
    buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim packet As String
Dim buffer As clsBuffer
Set buffer = New clsBuffer

buffer.WriteLong SPlayerWornEq
buffer.WriteLong GetPlayerEquipment(index, Armor)
buffer.WriteLong GetPlayerEquipment(index, weapon)
buffer.WriteLong GetPlayerEquipment(index, Helmet)
buffer.WriteLong GetPlayerEquipment(index, Whetstone) ' New
buffer.WriteLong GetPlayerEquipment(index, Boots) ' New
buffer.WriteLong GetPlayerEquipment(index, Charm) ' New
buffer.WriteLong GetPlayerEquipment(index, Ring) ' New
buffer.WriteLong GetPlayerEquipment(index, Enchant) ' New
buffer.WriteLong GetPlayerEquipment(index, Shield)
SendDataTo index, buffer.ToArray()

Set buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
Dim buffer As clsBuffer
Set buffer = New clsBuffer

buffer.WriteLong SMapWornEq
buffer.WriteLong index
buffer.WriteLong GetPlayerEquipment(index, Armor)
buffer.WriteLong GetPlayerEquipment(index, weapon)
buffer.WriteLong GetPlayerEquipment(index, Helmet)
buffer.WriteLong GetPlayerEquipment(index, Whetstone)
buffer.WriteLong GetPlayerEquipment(index, Boots)
buffer.WriteLong GetPlayerEquipment(index, Charm)
buffer.WriteLong GetPlayerEquipment(index, Ring)
buffer.WriteLong GetPlayerEquipment(index, Enchant)
buffer.WriteLong GetPlayerEquipment(index, Shield)

SendDataToMap GetPlayerMap(index), buffer.ToArray()

Set buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
Dim buffer As clsBuffer
Set buffer = New clsBuffer

buffer.WriteLong SMapWornEq
buffer.WriteLong PlayerNum
buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
buffer.WriteLong GetPlayerEquipment(PlayerNum, weapon)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Whetstone)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Boots)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Charm)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Ring)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Enchant)
buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)

SendDataTo index, buffer.ToArray()

Set buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHp
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMp
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong GetPlayerExp(index)
    buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStats(ByVal index As Long)
Dim i As Long
Dim packet As String
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(index, i)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If
    
    ' Send visibility message
    If GetPlayerAccess(index) > ADMIN_MONITOR Then
        If GetPlayerVisible(index) = 1 Then
            Call PlayerMsg(index, "(invisible)", AlertColor)
        End If
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SClassesData
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub


Sub SendNewCharClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SNewCharClasses
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
    'If Class(i).Next = 0 Then
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(i).stat(q)
        Next
    'End If
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub


Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXYMap
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal itemnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NPCNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NPCNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(index, i)
        buffer.WriteLong GetPlayerSpellUses(index, i)
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).x
            buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).y
        Next

    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapnum As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    ' Check if have player on map
    If PlayersOnMap(mapnum) = NO Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(mapnum).Resource_Count

    If ResourceCache(mapnum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(mapnum).Resource_Count
            buffer.WriteByte ResourceCache(mapnum).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(mapnum).ResourceData(i).x
            buffer.WriteLong ResourceCache(mapnum).ResourceData(i).y
        Next

    End If

    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SDoorAnimation
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapnum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString message
    buffer.WriteLong Color
    buffer.WriteLong MsgType
    buffer.WriteLong x
    buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap mapnum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendBlood(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub



Sub SendAnimation(ByVal mapnum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, buffer.ToArray
    Else
        SendDataToMap mapnum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCooldown
    buffer.WriteLong Slot
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal mapnum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Map] "
    buffer.WriteLong saycolour
    
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Global] "
    buffer.WriteLong saycolour
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendBlinded(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlinded
    buffer.WriteLong TempPlayer(index).BlindDuration
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStealthed(ByVal index As Long)
    Dim buffer As clsBuffer
    
    'Call SetPlayerColorA(index)
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStealthed
    buffer.WriteLong TempPlayer(index).StealthDuration
    buffer.WriteLong GetPlayerColorA(index)
    
    'SendDataToAll buffer.ToArray()
    'SendDataToAllBut index, buffer.ToArray()
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Bank(index).Item(i).num
        buffer.WriteLong Bank(index).Item(i).Value
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte Value
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte Value
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong shopNum
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    buffer.WriteLong tradeTarget
    buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal GName As String, ByVal GInviterName As String, ByVal guildTarget As Long)
Dim buffer As clsBuffer

Set buffer = New clsBuffer
buffer.WriteLong SGuildInvite

buffer.WriteString GName
buffer.WriteString GInviterName

SendDataTo guildTarget, buffer.ToArray()

Set buffer = Nothing
End Sub


Sub SendCloseTrade(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(index).TradeOffer(i).num
            buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(i).num).Type = ITEM_TYPE_CURRENCY Or Item(TempPlayer(index).TradeOffer(i).num).Stackable > 0 Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).num)).price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)
            buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).Stackable > 0 Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    buffer.WriteLong totalWorth
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    buffer.WriteByte Status
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTarget(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    buffer.WriteLong TempPlayer(index).Target
    buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim i As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        buffer.WriteLong Player(index).Hotbar(i).Slot
        buffer.WriteByte Player(index).Hotbar(i).sType
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLoginOk(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    buffer.WriteLong index
    buffer.WriteLong Player_HighIndex
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHighIndex()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

' Check if have player on map
If PlayersOnMap(GetPlayerMap(index)) = NO Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    buffer.WriteByte 1
    buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(partyNum).Member(i)
    Next
    buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim buffer As clsBuffer, i As Long, partyNum As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(partyNum).Member(i)
        Next
        buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(index, i)
        buffer.WriteLong Player(index).Vital(i)
    Next
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal mapnum As Long, ByVal index As Long)
Dim buffer As clsBuffer
' Check if have player on map
If PlayersOnMap(mapnum) = NO Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnItem
    buffer.WriteLong index
    buffer.WriteString MapItem(mapnum, index).playerName
    buffer.WriteLong MapItem(mapnum, index).num
    buffer.WriteLong MapItem(mapnum, index).Value
    buffer.WriteLong MapItem(mapnum, index).x
    buffer.WriteLong MapItem(mapnum, index).y
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChatBubble(ByVal mapnum As Long, ByVal Target As Long, ByVal targetType As Long, ByVal message As String, ByVal Colour As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    buffer.WriteLong Target
    buffer.WriteLong targetType
    buffer.WriteString message
    buffer.WriteLong Colour
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpecialEffect(ByVal index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 'fognum
            buffer.WriteLong Data2 'fog movement speed
            buffer.WriteLong Data3 'opacity
        Case EFFECT_TYPE_WEATHER
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 'weather type
            buffer.WriteLong Data2 'weather intensity
        Case EFFECT_TYPE_TINT
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 'red
            buffer.WriteLong Data2 'green
            buffer.WriteLong Data3 'blue
            buffer.WriteLong Data4 'alpha
    End Select
    
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub


Sub SendAttack(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong ServerPackets.SAttack
    buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendFlash(ByVal Target As Long, mapnum As Long, isNpc As Boolean)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFlash
    buffer.WriteLong Target
    If isNpc Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendProjectile(ByVal mapnum As Long, ByVal attacker As Long, ByVal AttackerType As Long, ByVal victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    Call buffer.WriteLong(SCreateProjectile)
    Call buffer.WriteLong(attacker)
    Call buffer.WriteLong(AttackerType)
    Call buffer.WriteLong(victim)
    Call buffer.WriteLong(targetType)
    Call buffer.WriteLong(Graphic)
    Call buffer.WriteLong(Rotate)
    Call buffer.WriteLong(RotateSpeed)
    Call SendDataToMap(mapnum, buffer.ToArray())
    
    Set buffer = Nothing
End Sub

Sub SendClientTime()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte GameSecondsPerSecond
    buffer.WriteByte GameMinutesPerMinute
    buffer.WriteByte GameSeconds
    buffer.WriteByte GameMinutes
    buffer.WriteByte GameHours
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
End Sub
Sub SendClientTimeTo(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte GameSecondsPerSecond
    buffer.WriteByte GameMinutesPerMinute
    buffer.WriteByte GameSeconds
    buffer.WriteByte GameMinutes
    buffer.WriteByte GameHours
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
End Sub

Sub SendNewClass(ByVal index As Long, ByVal BaseClass As Byte, ByVal Advancement1 As Byte, ByVal Advancement2 As Byte, ByVal Advancement3 As Byte)
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SNewClass
    
    buffer.WriteByte BaseClass
    buffer.WriteByte Advancement1
    buffer.WriteByte Advancement2
    buffer.WriteByte Advancement3
    
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub

Sub SendBossMsg(ByVal message As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBossMsg
    buffer.WriteString message
    buffer.WriteLong Color
        
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRankUpdate(ByVal index As Long)

Dim i As Byte

Dim buffer As clsBuffer

Set buffer = New clsBuffer

buffer.WriteLong SRankUpdate

For i = 1 To MAX_RANK

buffer.WriteLong Rank(i).Level

buffer.WriteString Trim$(Rank(i).Name)

Next i

SendDataTo index, buffer.ToArray()

Set buffer = Nothing

End Sub

Public Sub SendNews(ByVal index As Long)

Dim News As String, f As Long
   Dim buffer As clsBuffer
   Set buffer = New clsBuffer
   
   buffer.WriteLong SSendNews
   f = FreeFile
   Open App.path & "\data\news.txt" For Input As #f
       Line Input #f, News
   Close #f
   buffer.WriteString News
   
   SendDataTo index, buffer.ToArray
   
   Set buffer = Nothing
End Sub

Public Sub SendPatchNotes(ByVal index As Long)
Dim Patchnotes As String, f As Long
   Dim buffer As clsBuffer
   Set buffer = New clsBuffer
   
   buffer.WriteLong SSendPatchNotes
   f = FreeFile
   Open App.path & "\data\patchnotes.txt" For Input As #f
       Line Input #f, Patchnotes
   Close #f
   buffer.WriteString Patchnotes
   
   SendDataTo index, buffer.ToArray
   
   Set buffer = Nothing
End Sub
