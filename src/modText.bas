Attribute VB_Name = "modText"
Option Explicit
' Stuffs
Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As DX8TextureRec
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single
'Text buffer

Public Type ChatTextBuffer
    text As String
    color As Long
End Type

'Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Const FVF_SIZE As Long = 28

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal x As Long, ByVal Y As Long, ByVal color As Long, Optional ByVal Alpha As Long = 0, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    ' set the color
    Alpha = 255 - Alpha
    color = dx8Colour(color, Alpha)
    
    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = color
    
    'Set the texture
    Direct3D_Device.SetTexture 0, gTexture(UseFont.Texture.Texture).Texture
    'CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(i))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_SIZE * 4)
                
                'Set up the verticies
                TempVA(0).x = x + count
                TempVA(0).Y = Y + yOffset
                TempVA(1).x = TempVA(1).x + x + count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).x = TempVA(0).x
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).x = TempVA(1).x
                TempVA(3).Y = TempVA(2).Y
                
                'Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                
                'Draw the verticies
                Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)))
                
                'Shift over the the position to render the next character
                count = count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next j
        End If
    Next i
End Sub

Sub EngineInitFontTextures()
    ' FONT DEFAULT
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Default.Texture.Texture = NumTextures
    Font_Default.Texture.filepath = App.Path & FONT_PATH & "texdefault.png"
    LoadTexture Font_Default.Texture
    
    ' Georgia
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Georgia.Texture.Texture = NumTextures
    Font_Georgia.Texture.filepath = App.Path & FONT_PATH & "georgia.png"
    LoadTexture Font_Georgia.Texture
End Sub

Sub UnloadFontTextures()
    UnloadFont Font_Default
    UnloadFont Font_Georgia
End Sub
Sub UnloadFont(Font As CustomFont)
    Font.Texture.Texture = 0
End Sub


Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal filename As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single


    'Load the header information
    FileNum = FreeFile
    Open App.Path & FONT_PATH & filename For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).x = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).x = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).x = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
    Next LoopChar
End Sub

Sub EngineInitFontSettings()
    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"
End Sub
Public Function dx8Colour(ByVal colourNum As Long, ByVal Alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
        Case 17 'Orange
            dx8Colour = D3DColorARGB(Alpha, 255, 96, 0)
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI

End Function

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim name As String
Dim Text2X As Long
Dim Text2Y As Long
Dim GuildString As String
Dim ThreatColor As Long
'Dim Text2Y As Long
Dim Text3X As Long
Dim LevelText As String
Dim LevelTextSize As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                color = White
            Case 1
                color = Blue
            Case 2
                color = Cyan
            Case 3
                color = BrightGreen
            Case 4
                color = Pink
        End Select

    Else
        color = BrightRed
    End If
    
    If GetPlayerLevel(Index) <= GetPlayerLevel(MyIndex) - 10 Then
        ThreatColor = Grey
    ElseIf GetPlayerLevel(Index) <= GetPlayerLevel(MyIndex) - 5 Then
        ThreatColor = Green
    ElseIf GetPlayerLevel(Index) <= GetPlayerLevel(MyIndex) - 1 Then
        ThreatColor = Cyan
    ElseIf GetPlayerLevel(Index) >= GetPlayerLevel(MyIndex) + 1 Then
        ThreatColor = Yellow
    ElseIf GetPlayerLevel(Index) >= GetPlayerLevel(MyIndex) + 5 Then
        ThreatColor = D3DColorARGB(255, 150, 150, 0)
    ElseIf GetPlayerLevel(Index) >= GetPlayerLevel(MyIndex) + 10 Then
        ThreatColor = BrightRed
    Else
        ThreatColor = White
    End If

    LevelText = "Lv" & GetPlayerLevel(Index) & " " & Trim$(Class(GetPlayerClass(Index)).name)
    LevelTextSize = EngineGetTextWidth(Font_Default, LevelText)
    name = Trim$(Player(Index).name)
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    Text3X = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(LevelText))) / 2)
    GuildString = Player(Index).GuildTag
    Text2X = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(GuildString))) / 2)
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - 16
        Text2Y = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (Tex_Character(GetPlayerSprite(Index)).Height / 4) + 16
        Text2Y = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (Tex_Character(GetPlayerSprite(Index)).Height / 4) + 4
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Font_Default, LevelText, Text3X, Text2Y, ThreatColor
    RenderText Font_Default, name, TextX, TextY, color, 0
    If Not Player(Index).GuildName = vbNullString Then
        'Call DrawText(TexthDC, Text2X, Text2Y, GuildString, Color)
        RenderText Font_Default, GuildString, Text2X, Text2Y - 12, Player(Index).GuildColor, 0
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim TextX As Long, Text2X As Long
Dim TextY As Long, Text2Y As Long
Dim color As Long
Dim name As String, NameSize As Long
Dim npcNum As Long
Dim ThreatColor As Long
Dim Level As String, LevelSize As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    npcNum = MapNpc(Index).num
    
    If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
        ' get the colour
        If NPC(npcNum).Level <= GetPlayerLevel(MyIndex) - 10 Then
            ThreatColor = Grey
        ElseIf NPC(npcNum).Level <= GetPlayerLevel(MyIndex) - 5 Then
            ThreatColor = Green
        ElseIf NPC(npcNum).Level <= GetPlayerLevel(MyIndex) - 1 Then
            ThreatColor = Cyan
        ElseIf NPC(npcNum).Level >= GetPlayerLevel(MyIndex) + 1 Then
            ThreatColor = Yellow
        ElseIf NPC(npcNum).Level >= GetPlayerLevel(MyIndex) + 5 Then
            ThreatColor = D3DColorARGB(255, 150, 150, 0)
        ElseIf NPC(npcNum).Level >= GetPlayerLevel(MyIndex) + 10 Then
            ThreatColor = BrightRed
        Else
            ThreatColor = White
        End If
    Else
        ThreatColor = White
    End If

    Select Case NPC(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            color = Grey
        Case Else
            color = BrightGreen
    End Select

    name = Trim$(NPC(npcNum).name)
    NameSize = EngineGetTextWidth(Font_Default, name)
    Level = "Level " & Trim$(NPC(npcNum).Level)
    LevelSize = EngineGetTextWidth(Font_Default, Level)
    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - (NameSize / 2)
    Text2X = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - (LevelSize / 2)
    If NPC(npcNum).Sprite < 1 Or NPC(npcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - 12
        Text2Y = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4) + 16
        Text2Y = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4) + 4
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
        RenderText Font_Default, Level, Text2X, Text2Y, ThreatColor
        RenderText Font_Default, name, TextX, TextY, color, 0
    Else
    RenderText Font_Default, name, TextX, TextY, color, 0
    End If
    
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        'check if the npc is the next task to any quest: [?] symbol
        If Quest(i).name <> "" Then
            If Player(MyIndex).PlayerQuest(i).status = QUEST_STARTED Then
                If Quest(i).Task(Player(MyIndex).PlayerQuest(i).ActualTask).NPC = npcNum Then
                    name = "[?]"
                    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - (EngineGetTextWidth(Font_Default, name) / 2)
                    TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - 16
                    If NPC(npcNum).Sprite >= 1 And NPC(npcNum).Sprite <= NumCharacters Then
                        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4)
                    End If
                    If Not NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        Call RenderText(Font_Default, name, TextX, TextY - 12, Yellow)
                    Else
                        Call RenderText(Font_Default, name, TextX, TextY, Yellow)
                    End If
                    Exit For
                End If
            End If
            
            'check if the npc is the starter to any quest: [!] symbol
            'can accept the quest as a new one?
            If Player(MyIndex).PlayerQuest(i).status = QUEST_NOT_STARTED Or Player(MyIndex).PlayerQuest(i).status = QUEST_COMPLETED_BUT Then
                'the npc gives this quest?
                If NPC(npcNum).Quest = 1 Then
                    name = "[!]"
                    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - (EngineGetTextWidth(Font_Default, name) / 2)
                    TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - 16
                    If NPC(npcNum).Sprite >= 1 And NPC(npcNum).Sprite <= NumCharacters Then
                        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(NPC(npcNum).Sprite).Height / 4)
                    End If
                    If Not NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        Call RenderText(Font_Default, name, TextX, TextY - 12, Yellow)
                    Else
                        Call RenderText(Font_Default, name, TextX, TextY, Yellow)
                    End If
                    Exit For
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim Y As Long
    Dim tx As Long
    Dim ty As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.value Then
        For x = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, Y) Then
                    With Map.Tile(x, Y)
                        tx = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Font_Default, "B", tx, ty, BrightRed, 0
                            Case TILE_TYPE_WARP
                                RenderText Font_Default, "W", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_ITEM
                                RenderText Font_Default, "I", tx, ty, White, 0
                            Case TILE_TYPE_NPCAVOID
                                RenderText Font_Default, "N", tx, ty, White, 0
                            Case TILE_TYPE_KEY
                                RenderText Font_Default, "K", tx, ty, White, 0
                            Case TILE_TYPE_KEYOPEN
                                RenderText Font_Default, "O", tx, ty, White, 0
                            Case TILE_TYPE_RESOURCE
                                RenderText Font_Default, "B", tx, ty, Green, 0
                            Case TILE_TYPE_DOOR
                                RenderText Font_Default, "D", tx, ty, Brown, 0
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Font_Default, "S", tx, ty, Yellow, 0
                            Case TILE_TYPE_SHOP
                                RenderText Font_Default, "S", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_BANK
                                RenderText Font_Default, "B", tx, ty, Blue, 0
                            Case TILE_TYPE_HEAL
                                RenderText Font_Default, "H", tx, ty, BrightGreen, 0
                            Case TILE_TYPE_TRAP
                                RenderText Font_Default, "T", tx, ty, BrightRed, 0
                            Case TILE_TYPE_SLIDE
                                RenderText Font_Default, "S", tx, ty, BrightCyan, 0
                            Case TILE_TYPE_SOUND
                                RenderText Font_Default, "S", tx, ty, Orange, 0
                            Case TILE_TYPE_LIGHT
                                RenderText Font_Default, "L", tx, ty, Yellow
                            Case TILE_TYPE_CRAFT
                                RenderText Font_Default, "CR", tx, ty, Orange
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawActionMsg(ByVal Index As Long)
    Dim x As Long, Y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).Y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).Y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMain.ScaleWidth \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
            Y = 425
            
            Case ACTIONMSG_BOSS
            Time = 15000
            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_BOSS Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMain.ScaleWidth \ 2) - (EngineGetTextWidth(Font_Default, Trim$(ActionMsg(Index).Message)) / 2)
            Y = 114
            RenderTextureRectangle -2, 107, frmMain.ScaleWidth + 4, 28

    End Select
    
    x = ConvertMapX(x)
    Y = ConvertMapY(Y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        RenderText Font_Default, ActionMsg(Index).Message, x, Y, ActionMsg(Index).color, 0
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(Font As CustomFont, ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    getWidth = EngineGetTextWidth(Font, text)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub DrawEventName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If InMapEditor Then Exit Sub

    color = White

    name = Trim$(Map.MapEvents(Index).name)
    
    ' calc pos
    TextX = ConvertMapX(Map.MapEvents(Index).x * PIC_X) + Map.MapEvents(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    If Map.MapEvents(Index).GraphicType = 0 Then
        TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
    ElseIf Map.MapEvents(Index).GraphicType = 1 Then
        If Map.MapEvents(Index).GraphicNum < 1 Or Map.MapEvents(Index).GraphicNum > NumCharacters Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
        Else
            ' Determine location for text
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - (Tex_Character(Map.MapEvents(Index).GraphicNum).Height / 4) + 16
        End If
    ElseIf Map.MapEvents(Index).GraphicType = 2 Then
        If Map.MapEvents(Index).GraphicY2 > 0 Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - ((Map.MapEvents(Index).GraphicY2 - Map.MapEvents(Index).GraphicY) * 32) + 16
        Else
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 32 + 16
        End If
    End If

    ' Draw name
    RenderText Font_Default, name, TextX, TextY, color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String, x As Long, Y As Long, i As Long, MaxWidth As Long, x2 As Long, y2 As Long, colour As Long
    
    With chatBubble(Index)
        If .TargetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' it's on our map - get co-ords
                x = ConvertMapX((Player(.target).x * 32) + Player(.target).xOffset) + 16
                Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).yOffset) - 40
            End If
        ElseIf .TargetType = TARGET_TYPE_NPC Then
            ' it's on our map - get co-ords
            x = ConvertMapX((MapNpc(.target).x * 32) + MapNpc(.target).xOffset) + 16
            Y = ConvertMapY((MapNpc(.target).Y * 32) + MapNpc(.target).yOffset) - 40
        ElseIf .TargetType = TARGET_TYPE_EVENT Then
            x = ConvertMapX((Map.MapEvents(.target).x * 32) + Map.MapEvents(.target).xOffset) + 16
            Y = ConvertMapY((Map.MapEvents(.target).Y * 32) + Map.MapEvents(.target).yOffset) - 40
        End If
        
        ' word wrap the text
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
        ' find max width
        For i = 1 To UBound(theArray)
            If EngineGetTextWidth(Font_Default, theArray(i)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(i))
        Next
                
        ' calculate the new position
        x2 = x - (MaxWidth \ 2)
        y2 = Y - (UBound(theArray) * 12)
                
        ' render bubble - top left
        RenderTexture Tex_GUI(25), x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_GUI(25), x2 + MaxWidth, y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_GUI(25), x2, y2 - 5, 10, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_GUI(25), x2 - 9, Y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture Tex_GUI(25), x2 + MaxWidth, Y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture Tex_GUI(25), x2, Y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_GUI(25), x2 + (MaxWidth \ 2) + 6, Y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_GUI(25), x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_GUI(25), x2 + MaxWidth, y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_GUI(25), x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_GUI(25), x - 5, Y, 58, 19, 11, 11, 11, 11
                
        ' render each line centralised
        For i = 1 To UBound(theArray)
            RenderText Font_Georgia, theArray(i), x - (EngineGetTextWidth(Font_Default, theArray(i)) / 2), y2, DarkBrown
            y2 = y2 + 12
        Next
        ' check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

' Chat Box
Public Sub RenderChatTextBuffer()
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim i As Long

    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    Direct3D_Device.SetTexture 0, gTexture(Font_Default.Texture.Texture).Texture

    If ChatArrayUbound > 0 Then
        Direct3D_Device.SetStreamSource 0, ChatVBS, FVF_SIZE
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        Direct3D_Device.SetStreamSource 0, ChatVB, FVF_SIZE
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
    
End Sub

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim v As Single
Dim x As Single
Dim Y As Single
Dim y2 As Single
Dim i As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    ' set the offset of each line
    yOffset = 14

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    
    Chunk = ChatScroll
    
    'Get the number of characters in all the visible buffer
    Size = 0
    
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        Size = Size + Len(ChatTextBuffer(LoopC).text)
    Next
    
    Size = Size - j
    ChatArrayUbound = Size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    'Set the base position
    x = GUIWindow(GUI_CHAT).x + ChatOffsetX
    Y = GUIWindow(GUI_CHAT).Y + ChatOffsetY

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        'Set the temp color
        TempColor = ChatTextBuffer(LoopC).color
        
        'Set the Y position to be used
        y2 = Y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        
        'Loop through each line if there are line breaks (vbCrLf)
        count = 0   'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).text) <> 0 Then  'Dont bother with empty strings
            
            'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).text)
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).text, j, 1))
                
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                v = Row * Font_Default.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * Pos))
                    .color = TempColor
                    .x = (x) + count
                    .Y = (y2)
                    .TU = u
                    .TV = v
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * Pos))
                    .color = TempColor
                    .x = (x) + count
                    .Y = (y2) + Font_Default.HeaderInfo.CellHeight
                    .TU = u
                    .TV = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * Pos))
                    .color = TempColor
                    .x = (x) + count + Font_Default.HeaderInfo.CellWidth
                    .Y = (y2) + Font_Default.HeaderInfo.CellHeight
                    .TU = u + Font_Default.ColFactor
                    .TV = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                
                'Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * Pos)) = ChatVA(0 + (6 * Pos)) 'Top-left corner
                
                ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * Pos))
                    .color = TempColor
                    .x = (x) + count + Font_Default.HeaderInfo.CellWidth
                    .Y = (y2)
                    .TU = u + Font_Default.ColFactor
                    .TV = v
                    .RHW = 1
                End With

                ChatVA(5 + (6 * Pos)) = ChatVA(2 + (6 * Pos))

                'Update the character we are on
                Pos = Pos + 1

                'Shift over the the position to render the next character
                count = count + Font_Default.HeaderInfo.CharWidth(Ascii)
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).color
                End If
            Next
        End If
    Next LoopC
        
    If Not Direct3D_Device Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = Direct3D_Device.CreateVertexBuffer(FVF_SIZE * Pos * 6, 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_SIZE * Pos * 6, 0, ChatVAS(0)
        Set ChatVB = Direct3D_Device.CreateVertexBuffer(FVF_SIZE * Pos * 6, 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_SIZE * Pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
    
End Sub

Public Sub AddText(ByVal text As String, ByVal tColor As Long, Optional ByVal Alpha As Long = 255)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim B As Long
Dim color As Long

    color = dx8Colour(tColor, Alpha)

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
            
            'Add up the size
            Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            
            'Check for too large of a size
            If Size > ChatWidth Then
                
                'Check if the last space was too far back
                If i - lastSpace > 10 Then
                
                    'Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)), color
                    B = i - 1
                    Size = 0
                Else
                    'Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)), color
                    B = lastSpace + 1
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
            
            'This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If B <> i Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), B, i), color
            End If
        Next i
    Next TSLoop
    
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_Default.RowPitch = 0 Then Exit Sub
    
    If ChatScroll > 8 Then ChatScroll = ChatScroll + 1

    'Update the array
    UpdateChatArray
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal text As String, ByVal color As Long)
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).text = text
    ChatTextBuffer(1).color = color
    
    ' set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
End Sub
Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, Size As Long, lastSpace As Long, B As Long
    
    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If
    
    ' default values
    B = 1
    lastSpace = 1
    Size = 0
    
    For i = 1 To Len(text)
        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
        
        'Add up the size
        Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(text, i, 1)))
        
        'Check for too large of a size
        If Size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (i - 1) - B))
                B = i - 1
                Size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = EngineGetTextWidth(Font_Default, Mid$(text, lastSpace, i - lastSpace))
            End If
        End If
        
        ' Remainder
        If i = Len(text) Then
            If B <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim B As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": lastSpace = i
                    Case "_": lastSpace = i
                    Case "-": lastSpace = i
                End Select
    
                'Add up the size
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
 
                'Check for too large of a size
                If Size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)) & vbNewLine
                        B = i - 1
                        Size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If B <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function
 

Public Sub UpdateShowChatText()
Dim CHATOFFSET As Long, i As Long, x As Long

    CHATOFFSET = 52
    
    If EngineGetTextWidth(Font_Default, MyText) > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
        For i = Len(MyText) To 1 Step -1
            x = x + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(MyText, i, 1)))
            If x > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
                RenderChatText = Right$(MyText, Len(MyText) - i + 1)
                Exit For
            End If
        Next
    Else
        RenderChatText = MyText
    End If
End Sub

Sub DrawBossMsg()
    Dim x As Long, Y As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If BossMsg.Created = 0 Then Exit Sub
    
    Time = 15000
    x = (frmMain.ScaleWidth \ 2) - (EngineGetTextWidth(Font_Default, Trim$(BossMsg.Message)) / 2)
    Y = 114
    
    If GetTickCount < BossMsg.Created + Time Then
        RenderTextureRectangle -2, 107, frmMain.ScaleWidth + 4, 28
        RenderText Font_Default, Trim$(BossMsg.Message), x, Y, BossMsg.color
    Else
        BossMsg.Message = vbNullString
        BossMsg.Created = 0
        BossMsg.color = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBossMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

