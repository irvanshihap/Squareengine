Attribute VB_Name = "modGraphics"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Private Direct3DX As D3DX8

'The 2D (Transformed and Lit) vertex format.
Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

Public ScreenWidth As Long
Public ScreenHeight As Long

'Graphic Textures
Public Tex_GUI() As DX8TextureRec
Public Tex_Buttons() As DX8TextureRec
Public Tex_Buttons_h() As DX8TextureRec
Public Tex_Buttons_c() As DX8TextureRec
Public Tex_Item() As DX8TextureRec ' arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Blood As DX8TextureRec ' singes
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Shadow As DX8TextureRec
Public Tex_Projectile() As DX8TextureRec
Public Tex_Lightmap As DX8TextureRec
Public Tex_Light As DX8TextureRec
Public Tex_Panorama() As DX8TextureRec

' Number of graphic files
Public NumGUIs As Long
Public NumButtons As Long
Public NumButtons_c As Long
Public NumButtons_h As Long
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumProjectiles As Long
Public NumPanoramas As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    Loaded As Boolean
    UnloadTimer As Long
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    ScreenWidth = 800
    ScreenHeight = 600
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = ScreenWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = ScreenHeight 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
    
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    ' Initialise the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDX8", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TryCreateDirectX8Device() As Boolean
Dim I As Long

On Error GoTo nexti

    For I = 1 To 4
        Select Case I
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            'Case 2
            '    Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
            '    TryCreateDirectX8Device = True
            '    Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(Value As Long) As Long
Dim I As Long
    Do While 2 ^ I < Value
        I = I + 1
    Loop
    GetNearestPOT = 2 ^ I
End Function
Public Sub SetTexture(ByRef TextureRec As DX8TextureRec)
If TextureRec.Texture > NumTextures Then TextureRec.Texture = NumTextures
If TextureRec.Texture < 0 Then TextureRec.Texture = 0

If Not TextureRec.Texture = 0 Then
    If Not gTexture(TextureRec.Texture).Loaded Then
        Call LoadTexture(TextureRec)
    End If
End If

End Sub
Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, I As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            I = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, I, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(I)
            I = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, I)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (I)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    gTexture(TextureRec.Texture).Loaded = True
    gTexture(TextureRec.Texture).UnloadTimer = GetTickCount
    Exit Sub
errorhandler:
    HandleError "LoadTexture", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub LoadTextures()
Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckGUIs
    Call CheckButtons
    Call CheckButtons_c
    Call CheckButtons_h
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckProjectiles
    Call CheckPanoramas
    
    NumTextures = NumTextures + 12
    
    ReDim Preserve gTexture(NumTextures)
    Tex_Light.filepath = App.Path & "\data files\graphics\misc\light.png"
    Tex_Light.Texture = NumTextures - 11
    Tex_Lightmap.filepath = App.Path & "\data files\graphics\misc\lightmap.png"
    Tex_Lightmap.Texture = NumTextures - 10
    Tex_Shadow.filepath = App.Path & "\data files\graphics\misc\shadow.png"
    Tex_Shadow.Texture = NumTextures - 9
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader.png"
    Tex_Fade.Texture = NumTextures - 8
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather.png"
    Tex_Weather.Texture = NumTextures - 7
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white.png"
    Tex_White.Texture = NumTextures - 6
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction.png"
    Tex_Direction.Texture = NumTextures - 5
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target.png"
    Tex_Target.Texture = NumTextures - 4
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc.png"
    Tex_Misc.Texture = NumTextures - 3
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood.png"
    Tex_Blood.Texture = NumTextures - 2
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select.png"
    Tex_Selection.Texture = NumTextures
    
    EngineInitFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures(Optional ByVal Complete As Boolean = False)
Dim I As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    If Complete = False Then
        For I = 1 To NumTextures
            If gTexture(I).UnloadTimer > GetTickCount + 150000 Then
                Set gTexture(I).Texture = Nothing
                ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
                gTexture(I).UnloadTimer = 0
                gTexture(I).Loaded = False
            End If
        Next
    Else
    
    For I = 1 To NumTextures
        Set gTexture(I).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
    Next
    
    ReDim gTexture(1)

    
    For I = 1 To NumTileSets
        Tex_Tileset(I).Texture = 0
    Next

    For I = 1 To numitems
        Tex_Item(I).Texture = 0
    Next

    For I = 1 To NumCharacters
        Tex_Character(I).Texture = 0
    Next
    
    For I = 1 To NumPaperdolls
        Tex_Paperdoll(I).Texture = 0
    Next
    
    For I = 1 To NumResources
        Tex_Resource(I).Texture = 0
    Next
    
    For I = 1 To NumAnimations
        Tex_Animation(I).Texture = 0
    Next
    
    For I = 1 To NumSpellIcons
        Tex_SpellIcon(I).Texture = 0
    Next
    
    For I = 1 To NumFaces
        Tex_Face(I).Texture = 0
    Next
    
    For I = 1 To NumGUIs
        Tex_GUI(I).Texture = 0
    Next
    
    For I = 1 To NumButtons
        Tex_Buttons(I).Texture = 0
    Next
    
    For I = 1 To NumButtons_c
        Tex_Buttons_c(I).Texture = 0
    Next
    
    For I = 1 To NumButtons_h
        Tex_Buttons_h(I).Texture = 0
    Next
    
    For I = 1 To NumProjectiles
        Tex_Projectile(I).Texture = 0
    Next
    
    For I = 1 To NumPanoramas
        Tex_Panorama(I).Texture = 0
    Next
    
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    Tex_Bars.Texture = 0
    Tex_White.Texture = 0
    Tex_Weather.Texture = 0
    Tex_Fade.Texture = 0
    Tex_Shadow.Texture = 0
    
    UnloadFontTextures
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sx As Single, ByVal sy As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional color As Long = -1, Optional ByVal Degrees As Single = 0)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    Dim RadAngle As Single 'The angle in Radians
    Dim CenterX As Single
    Dim CenterY As Single
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single
    Dim I As Long
    
    SetTexture TextureRec
    
    TextureNum = TextureRec.Texture
    
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    
    If sy + sHeight > textureHeight Then Exit Sub
    If sx + sWidth > textureWidth Then Exit Sub
    If sx < 0 Then Exit Sub
    If sy < 0 Then Exit Sub

    sx = sx - 0.5
    sy = sy - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sx / textureWidth)
    sourceY = (sy / textureHeight)
    sourceWidth = ((sx + sWidth) / textureWidth)
    sourceHeight = ((sy + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    'Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = dX + (dWidth * 0.5)
        CenterY = dY + (dHeight * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For I = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Vertex_List(I).X - CenterX) * CosRad - (Vertex_List(I).Y - CenterY) * SinRad
            NewY = CenterY + (Vertex_List(I).Y - CenterY) * CosRad + (Vertex_List(I).X - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            Vertex_List(I).X = NewX
            Vertex_List(I).Y = NewY
        Next
    End If
    
    Call Direct3D_Device.SetTexture(0, gTexture(TextureNum).Texture)
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRECT As RECT, drect As RECT, Optional colour As Long = -1)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture TextureRec, drect.Left, drect.Top, sRECT.Left, sRECT.Top, drect.Right - drect.Left, drect.Bottom - drect.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, colour

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
Dim Rec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    Rec.Top = 24
    Rec.Left = 0
    Rec.Right = Rec.Left + 32
    Rec.Bottom = Rec.Top + 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' render dir blobs
    For I = 1 To 4
        Rec.Left = (I - 1) * 8
        Rec.Right = Rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(I)) Then
            Rec.Top = 8
        Else
            Rec.Top = 16
        End If
        Rec.Bottom = Rec.Top + 8
        'render!
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(I), ConvertMapY(Y * PIC_Y) + DirArrowY(I), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDirection", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    X = X - ((Width - 32) / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If
    
    If Targettmr < GetTickCount Then
        If Curtarget = 1 Then
            Curtarget = 0
        Else
            Curtarget = Curtarget + 1
        End If
        Targettmr = GetTickCount + 200
    End If
    
    Select Case Curtarget
        Case 0
            RenderTexture Tex_Target, X, Y - 2, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, (sRECT.Bottom - sRECT.Top) + 4, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        Case 1
            RenderTexture Tex_Target, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
    End Select
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTarget", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Tex_Target.Width / 2
        .Right = .Left + Width
    End With
    
    X = X - ((Width - 32) / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Target, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHover", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim Rec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For I = MapLayer.Ground To MapLayer.Mask2
            If Autotile(X, Y).layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .layer(I).X * 32, .layer(I).Y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(X, Y).layer(I).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "DrawMapTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim Rec As RECT
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For I = MapLayer.Fringe To MapLayer.Fringe2
            If Autotile(X, Y).layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .layer(I).X * 32, .layer(I).Y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(X, Y).layer(I).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile I, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile I, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapFringeTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim Rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    'load blood then
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        ' check if we should be seeing it
        If .timer + 20000 < GetTickCount Then Exit Sub
        
        Rec.Top = 0
        Rec.Bottom = PIC_Y
        Rec.Left = (.Sprite - 1) * PIC_X
        Rec.Right = Rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBlood", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal layer As Long)
Dim Sprite As Integer, sRECT As RECT, I As Long, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim X As Long, Y As Long, lockindex As Long
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(layer)
    
    ' total width divided by frame count
    Width = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).width / frameCount
    Height = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).height
    
    With sRECT
        .Top = (Height * ((AnimInstance(Index).frameIndex(layer) - 1) \ AnimColumns))
        .Bottom = .Top + Height
        .Left = (Width * (((AnimInstance(Index).frameIndex(layer) - 1) Mod AnimColumns)))
        .Right = .Left + Width
    End With
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Animation(Sprite), X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
End Sub
Public Sub ScreenshotMap()
Dim X As Long, Y As Long, I As Long, Rec As RECT, drec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picSSMap.Cls
    
    ' render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For I = MapLayer.Ground To MapLayer.Mask2
                    ' skip tile?
                    If (.layer(I).Tileset > 0 And .layer(I).Tileset <= NumTileSets) And (.layer(I).X > 0 Or .layer(I).Y > 0) Then
                        ' sort out rec
                        Rec.Top = .layer(I).Y * PIC_Y
                        Rec.Bottom = Rec.Top + PIC_Y
                        Rec.Left = .layer(I).X * PIC_X
                        Rec.Right = Rec.Left + PIC_X
                        
                        drec.Left = X * PIC_X
                        drec.Top = Y * PIC_Y
                        drec.Right = drec.Left + (Rec.Right - Rec.Left)
                        drec.Bottom = drec.Top + (Rec.Bottom - Rec.Top)
                        ' render
                        RenderTextureByRects Tex_Tileset(.layer(I).Tileset), Rec, drec
                    End If
                Next
            End With
        Next
    Next
    
    ' render the resources
    For Y = 0 To Map.MaxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For I = 1 To Resource_Index
                        If MapResource(I).Y = Y Then
                            Call DrawMapResource(I, True)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For I = MapLayer.Fringe To MapLayer.Fringe2
                    ' skip tile?
                    If (.layer(I).Tileset > 0 And .layer(I).Tileset <= NumTileSets) And (.layer(I).X > 0 Or .layer(I).Y > 0) Then
                        ' sort out rec
                        Rec.Top = .layer(I).Y * PIC_Y
                        Rec.Bottom = Rec.Top + PIC_Y
                        Rec.Left = .layer(I).X * PIC_X
                        Rec.Right = Rec.Left + PIC_X
                        
                        drec.Left = X * PIC_X
                        drec.Top = Y * PIC_Y
                        drec.Right = drec.Left + (Rec.Right - Rec.Left)
                        drec.Bottom = drec.Top + (Rec.Bottom - Rec.Top)
                        ' render
                        RenderTextureByRects Tex_Tileset(.layer(I).Tileset), Rec, drec
                    End If
                Next
            End With
        Next
    Next
    
    ' dump and save
    frmMain.picSSMap.Width = (Map.MaxX + 1) * 32
    frmMain.picSSMap.Height = (Map.MaxY + 1) * 32
    Rec.Top = 0
    Rec.Left = 0
    Rec.Bottom = (Map.MaxX + 1) * 32
    Rec.Right = (Map.MaxY + 1) * 32
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"
    
    ' let them know we did it
    AddText "Screenshot of map #" & GetPlayerMap(MyIndex) & " saved.", BrightGreen
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotMap", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim Rec As RECT
Dim X As Long, Y As Long
Dim I As Long, Alpha As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With Rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    

    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
            If ConvertMapY(GetPlayerY(I)) < ConvertMapY(MapResource(Resource_num).Y) And ConvertMapY(GetPlayerY(I)) > ConvertMapY(MapResource(Resource_num).Y) - (Tex_Resource(Resource_sprite).Height) / 32 Then
                If ConvertMapX(GetPlayerX(I)) >= ConvertMapX(MapResource(Resource_num).X) - ((Tex_Resource(Resource_sprite).Width / 2) / 32) And ConvertMapX(GetPlayerX(I)) <= ConvertMapX(MapResource(Resource_num).X) + ((Tex_Resource(Resource_sprite).Width / 2) / 32) Then
                    Alpha = 150
                Else
                    Alpha = 255
                End If
            Else
                Alpha = 255
            End If
        End If
    Next

    
    ' render it
    If Not screenShot Then
        Call DrawResource(Resource_sprite, Alpha, X, Y, Rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, Rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal Alpha As Long, ByVal dX As Long, dY As Long, Rec As RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    X = ConvertMapX(dX)
    Y = ConvertMapY(dY)
    
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
    
    RenderTexture Tex_Resource(Resource), X, Y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, Alpha)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, Rec As RECT)
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)

    If Y < 0 Then
        With Rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With Rec
            .Left = .Left - X
        End With
        X = 0
    End If
    RenderTexture Tex_Resource(Resource), X, Y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim I As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SetTexture Tex_Bars
    ' dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 4
    
    ' render health bars
    For I = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(I).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(I).Vital(Vitals.HP) > 0 And MapNpc(I).Vital(Vitals.HP) < NPC(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(I).X * PIC_X + MapNpc(I).xOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(I).Y * PIC_Y + MapNpc(I).yOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(I).Vital(Vitals.HP) / sWidth) / (NPC(npcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                
                ' draw the bar proper
                With sRECT
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer).Spell).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            
            ' draw the bar proper
            With sRECT
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRECT
            .Top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
       
        ' draw the bar proper
        With sRECT
            .Top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For I = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(I)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).xOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).yOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        Next
    End If
                    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBars", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayer(ByVal Index As Long, Optional a As Byte = 255, Optional R As Byte = 255, Optional G As Byte = 255, Optional B As Byte = 255)
Dim Anim As Byte, I As Long, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long
Dim Rec As RECT
Dim AttackSpeed As Long

    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)
    

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
        AttackSpeed = Item(GetPlayerEquipment(Index, weapon)).speed
    Else
        AttackSpeed = 1000
    End If
    

    If VXFRAME = False Then
        ' Reset frame
        If Player(Index).Step = 3 Then
            Anim = 0
        ElseIf Player(Index).Step = 1 Then
            Anim = 2
        End If
    Else
        Anim = 1
    End If
    
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            If VXFRAME = False Then
                Anim = 3
            Else
                Anim = 2
            End If
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    If isAnimated(GetPlayerSprite(Index)) Then
    If Player(Index).AnimTimer + 100 <= GetTickCount Then
    Player(Index).Anim = Player(Index).Anim + 1
    If Player(Index).Anim >= 3 Then Player(Index).Anim = 0
    Player(Index).AnimTimer = GetTickCount
    End If
    Anim = Player(Index).Anim
    End If

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With Rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        If VXFRAME = False Then
            .Left = Anim * (Tex_Character(Sprite).Width / 4)
            .Right = .Left + (Tex_Character(Sprite).Width / 4)
        Else
            .Left = Anim * (Tex_Character(Sprite).Width / 3)
            .Right = .Left + (Tex_Character(Sprite).Width / 3)
        End If
    End With

    ' Calculate the X
    If VXFRAME = False Then
        X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    Else
        X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    End If
    
    ' render player shadow
    'If hasShadow(Sprite) Then RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 150)
    'RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y + 4), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 150)
    
    'If StealthDuration > 0 Then
      '  Call DrawSprite(Sprite, x, y, rec, 100, 255, 255, 255, True)
      'SendVisibility
    'Else
    
     ' render the actual sprite
    'If GetTickCount > Player(Index).StartFlash Then
       ' Call DrawSprite(Sprite, x, y, rec)
      '  Player(Index).StartFlash = 0
    'Else
    
      '  Call DrawSprite(Sprite, x, y, rec, True)
    'End If
    'End If
    
    ' check for paperdolling
    'For I = 1 To UBound(PaperdollOrder)
       ' If GetPlayerEquipment(Index, PaperdollOrder(I)) > 0 Then
           ' If Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll > 0 Then
                'Call DrawPaperdoll(x, y, Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll, anim, spritetop, 255, 255, 255, 255)
            'End If
        'End If
    'Next

    
    Select Case GetPlayerDir(Index)
    
    
    Case DIR_DOWN
    
    
    If GetTickCount > Player(Index).StartFlash Then
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B)
        Player(Index).StartFlash = 0
    Else
    
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B, True)
    End If
    
    
    'Helmet
    If GetPlayerEquipment(Index, helmet) > 0 Then
    
    If Item(GetPlayerEquipment(Index, helmet)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, helmet)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, helmet)).R, 255 - Item(GetPlayerEquipment(Index, helmet)).G, 255 - Item(GetPlayerEquipment(Index, helmet)).B)
       
    End If
    End If
    
    'Charm
    If GetPlayerEquipment(Index, charm) > 0 Then
    
    If Item(GetPlayerEquipment(Index, charm)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, charm)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, charm)).R, 255 - Item(GetPlayerEquipment(Index, charm)).G, 255 - Item(GetPlayerEquipment(Index, charm)).B)
       
    End If
    End If
    
    'Boots
    If GetPlayerEquipment(Index, boots) > 0 Then
    
    If Item(GetPlayerEquipment(Index, boots)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, boots)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, boots)).R, 255 - Item(GetPlayerEquipment(Index, boots)).G, 255 - Item(GetPlayerEquipment(Index, boots)).B)
       
    End If
    End If
    
    'Armor
    If GetPlayerEquipment(Index, armor) > 0 Then
    
    If Item(GetPlayerEquipment(Index, armor)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, armor)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, armor)).R, 255 - Item(GetPlayerEquipment(Index, armor)).G, 255 - Item(GetPlayerEquipment(Index, armor)).B)
       
    End If
    End If
    
    
    'Whetstone
    If GetPlayerEquipment(Index, whetstone) > 0 Then
    
    If Item(GetPlayerEquipment(Index, whetstone)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, whetstone)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, whetstone)).R, 255 - Item(GetPlayerEquipment(Index, whetstone)).G, 255 - Item(GetPlayerEquipment(Index, whetstone)).B)
       
    End If
    End If
    
    
    'Ring
    If GetPlayerEquipment(Index, ring) > 0 Then
    
    If Item(GetPlayerEquipment(Index, ring)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, ring)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, ring)).R, 255 - Item(GetPlayerEquipment(Index, ring)).G, 255 - Item(GetPlayerEquipment(Index, ring)).B)
       
    End If
    End If
    
    'Enchant
    If GetPlayerEquipment(Index, enchant) > 0 Then
    
    If Item(GetPlayerEquipment(Index, enchant)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, enchant)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, enchant)).R, 255 - Item(GetPlayerEquipment(Index, enchant)).G, 255 - Item(GetPlayerEquipment(Index, enchant)).B)
       
    End If
    End If
    
    'Shield
    If GetPlayerEquipment(Index, shield) > 0 Then
    
    If Item(GetPlayerEquipment(Index, shield)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, shield)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, shield)).R, 255 - Item(GetPlayerEquipment(Index, shield)).G, 255 - Item(GetPlayerEquipment(Index, shield)).B)
       
    End If
    End If
    
    'Weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
    
    If Item(GetPlayerEquipment(Index, weapon)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, weapon)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, weapon)).R, 255 - Item(GetPlayerEquipment(Index, weapon)).G, 255 - Item(GetPlayerEquipment(Index, weapon)).B)
       
    End If
    End If
    
    Case DIR_UP
    
    'Shield
    If GetPlayerEquipment(Index, shield) > 0 Then
    
    If Item(GetPlayerEquipment(Index, shield)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, shield)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, shield)).R, 255 - Item(GetPlayerEquipment(Index, shield)).G, 255 - Item(GetPlayerEquipment(Index, shield)).B)
    End If
    End If
    
    'Weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
    
    If Item(GetPlayerEquipment(Index, weapon)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, weapon)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, weapon)).R, 255 - Item(GetPlayerEquipment(Index, weapon)).G, 255 - Item(GetPlayerEquipment(Index, weapon)).B)
       
    End If
    End If
    
    
    
    If GetTickCount > Player(Index).StartFlash Then
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B)
        Player(Index).StartFlash = 0
    Else
    
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B, True)
    End If
    
    
    'Charm
    If GetPlayerEquipment(Index, charm) > 0 Then
    
    If Item(GetPlayerEquipment(Index, charm)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, charm)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, charm)).R, 255 - Item(GetPlayerEquipment(Index, charm)).G, 255 - Item(GetPlayerEquipment(Index, charm)).B)
       
    End If
    End If
    
    'Boots
    If GetPlayerEquipment(Index, boots) > 0 Then
    
    If Item(GetPlayerEquipment(Index, boots)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, boots)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, boots)).R, 255 - Item(GetPlayerEquipment(Index, boots)).G, 255 - Item(GetPlayerEquipment(Index, boots)).B)
       
    End If
    End If
    
    
    'Helmet
    If GetPlayerEquipment(Index, helmet) > 0 Then
    
    If Item(GetPlayerEquipment(Index, helmet)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, helmet)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, helmet)).R, 255 - Item(GetPlayerEquipment(Index, helmet)).G, 255 - Item(GetPlayerEquipment(Index, helmet)).B)
       
    End If
    End If
    
    'Armor
    If GetPlayerEquipment(Index, armor) > 0 Then
    
    If Item(GetPlayerEquipment(Index, armor)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, armor)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, armor)).R, 255 - Item(GetPlayerEquipment(Index, armor)).G, 255 - Item(GetPlayerEquipment(Index, armor)).B)
       
    End If
    End If
    
        'Whetstone
    If GetPlayerEquipment(Index, whetstone) > 0 Then
    
    If Item(GetPlayerEquipment(Index, whetstone)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, whetstone)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, whetstone)).R, 255 - Item(GetPlayerEquipment(Index, whetstone)).G, 255 - Item(GetPlayerEquipment(Index, whetstone)).B)
       
    End If
    End If
    
    
    'Ring
    If GetPlayerEquipment(Index, ring) > 0 Then
    
    If Item(GetPlayerEquipment(Index, ring)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, ring)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, ring)).R, 255 - Item(GetPlayerEquipment(Index, ring)).G, 255 - Item(GetPlayerEquipment(Index, ring)).B)
       
    End If
    End If
    
    'Enchant
    If GetPlayerEquipment(Index, enchant) > 0 Then
    
    If Item(GetPlayerEquipment(Index, enchant)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, enchant)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, enchant)).R, 255 - Item(GetPlayerEquipment(Index, enchant)).G, 255 - Item(GetPlayerEquipment(Index, enchant)).B)
       
    End If
    End If
    
    Case DIR_LEFT
    
    'Weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
    
    If Item(GetPlayerEquipment(Index, weapon)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, weapon)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, weapon)).R, 255 - Item(GetPlayerEquipment(Index, weapon)).G, 255 - Item(GetPlayerEquipment(Index, weapon)).B)
       
    End If
    End If
    
    If GetTickCount > Player(Index).StartFlash Then
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B)
        Player(Index).StartFlash = 0
    Else
    
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B, True)
    End If
    
    'Helmet
    If GetPlayerEquipment(Index, helmet) > 0 Then
    
    If Item(GetPlayerEquipment(Index, helmet)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, helmet)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, helmet)).R, 255 - Item(GetPlayerEquipment(Index, helmet)).G, 255 - Item(GetPlayerEquipment(Index, helmet)).B)
       
    End If
    End If
    
    'Charm
    If GetPlayerEquipment(Index, charm) > 0 Then
    
    If Item(GetPlayerEquipment(Index, charm)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, charm)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, charm)).R, 255 - Item(GetPlayerEquipment(Index, charm)).G, 255 - Item(GetPlayerEquipment(Index, charm)).B)
       
    End If
    End If
    
    'Boots
    If GetPlayerEquipment(Index, boots) > 0 Then
    
    If Item(GetPlayerEquipment(Index, boots)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, boots)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, boots)).R, 255 - Item(GetPlayerEquipment(Index, boots)).G, 255 - Item(GetPlayerEquipment(Index, boots)).B)
       
    End If
    End If
    
    
    'Armor
    If GetPlayerEquipment(Index, armor) > 0 Then
    
    If Item(GetPlayerEquipment(Index, armor)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, armor)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, armor)).R, 255 - Item(GetPlayerEquipment(Index, armor)).G, 255 - Item(GetPlayerEquipment(Index, armor)).B)
       
    End If
    End If
    
    'Whetstone
    If GetPlayerEquipment(Index, whetstone) > 0 Then
    
    If Item(GetPlayerEquipment(Index, whetstone)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, whetstone)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, whetstone)).R, 255 - Item(GetPlayerEquipment(Index, whetstone)).G, 255 - Item(GetPlayerEquipment(Index, whetstone)).B)
       
    End If
    End If
    
    
    'Ring
    If GetPlayerEquipment(Index, ring) > 0 Then
    
    If Item(GetPlayerEquipment(Index, ring)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, ring)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, ring)).R, 255 - Item(GetPlayerEquipment(Index, ring)).G, 255 - Item(GetPlayerEquipment(Index, ring)).B)
       
    End If
    End If
    
    'Enchant
    If GetPlayerEquipment(Index, enchant) > 0 Then
    
    If Item(GetPlayerEquipment(Index, enchant)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, enchant)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, enchant)).R, 255 - Item(GetPlayerEquipment(Index, enchant)).G, 255 - Item(GetPlayerEquipment(Index, enchant)).B)
       
    End If
    End If
    
    'Shield
    If GetPlayerEquipment(Index, shield) > 0 Then
    
    If Item(GetPlayerEquipment(Index, shield)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, shield)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, shield)).R, 255 - Item(GetPlayerEquipment(Index, shield)).G, 255 - Item(GetPlayerEquipment(Index, shield)).B)
       
    End If
    End If
    
    Case DIR_RIGHT
    
    'Shield
    If GetPlayerEquipment(Index, shield) > 0 Then
    
    If Item(GetPlayerEquipment(Index, shield)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, shield)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, shield)).R, 255 - Item(GetPlayerEquipment(Index, shield)).G, 255 - Item(GetPlayerEquipment(Index, shield)).B)
       
    End If
    End If
    
    
    If GetTickCount > Player(Index).StartFlash Then
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B)
        Player(Index).StartFlash = 0
    Else
    
        Call DrawSprite(Sprite, X, Y, Rec, a, R, G, B, True)
    End If
    
    'Helmet
    If GetPlayerEquipment(Index, helmet) > 0 Then
    
    If Item(GetPlayerEquipment(Index, helmet)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, helmet)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, helmet)).R, 255 - Item(GetPlayerEquipment(Index, helmet)).G, 255 - Item(GetPlayerEquipment(Index, helmet)).B)
       
    End If
    End If
    
    'Charm
    If GetPlayerEquipment(Index, charm) > 0 Then
    
    If Item(GetPlayerEquipment(Index, charm)).Paperdoll > 0 Then
   Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, charm)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, charm)).R, 255 - Item(GetPlayerEquipment(Index, charm)).G, 255 - Item(GetPlayerEquipment(Index, charm)).B)
       
    End If
    End If
    
    'Boots
    If GetPlayerEquipment(Index, boots) > 0 Then
    
    If Item(GetPlayerEquipment(Index, boots)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, boots)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, boots)).R, 255 - Item(GetPlayerEquipment(Index, boots)).G, 255 - Item(GetPlayerEquipment(Index, boots)).B)
       
    End If
    End If
    
    'Armor
    If GetPlayerEquipment(Index, armor) > 0 Then
    
    If Item(GetPlayerEquipment(Index, armor)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, armor)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, armor)).R, 255 - Item(GetPlayerEquipment(Index, armor)).G, 255 - Item(GetPlayerEquipment(Index, armor)).B)
       
    End If
    End If
    
    'Whetstone
    If GetPlayerEquipment(Index, whetstone) > 0 Then
    
    If Item(GetPlayerEquipment(Index, whetstone)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, whetstone)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, whetstone)).R, 255 - Item(GetPlayerEquipment(Index, whetstone)).G, 255 - Item(GetPlayerEquipment(Index, whetstone)).B)
       
    End If
    End If
    
    'Ring
    If GetPlayerEquipment(Index, ring) > 0 Then
    
    If Item(GetPlayerEquipment(Index, ring)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, ring)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, ring)).R, 255 - Item(GetPlayerEquipment(Index, ring)).G, 255 - Item(GetPlayerEquipment(Index, ring)).B)
       
    End If
    End If
    
    'Enchant
    If GetPlayerEquipment(Index, enchant) > 0 Then
    
    If Item(GetPlayerEquipment(Index, enchant)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, enchant)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, enchant)).R, 255 - Item(GetPlayerEquipment(Index, enchant)).G, 255 - Item(GetPlayerEquipment(Index, enchant)).B)
       
    End If
    End If
    
    'Weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
    
    If Item(GetPlayerEquipment(Index, weapon)).Paperdoll > 0 Then
    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, weapon)).Paperdoll, Anim, spritetop, a, 255 - Item(GetPlayerEquipment(Index, weapon)).R, 255 - Item(GetPlayerEquipment(Index, weapon)).G, 255 - Item(GetPlayerEquipment(Index, weapon)).B)
       
    End If
    End If
    
    End Select
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayer", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub DrawNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte, I As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim Rec As RECT
Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    
    Sprite = NPC(MapNpc(MapNpcNum).num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    'AttackSpeed = 1000
    
    If NPC(MapNpc(MapNpcNum).num).AttackSpeed > 0 Then
        AttackSpeed = NPC(MapNpc(MapNpcNum).num).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    If VXFRAME = False Then
        ' Reset frame
        If MapNpc(MapNpcNum).Step = 3 Then
            Anim = 0
        ElseIf MapNpc(MapNpcNum).Step = 1 Then
            Anim = 2
        End If
    Else
        Anim = 1
    End If

    ' Reset frame
    'anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            If VXFRAME = False Then
                Anim = 3
            Else
                Anim = 2
            End If
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    ' animated npcs
    If isAnimated(NPC(MapNpc(MapNpcNum).num).Sprite) Then
    With MapNpc(MapNpcNum)
    If .AnimTimer + 100 <= GetTickCount Then
    .Anim = .Anim + 1
    If .Anim >= 3 Then .Anim = 0
    .AnimTimer = GetTickCount
    End If
    Anim = .Anim
    End With
    End If

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With Rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        If VXFRAME = False Then
            .Left = Anim * (Tex_Character(Sprite).Width / 4)
            .Right = .Left + (Tex_Character(Sprite).Width / 4)
        Else
            .Left = Anim * (Tex_Character(Sprite).Width / 3)
            .Right = .Left + (Tex_Character(Sprite).Width / 3)
        End If
    End With

    ' Calculate the X
    If VXFRAME = False Then
        X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    Else
        X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset
    End If
    
    ' render player shadow
   ' If hasShadow(Sprite) Then RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 150)
    'RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y + 4), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    ' render the actual sprite
    If GetTickCount > MapNpc(MapNpcNum).StartFlash Then
        Call DrawSprite(Sprite, X, Y, Rec, 255 - NPC(MapNpc(MapNpcNum).num).a, 255 - NPC(MapNpc(MapNpcNum).num).R, 255 - NPC(MapNpc(MapNpcNum).num).G, 255 - NPC(MapNpc(MapNpcNum).num).B)
        MapNpc(MapNpcNum).StartFlash = 0
    Else
        Call DrawSprite(Sprite, X, Y, Rec, 255 - NPC(MapNpc(MapNpcNum).num).a, 255 - NPC(MapNpc(MapNpcNum).num).R, 255 - NPC(MapNpc(MapNpcNum).num).G, 255 - NPC(MapNpc(MapNpcNum).num).B, True)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long, Optional a As Byte = 255, Optional R As Byte = 255, Optional G As Byte = 255, Optional B As Byte = 255)
Dim Rec As RECT
Dim X As Long, Y As Long, I As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    With Rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        If VXFRAME = False Then
            .Left = Anim * (Tex_Paperdoll(Sprite).Width / 4)
            .Right = .Left + (Tex_Paperdoll(Sprite).Width / 4)
        Else
            .Left = Anim * (Tex_Paperdoll(Sprite).Width / 3)
            .Right = .Left + (Tex_Paperdoll(Sprite).Width / 3)
        End If
    End With
    
    ' clipping
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With Rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With Rec
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Paperdoll(Sprite), X, Y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorARGB(a, R, G, B)
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, Rec As RECT, Optional a As Byte = 255, Optional R As Byte = 255, Optional G As Byte = 255, Optional B As Byte = 255, Optional Flash As Boolean = False)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
    
    If Flash = True Then
        RenderTexture Tex_Character(Sprite), X, Y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(R, G, B, Round(a / 2))
    Else
        RenderTexture Tex_Character(Sprite), X, Y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(R, G, B, a)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, color As Long, X As Long, Y As Long, RenderState As Long

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    color = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)

    RenderState = 0
    ' render state
    Select Case RenderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For X = 0 To ((Map.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, color
        Next
    Next
    
    ' reset render state
    If RenderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
Dim color As Long

    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    
    If BlindDuration > 0 Then
    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, 255)
    End If
    
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, color
End Sub

Public Sub DrawWeather()
Dim color As Long, I As Long, SpriteLeft As Long
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(I).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(I).X), ConvertMapY(WeatherParticle(I).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Sub DrawAnimatedInvItems()
Dim I As Long
Dim itemNum As Long, ItemPic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim Rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For I = 1 To MAX_MAP_ITEMS

        If MapItem(I).num > 0 Then
            ItemPic = Item(MapItem(I).num).Pic

            If ItemPic < 1 Or ItemPic > numitems Then Exit Sub
            MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(I).Frame < MaxFrames - 1 Then
                MapItem(I).Frame = MapItem(I).Frame + 1
            Else
                MapItem(I).Frame = 1
            End If
        End If

    Next

    For I = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, I)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic

            If ItemPic > 0 And ItemPic <= numitems Then
                If Tex_Item(ItemPic).Width > 64 Then
                    MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(I) < MaxFrames - 1 Then
                        InvItemFrame(I) = InvItemFrame(I) + 1
                    Else
                        InvItemFrame(I) = 1
                    End If

                    With Rec
                        .Top = 0
                        .Bottom = 32
                        .Left = (Tex_Item(ItemPic).Width / 2) + (InvItemFrame(I) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' We'll now re-Draw the item, and place the currency value over it again :P
                    RenderTextureByRects Tex_Item(ItemPic), Rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, I))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, Yellow, 0
                    End If
                End If
            End If
        End If

    Next

    'frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimatedInvItems", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long
Dim sRECT As RECT
Dim drect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmEditor_Map.scrlPictureX.Value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.Value * PIC_Y
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRECT.Left = frmEditor_Map.scrlPictureX.Value * PIC_X
    sRECT.Top = frmEditor_Map.scrlPictureY.Value * PIC_Y
    sRECT.Right = sRECT.Left + Width
    sRECT.Bottom = sRECT.Top + Height
    
    drect.Top = 0
    drect.Bottom = Height
    drect.Left = 0
    drect.Right = Width
    
    RenderTextureByRects Tex_Tileset(Tileset), sRECT, drect
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    
    With destRect
        .X1 = (EditorTileX * 32) - sRECT.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRECT.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    
    DrawSelectionBox destRect
        
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picBack.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    
    'Now render the selection tiles and we are done!
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picBack.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawSelectionBox(drect As D3DRECT)
Dim Width As Long, Height As Long, X As Long, Y As Long
    Width = drect.X2 - drect.X1
    Height = drect.Y2 - drect.Y1
    X = drect.X1
    Y = drect.Y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, X, Y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Selection, X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Selection, X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Selection, X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Selection, X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawTileOutline()
Dim Rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.Value Then Exit Sub

    With Rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTileOutline", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterDrawSprite()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.Value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    SetTexture Tex_Character(Sprite)
    
    If VXFRAME = False Then
        Width = Tex_Character(Sprite).Width / 4
    Else
        Width = Tex_Character(Sprite).Width / 3
    End If
    
    Height = Tex_Character(Sprite).Height / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    drect.Top = 0
    drect.Bottom = Height
    drect.Left = 0
    drect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(Sprite), sRECT, drect
    
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMenu.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterDrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawMapItem()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemNum), sRECT, drect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawKey()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = Item(frmEditor_Map.scrlMapKey.Value).Pic

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    
    RenderTextureByRects Tex_Item(itemNum), sRECT, drect
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapKey.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawItem()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_Item.scrlPic.Value

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    drect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemNum), sRECT, drect, D3DColorARGB(255 - frmEditor_Item.scrlA, 255 - frmEditor_Item.scrlR, 255 - frmEditor_Item.scrlG, 255 - frmEditor_Item.scrlB)
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRECT.Left = 0
    If VXFRAME = False Then
        sRECT.Right = Tex_Paperdoll(Sprite).Width / 4
    Else
        sRECT.Right = Tex_Paperdoll(Sprite).Width / 3
    End If
    ' same for destination as source
    drect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRECT, drect
                    
    With destRect
        .X1 = 0
        If VXFRAME = False Then
            .X2 = Tex_Paperdoll(Sprite).Width / 4
        Else
            .X2 = Tex_Paperdoll(Sprite).Width / 3
        End If
        .Y1 = 0
        .Y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picPaperdoll.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
Dim iconnum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.Value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(iconnum), sRECT, drect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_DrawAnim()
Dim I As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim sx As Long, sy As Long, sRECT As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    sRECT.Top = 0
    sRECT.Bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For I = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(I).Value
        
        If Animationnum <= 0 Or Animationnum > NumAnimations Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(I)
            FrameCount = frmEditor_Animation.scrlFrameCount(I)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(I) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(I) >= FrameCount Then
                    AnimEditorFrame(I) = 1
                Else
                    AnimEditorFrame(I) = AnimEditorFrame(I) + 1
                End If
                AnimEditorTimer(I) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(I).Value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    Height = 192

                    sy = (Height * ((AnimEditorFrame(I) - 1) \ AnimColumns))
                    sx = (Width * (((AnimEditorFrame(I) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call Direct3D_Device.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call Direct3D_Device.BeginScene
                    
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Animation(Animationnum), 0, 0, sx, sy, Width, Height, Width, Height
                    
                    ' Finish Rendering
                    Call Direct3D_Device.EndScene
                    Call Direct3D_Device.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(I).hWnd, ByVal 0)
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
Dim Sprite As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 2 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    drect.Top = 0
    drect.Bottom = SIZE_Y
    drect.Left = 0
    drect.Right = SIZE_X
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRECT, drect, D3DColorARGB(255 - frmEditor_NPC.scrlA.Value, 255 - frmEditor_NPC.scrlR.Value, 255 - frmEditor_NPC.scrlG.Value, 255 - frmEditor_NPC.scrlB.Value)
    
    With destRect
        .X1 = 0
        .X2 = SIZE_X
        .Y1 = 0
        .Y2 = SIZE_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawSprite()
Dim Sprite As Long
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        drect.Top = 0
        drect.Bottom = Tex_Resource(Sprite).Height
        drect.Left = 0
        drect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, drect
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        drect.Top = 0
        drect.Bottom = Tex_Resource(Sprite).Height
        drect.Left = 0
        drect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, drect
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim I As Long
Dim Rec As RECT
Dim rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
   On Error GoTo errorhandler
    
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera

    ' unload any textures we need to unload
    UnloadTextures
   Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
        
        Direct3D_Device.BeginScene
        
        If Map.Panorama > 0 Then
                RenderTexture Tex_Panorama(Map.Panorama), ParallaxX, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, frmMain.ScaleWidth, frmMain.ScaleHeight
                RenderTexture Tex_Panorama(Map.Panorama), ParallaxX + frmMain.ScaleWidth, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, frmMain.ScaleWidth, frmMain.ScaleHeight
            End If
            
            ' blit lower tiles
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                            Call DrawMapTile(X, Y)
                        End If
                    Next
                Next
            End If
        
            ' render the decals
            For I = 1 To MAX_BYTE
                Call DrawBlood(I)
            Next
        
            ' Blit out the items
            If numitems > 0 Then
                For I = 1 To MAX_MAP_ITEMS
                    If MapItem(I).num > 0 Then
                        Call DrawItem(I)
                    End If
                Next
            End If
            
            If Map.CurrentEvents > 0 Then
                For I = 1 To Map.CurrentEvents
                    If Map.MapEvents(I).Position = 0 Then
                        DrawEvent I
                    End If
                Next
            End If
            
            ' draw animations
            If NumAnimations > 0 Then
                For I = 1 To MAX_BYTE
                    If AnimInstance(I).Used(0) Then
                        DrawAnimation I, 0
                    End If
                Next
            End If
        
            ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
            For Y = 0 To Map.MaxY
                If NumCharacters > 0 Then
                
                    If Map.CurrentEvents > 0 Then
                        For I = 1 To Map.CurrentEvents
                            If Map.MapEvents(I).Position = 1 Then
                                If Y = Map.MapEvents(I).Y Then
                                    DrawEvent I
                                End If
                            End If
                        Next
                    End If
                    
                    ' Players
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                            If Player(I).Y = Y And (Not GetPlayerVisible(I) = 1 Or I = MyIndex) Then
                                Call DrawPlayer(I, GetPlayerColorA(I), GetPlayerColorR(I), GetPlayerColorG(I), GetPlayerColorB(I))
                            End If
                        End If
                    Next
                    
                    
                
                    ' Npcs
                    For I = 1 To Npc_HighIndex
                        If MapNpc(I).Y = Y Then
                            Call DrawNpc(I)
                        End If
                    Next
                End If
                
                ' Resources
                If NumResources > 0 Then
                    If Resources_Init Then
                        If Resource_Index > 0 Then
                            For I = 1 To Resource_Index
                                If MapResource(I).Y = Y Then
                                    Call DrawMapResource(I)
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            
            If NumProjectiles > 0 Then
                Call DrawProjectile
            End If
            
            ' animations
            If NumAnimations > 0 Then
                For I = 1 To MAX_BYTE
                    If AnimInstance(I).Used(1) Then
                        DrawAnimation I, 1
                    End If
                Next
            End If
        
            ' blit out upper tiles
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                            Call DrawMapFringeTile(X, Y)
                        End If
                    Next
                Next
            End If
            
             ' blit out lights
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                        If Map.Tile(X, Y).Type = TILE_TYPE_LIGHT Then
                        If DayTime = False Then
                            If Not Map.DayNight = 2 Then Call DrawLight(X * 32, Y * 32, Map.Tile(X, Y).data1, Map.Tile(X, Y).Data2, Map.Tile(X, Y).Data3, Map.Tile(X, Y).Data4)
                        End If
                    End If
                        End If
                    Next
                Next
            End If
            
            If Map.CurrentEvents > 0 Then
                For I = 1 To Map.CurrentEvents
                    If Map.MapEvents(I).Position = 2 Then
                        DrawEvent I
                    End If
                Next
            End If
            
            DrawWeather
            DrawFog
            
            
            DrawTint
            
            ' blit out a square at mouse cursor
            If InMapEditor Then
                If frmEditor_Map.optBlock.Value = True Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                Call DrawDirection(X, Y)
                            End If
                        Next
                    Next
                End If
                Call DrawTileOutline
            End If
            
            ' Render the bars
            DrawBars
            
            ' Draw the target icon
            If myTarget > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    DrawTarget (Player(myTarget).X * 32) + Player(myTarget).xOffset, (Player(myTarget).Y * 32) + Player(myTarget).yOffset
                ElseIf myTargetType = TARGET_TYPE_NPC Then
                    DrawTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).yOffset
                End If
            End If
            
            ' Draw the hover icon
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    If Player(I).Map = Player(MyIndex).Map Then
                        If CurX = Player(I).X And CurY = Player(I).Y Then
                            If myTargetType = TARGET_TYPE_PLAYER And myTarget = I Or GetPlayerVisible(I) = 1 Then
                                ' dont render lol
                            Else
                                DrawHover TARGET_TYPE_PLAYER, I, (Player(I).X * 32) + Player(I).xOffset, (Player(I).Y * 32) + Player(I).yOffset
                            End If
                        End If
                    End If
                End If
            Next
            For I = 1 To Npc_HighIndex
                If MapNpc(I).num > 0 Then
                    If CurX = MapNpc(I).X And CurY = MapNpc(I).Y Then
                        If myTargetType = TARGET_TYPE_NPC And myTarget = I Then
                            ' dont render lol
                        Else
                            DrawHover TARGET_TYPE_NPC, I, (MapNpc(I).X * 32) + MapNpc(I).xOffset, (MapNpc(I).Y * 32) + MapNpc(I).yOffset
                        End If
                    End If
                End If
            Next
            
            If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
            
            ' Get rec
            With Rec
                .Top = Camera.Top
                .Bottom = .Top + ScreenY
                .Left = Camera.Left
                .Right = .Left + ScreenX
            End With
                
            ' rec_pos
            With rec_pos
                .Bottom = ScreenY
                .Right = ScreenX
            End With
                
            With srcRect
                .X1 = 0
                .X2 = frmMain.ScaleWidth
                .Y1 = 0
                .Y2 = frmMain.ScaleHeight
            End With
            
            If BFPS Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS), 12, 100, Yellow, 0
            End If
            
            ' draw cursor, player X and Y locations
            If BLoc Then
                RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 12, 114, Yellow, 0
                RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 12, 128, Yellow, 0
                RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 12, 142, Yellow, 0
            End If
            
            ' draw player names
            For I = 1 To Player_HighIndex
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) And (Not GetPlayerVisible(I) = 1 Or I = MyIndex) Then
                    Call DrawPlayerName(I)
                End If
            Next
            
            For I = 1 To Map.CurrentEvents
                If Map.MapEvents(I).Visible = 1 Then
                    If Map.MapEvents(I).ShowName = 1 Then
                        DrawEventName (I)
                    End If
                End If
            Next
            
            ' draw npc names
            For I = 1 To Npc_HighIndex
                If MapNpc(I).num > 0 Then
                    Call DrawNpcName(I)
                End If
            Next
            
                ' draw the messages
            For I = 1 To MAX_BYTE
                If chatBubble(I).active Then
                    DrawChatBubble I
                End If
            Next
            
            For I = 1 To Action_HighIndex
                Call DrawActionMsg(I)
            Next I
            
            If OverlayVisible Then Call DrawOverlay(150, 0, 0, 0)
            If Not NightDisabled Then DrawNight
            
            'If BMENU Then
              '  DrawGuildMenu
            'End If
            
            DrawBossMsg
            
            If Not hideGUI Then DrawGUI
            
            If Not hideGUI Then
            RenderText Font_Default, KeepTwoDigit(GameHours) & ":" & KeepTwoDigit(GameMinutes) & ":" & KeepTwoDigit(GameSeconds), DrawMapNameX + 400, DrawMapNameY, Yellow
            RenderText Font_Default, Map.name, DrawMapNameX, DrawMapNameY, DrawMapNameColor
            End If
            If InMapEditor And frmEditor_Map.optEvent.Value = True Then DrawEvents
            If InMapEditor Then Call DrawMapAttributes
            
            
            If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
            If FlashTimer > GetTickCount Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, -1
            
            RenderTexture Tex_GUI(36), GlobalX, GlobalY, 0, 0, 32, 32, 32, 32
        Direct3D_Device.EndScene
        
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If

    ' Error handler
    Exit Sub
    
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If Options.Debug = 1 Then
            HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Unrecoverable DX8 error."
        DestroyGame
    End If
End Sub

Sub HandleDeviceLost()
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures True
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
    
   LoadTextures
   
End Sub

Private Function DirectX_ReInit() As Boolean

    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 800 ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 600 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
Dim OffsetX As Long
Dim OffsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    OffsetX = Player(MyIndex).xOffset + PIC_X
    OffsetY = Player(MyIndex).yOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        OffsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                OffsetX = Player(MyIndex).xOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        OffsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                OffsetY = Player(MyIndex).yOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        OffsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                OffsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        OffsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                OffsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = OffsetY
        .Bottom = .Top + ScreenY
        .Left = OffsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim I As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).layer(I).Tileset > 0 And Map.Tile(X, Y).layer(I).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).layer(I).Tileset) = True
                End If
            Next
        Next
    Next
    
    For I = 1 To NumTileSets
        If tilesetInUse(I) Then
        
        Else
            ' unload tileset
            'Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            'Set Tex_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub DrawEvents()
Dim sRECT As RECT
Dim Width As Long, Height As Long, I As Long, X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For I = 1 To Map.EventCount
        If Map.Events(I).pageCount <= 0 Then
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(X), ConvertMapY(Y), sRECT.Left, sRECT.Right, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        X = Map.Events(I).X * 32
        Y = Map.Events(I).Y * 32
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
    
        
        If I > Map.EventCount Then Exit Sub
        If 1 > Map.Events(I).pageCount Then Exit Sub
        Select Case Map.Events(I).Pages(1).GraphicType
            Case 0
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Case 1
                If Map.Events(I).Pages(1).Graphic > 0 And Map.Events(I).Pages(1).Graphic <= NumCharacters Then
                    
                    sRECT.Top = (Map.Events(I).Pages(1).GraphicY * (Tex_Character(Map.Events(I).Pages(1).Graphic).Height / 4))
                    
                    If VXFRAME = False Then
                        sRECT.Left = (Map.Events(I).Pages(1).GraphicX * (Tex_Character(Map.Events(I).Pages(1).Graphic).Width / 4))
                    Else
                        sRECT.Left = (Map.Events(I).Pages(1).GraphicX * (Tex_Character(Map.Events(I).Pages(1).Graphic).Width / 3))
                    End If
                    
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Character(Map.Events(I).Pages(1).Graphic), X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            Case 2
                If Map.Events(I).Pages(1).Graphic > 0 And Map.Events(I).Pages(1).Graphic < NumTileSets Then
                    sRECT.Top = Map.Events(I).Pages(1).GraphicY * 32
                    sRECT.Left = Map.Events(I).Pages(1).GraphicX * 32
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Tileset(Map.Events(I).Pages(1).Graphic), X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        End Select
nextevent:
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEvents", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorEvent_DrawGraphic()
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - sRECT.Left)
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width
                    End If
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.Top + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - sRECT.Top)
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height
                    End If
                    
                    With drect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    With destRect
                        .X1 = drect.Left
                        .X2 = drect.Right
                        .Y1 = drect.Top
                        .Y2 = drect.Bottom
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRECT, drect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            If VXFRAME = False Then
                                .X1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4)) - sRECT.Left
                                .X2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) + .X1
                            Else
                                .X1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 3)) - sRECT.Left
                                .X2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 3) + .X1
                            End If
                            .Y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4)) - sRECT.Top
                            .Y2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) + .Y1
                        End With

                    Else
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .X1 = drect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = drect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + 800
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.vScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.Top + 512
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                    
                    If sRECT.Left = -1 Then sRECT.Left = 0
                    If sRECT.Top = -1 Then sRECT.Top = 0
                    
                    With drect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, drect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = PIC_X + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = PIC_Y + .Y1
                        End With

                    Else
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .X1 = drect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = drect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    If VXFRAME = False Then
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                        sRECT.Right = sRECT.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    Else
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 3)
                        sRECT.Right = sRECT.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 3)
                    End If
                    sRECT.Bottom = sRECT.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    With drect
                        drect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                        drect.Bottom = drect.Top + (sRECT.Bottom - sRECT.Top)
                        drect.Left = (121 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                        drect.Right = drect.Left + (sRECT.Right - sRECT.Left)
                    End With
                    With destRect
                        .X1 = drect.Left
                        .X2 = drect.Right
                        .Y1 = drect.Top
                        .Y2 = drect.Bottom
                    End With
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRECT, drect
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + 32
                        sRECT.Right = sRECT.Left + 32
                        With drect
                            drect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            drect.Bottom = drect.Top + (sRECT.Bottom - sRECT.Top)
                            drect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            drect.Right = drect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRect
                            .X1 = drect.Left
                            .X2 = drect.Right
                            .Y1 = drect.Top
                            .Y2 = drect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, drect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    Else
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRECT.Right = sRECT.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With drect
                            drect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            drect.Bottom = drect.Top + (sRECT.Bottom - sRECT.Top)
                            drect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            drect.Right = drect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRect
                            .X1 = drect.Left
                            .X2 = drect.Right
                            .Y1 = drect.Top
                            .Y2 = drect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, drect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    End If
                End If
        End Select
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEvent(ID As Long)
    Dim X As Long, Y As Long, Width As Long, Height As Long, sRECT As RECT, drect As RECT, Anim As Long, spritetop As Long
    If Map.MapEvents(ID).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(ID).GraphicType
        Case 0
            Exit Sub
            
        Case 1
            If Map.MapEvents(ID).GraphicNum <= 0 Or Map.MapEvents(ID).GraphicNum > NumCharacters Then Exit Sub
            If VXFRAME = False Then
                Width = Tex_Character(Map.MapEvents(ID).GraphicNum).Width / 4
            Else
                Width = Tex_Character(Map.MapEvents(ID).GraphicNum).Width / 3
            End If
            Height = Tex_Character(Map.MapEvents(ID).GraphicNum).Height / 4
            ' Reset frame
            If VXFRAME = False Then
                If Map.MapEvents(ID).Step = 3 Then
                    Anim = 0
                ElseIf Map.MapEvents(ID).Step = 1 Then
                    Anim = 2
                End If
            Else
                
            End If
            
            Select Case Map.MapEvents(ID).Dir
                Case DIR_UP
                    If (Map.MapEvents(ID).yOffset > 8) Then Anim = Map.MapEvents(ID).Step
                Case DIR_DOWN
                    If (Map.MapEvents(ID).yOffset < -8) Then Anim = Map.MapEvents(ID).Step
                Case DIR_LEFT
                    If (Map.MapEvents(ID).xOffset > 8) Then Anim = Map.MapEvents(ID).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(ID).xOffset < -8) Then Anim = Map.MapEvents(ID).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(ID).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
            
            If Map.MapEvents(ID).WalkAnim = 1 Then Anim = 0
            
            If Map.MapEvents(ID).Moving = 0 Then Anim = Map.MapEvents(ID).GraphicX
            
            With sRECT
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = Anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            X = Map.MapEvents(ID).X * PIC_X + Map.MapEvents(ID).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = Map.MapEvents(ID).Y * PIC_Y + Map.MapEvents(ID).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                Y = Map.MapEvents(ID).Y * PIC_Y + Map.MapEvents(ID).yOffset
            End If
        
            ' render the actual sprite
            Call DrawSprite(Map.MapEvents(ID).GraphicNum, X, Y, sRECT)
            
        Case 2
            If Map.MapEvents(ID).GraphicNum < 1 Or Map.MapEvents(ID).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(ID).GraphicY2 > 0 Or Map.MapEvents(ID).GraphicX2 > 0 Then
                With sRECT
                    .Top = Map.MapEvents(ID).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(ID).GraphicY2 - Map.MapEvents(ID).GraphicY) * 32)
                    .Left = Map.MapEvents(ID).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(ID).GraphicX2 - Map.MapEvents(ID).GraphicX) * 32)
                End With
            Else
                With sRECT
                    .Top = Map.MapEvents(ID).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(ID).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            X = Map.MapEvents(ID).X * 32
            Y = Map.MapEvents(ID).Y * 32
            
            X = X - ((sRECT.Right - sRECT.Left) / 2)
            Y = Y - (sRECT.Bottom - sRECT.Top) + 32
            
            
            If Map.MapEvents(ID).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(ID).GraphicNum), ConvertMapX(Map.MapEvents(ID).X * 32), ConvertMapY((Map.MapEvents(ID).Y - ((Map.MapEvents(ID).GraphicY2 - Map.MapEvents(ID).GraphicY) - 1)) * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(ID).GraphicNum), ConvertMapX(Map.MapEvents(ID).X * 32), ConvertMapY(Map.MapEvents(ID).Y * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
    End Select
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, Z As Single, RHW As Single, color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.color = color
    'Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
' round it
Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
' if it rounded down, force it up
If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
End Sub

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.Visible Then
        If frmMenu.picCharacter.Visible Then NewCharacterDrawSprite
    End If
    
    If frmEditor_Animation.Visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
        EditorItem_DrawProjectile
    End If
    
    If frmEditor_Map.Visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
        If frmEditor_Map.fraMapKey.Visible Then EditorMap_DrawKey
        If frmEditor_Map.fraLight.Visible Then EditorMap_DrawLight
    End If
    
    If frmEditor_NPC.Visible Then
        EditorNpc_DrawSprite
    End If
    
    If frmEditor_Resource.Visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
    End If
    
    If frmEditor_Events.Visible Then
        EditorEvent_DrawGraphic
    End If
End Sub
Public Sub DrawGUI()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' render shadow
    'EngineRenderRectangle Tex_GUI(27), 0, 0, 0, 0, 800, 64, 1, 64, 800, 64
    'EngineRenderRectangle Tex_GUI(26), 0, 600 - 64, 0, 0, 800, 64, 1, 64, 800, 64
    RenderTexture Tex_GUI(23), 0, 0, 0, 0, 800, 64, 1, 64
    RenderTexture Tex_GUI(22), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    ' render chatbox
        If Not inChat Then
            If chatOn Then
                Width = 412
                Height = 145
                RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).X, GUIWindow(GUI_CHAT).Y, 0, 0, Width, Height, Width, Height
                RenderText Font_Default, RenderChatText & chatShowLine, GUIWindow(GUI_CHAT).X + 38, GUIWindow(GUI_CHAT).Y + 126, White
                ' draw buttons
                For I = 34 To 35
                    ' set co-ordinate
                    X = GUIWindow(GUI_CHAT).X + Buttons(I).X
                    Y = GUIWindow(GUI_CHAT).Y + Buttons(I).Y
                    Width = Buttons(I).Width
                    Height = Buttons(I).Height
                    ' check for state
                    If Buttons(I).state = 2 Then
                        ' we're clicked boyo
                        'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                    ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
                        ' we're hoverin'
                        'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                        ' play sound if needed
                        If Not lastButtonSound = I Then
                            PlaySound Sound_ButtonHover, -1, -1
                            lastButtonSound = I
                        End If
                    Else
                        ' we're normal
                        'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                        ' reset sound if needed
                        If lastButtonSound = I Then lastButtonSound = 0
                    End If
                Next
            Else
                RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).X, GUIWindow(GUI_CHAT).Y + 123, 0, 123, 412, 22, 412, 22
            End If
            RenderChatTextBuffer
        Else
            If GUIWindow(GUI_CURRENCY).Visible Then DrawCurrency
            If GUIWindow(GUI_EVENTCHAT).Visible Then DrawEventChat
            If GUIWindow(GUI_QUESTDIALOGUE).Visible Then DrawQuestDialogue
        End If
    
    DrawGUIBars
    
    
    ' render menu
    If GUIWindow(GUI_MENU).Visible Then DrawMenu
    
    ' render hotbar
    If GUIWindow(GUI_HOTBAR).Visible Then DrawHotbar
    
    ' render menus
    If GUIWindow(GUI_INVENTORY).Visible Then DrawInventory
    If GUIWindow(GUI_SPELLS).Visible Then DrawSkills
    If GUIWindow(GUI_CHARACTER).Visible Then DrawCharacter
    If GUIWindow(GUI_OPTIONS).Visible Then DrawOptions
    If GUIWindow(GUI_PARTY).Visible Then DrawParty
    If GUIWindow(GUI_SHOP).Visible Then DrawShop
    If GUIWindow(GUI_BANK).Visible Then DrawBank
    If GUIWindow(GUI_TRADE).Visible Then DrawTrade
    If GUIWindow(GUI_DIALOGUE).Visible Then DrawDialogue
    If GUIWindow(GUI_GUILD).Visible Then DrawGuildMenu
    If GUIWindow(GUI_QUESTLOG).Visible Then DrawQuestLog
    If GUIWindow(GUI_NEWCLASS).Visible Then DrawNewClass
    If GUIWindow(GUI_NEWS).Visible = True Then DrawNews
    
    
    ' Drag and drop
    DrawDragItem
    DrawDragSpell
    
    ' Descriptions
    If Not mouseClicked Then
        DrawInventoryItemDesc
        DrawCharacterItemDesc
        DrawPlayerSpellDesc
        DrawBankItemDesc
        DrawTradeItemDesc
    End If
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' check if the tile can be rendered
        If .layer(layerNum).Tileset <= 0 Or .layer(layerNum).Tileset > NumTileSets Then
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_NORMAL
        Else
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).layer(layerNum).X * 32) + Autotile(X, Y).layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).layer(layerNum).Y * 32) + Autotile(X, Y).layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).layer(layerNum).Tileset <> Map.Tile(X2, Y2).layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).layer(layerNum).X <> Map.Tile(X2, Y2).layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).layer(layerNum).Y <> Map.Tile(X2, Y2).layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    RenderTexture Tex_Tileset(Map.Tile(X, Y).layer(layerNum).Tileset), destX, destY, Autotile(X, Y).layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, -1
End Sub

Public Sub DrawItem(ByVal itemNum As Long)
Dim PicNum As Integer, dontRender As Boolean, I As Long, tmpIndex As Long, colour As Long
    
    PicNum = Item(MapItem(itemNum).num).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

     ' if it's not us then don't render
    If MapItem(itemNum).PlayerName <> vbNullString Then
        If Trim$(MapItem(itemNum).PlayerName) <> Trim$(GetPlayerName(MyIndex)) Then
            dontRender = True
        End If
        ' make sure it's not a party drop
        If Party.Leader > 0 Then
            For I = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(I)
                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemNum).PlayerName) Then
                        dontRender = False
                    End If
                End If
            Next
        End If
    End If
    
    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemNum).X * PIC_X), ConvertMapY(MapItem(itemNum).Y * PIC_Y), 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(MapItem(itemNum).num).a, 255 - Item(MapItem(itemNum).num).R, 255 - Item(MapItem(itemNum).num).G, 255 - Item(MapItem(itemNum).num).B)
        
        If AltDown Then
            ' work out name colour
            Select Case Item(MapItem(itemNum).num).Rarity
                Case 0
                    colour = White
                Case 1
                    colour = Grey
                Case 2
                    colour = White
                Case 3
                    colour = Green
                Case 4
                    colour = Cyan
                Case 5
                    colour = Magenta
            End Select
            RenderText Font_Default, Trim$(Item(MapItem(itemNum).num).name), ConvertMapX((MapItem(itemNum).X * PIC_X) + 16) - (EngineGetTextWidth(Font_Default, Trim$(Item(MapItem(itemNum).num).name)) / 2), ConvertMapY((MapItem(itemNum).Y * PIC_Y) - 14), colour
        End If
    End If
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemNum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not itemNum > 0 Then Exit Sub
    
    PicNum = Item(itemNum).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    'EngineRenderRectangle Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, spellnum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    spellnum = PlayerSpells(DragSpell).Spell
    If Not spellnum > 0 Then Exit Sub
    
    PicNum = Spell(spellnum).Icon

    If PicNum < 1 Or PicNum > NumSpellIcons Then Exit Sub

    'EngineRenderRectangle Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_SpellIcon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawHotbar()
Dim I As Long, X As Long, Y As Long, t As Long, sS As String
Dim Width As Long, Height As Long, color As Long

    For I = 1 To MAX_HOTBAR
        ' draw the box
        X = GUIWindow(GUI_HOTBAR).X + ((I - 1) * (5 + 36))
        Y = GUIWindow(GUI_HOTBAR).Y
        Width = 36
        Height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        RenderTexture Tex_GUI(2), X, Y, 0, 0, Width, Height, Width, Height
        ' draw the icon
        Select Case Hotbar(I).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(I).Slot).name) > 0 Then
                    If Item(Hotbar(I).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_Item(Item(Hotbar(I).Slot).Pic), X + 2, Y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(Hotbar(I).Slot).a, 255 - Item(Hotbar(I).Slot).R, 255 - Item(Hotbar(I).Slot).G, 255 - Item(Hotbar(I).Slot).B)
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(I).Slot).name) > 0 Then
                    If Spell(Hotbar(I).Slot).Icon > 0 Then
                        ' render normal icon
                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_SpellIcon(Spell(Hotbar(I).Slot).Icon), X + 2, Y + 2, 0, 0, 32, 32, 32, 32
                        ' we got the spell?
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t).Spell > 0 Then
                                If PlayerSpells(t).Spell = Hotbar(I).Slot Then
                                    If SpellCD(t) > 0 Then
                                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                                        RenderTexture Tex_SpellIcon(Spell(Hotbar(I).Slot).Icon), X + 2, Y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
        ' draw the numbers
        sS = str(I)
        If I = 10 Then sS = "0"
        If I = 11 Then sS = " -"
        If I = 12 Then sS = " ="
        RenderText Font_Default, sS, X + 4, Y + 20, White
    Next
End Sub
Public Sub DrawInventory()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long
Dim Amount As String
Dim colour As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_INVENTORY).X, GUIWindow(GUI_INVENTORY).Y, 0, 0, Width, Height, Width, Height
    
    For I = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, I)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    If TradeYourOffer(X).num = I Then
                        GoTo NextLoop
                    End If
                Next
            End If
            
            ' exit out if dragging
            If DragInvSlotNum = I Then GoTo NextLoop

            If ItemPic > 0 And ItemPic <= numitems Then
                Top = GUIWindow(GUI_INVENTORY).Y + InvTop - 2 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                Left = GUIWindow(GUI_INVENTORY).X + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
                ' If item is a stack - draw the amount you have
                If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                    Y = Top + 21
                    X = Left - 4
                    Amount = CStr(GetPlayerInvItemValue(MyIndex, I))
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawInventoryItemDesc()
Dim invSlot As Long, isSB As Boolean
    
    If mouseClicked Then Exit Sub
    
    If Not GUIWindow(GUI_INVENTORY).Visible Then Exit Sub
    If DragInvSlotNum > 0 Then Exit Sub
    
    invSlot = IsInvItem(GlobalX, GlobalY)
    If invSlot > 0 Then
        If GetPlayerInvItemNum(MyIndex, invSlot) > 0 Then
            'If Item(GetPlayerInvItemNum(MyIndex, invSlot)).BindType > 0 And PlayerInv(invSlot).bound > 0 Then isSB = True
            DrawItemDesc GetPlayerInvItemNum(MyIndex, invSlot), GUIWindow(GUI_INVENTORY).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).Y, isSB
            ' value
            If InShop > 0 Then
                DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).Y + GUIWindow(GUI_DESCRIPTION).Height + 85
            End If
        End If
    End If
End Sub

Public Sub DrawShopItemDesc()
Dim shopSlot As Long
    
    If mouseClicked Then Exit Sub
    
    If Not GUIWindow(GUI_SHOP).Visible Then Exit Sub
    
    shopSlot = IsShopItem(GlobalX, GlobalY)
    If shopSlot > 0 Then
        If Shop(InShop).TradeItem(shopSlot).Item > 0 Then
            DrawItemDesc Shop(InShop).TradeItem(shopSlot).Item, GUIWindow(GUI_SHOP).X + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).Y
            DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).X + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).Y + GUIWindow(GUI_DESCRIPTION).Height + 85
        End If
    End If
End Sub

Public Sub DrawCharacterItemDesc()
Dim eqSlot As Long, isSB As Boolean

    If mouseClicked Then Exit Sub
    
    If Not GUIWindow(GUI_CHARACTER).Visible Then Exit Sub
    
    eqSlot = IsEqItem(GlobalX, GlobalY)
    If eqSlot > 0 Then
        If GetPlayerEquipment(MyIndex, eqSlot) > 0 Then
            If Item(GetPlayerEquipment(MyIndex, eqSlot)).BindType > 0 Then isSB = True
            DrawItemDesc GetPlayerEquipment(MyIndex, eqSlot), GUIWindow(GUI_CHARACTER).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_CHARACTER).Y, isSB
        End If
    End If
End Sub

Public Sub DrawItemCost(ByVal isShop As Boolean, ByVal slotNum As Long, ByVal X As Long, ByVal Y As Long)
Dim CostItem As Long, CostValue As Long, itemNum As Long, sString As String, Width As Long, Height As Long

    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    ' draw the window
    Width = 190
    Height = 36

    RenderTexture Tex_GUI(24), X, Y, 0, 0, Width, Height, Width, Height
    
    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        itemNum = GetPlayerInvItemNum(MyIndex, slotNum)
        If itemNum = 0 Then Exit Sub
        CostItem = 1
        CostValue = (Item(itemNum).Price / 100) * Shop(InShop).BuyRate
        sString = "Harga jual"
    Else
        itemNum = Shop(InShop).TradeItem(slotNum).Item
        If itemNum = 0 Then Exit Sub
        CostItem = Shop(InShop).TradeItem(slotNum).CostItem
        CostValue = Shop(InShop).TradeItem(slotNum).CostValue
        sString = "Harga beli"
    End If
    
    'EngineRenderRectangle Tex_Item(Item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(Item(CostItem).Pic), X + 155, Y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
    
    RenderText Font_Default, sString, X + 4, Y + 3, White
    
    RenderText Font_Default, CostValue & " " & Trim$(Item(CostItem).name), X + 4, Y + 18, White
End Sub

Public Sub DrawItemDesc(ByVal itemNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal soulBound As Boolean = False)
Dim colour As Long, descString As String, theName As String, className As String, levelTxt As String, sInfo() As String, I As Long, Width As Long, Height As Long
Dim WeightTxt As String

    ' get out
    If itemNum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Item(itemNum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 210
    End If
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Item(itemNum).Pic > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 64, 64
       ' RenderTexture Tex_Item(Item(itemNum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32
       
       RenderTexture Tex_Item(Item(itemNum).Pic), X + 16, Y + 27, 0, 0, 64, 64, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
    End If
    
    If Not Trim$(Item(itemNum).Desc) = vbNullString Or Not Trim$(Item(itemNum).Desc) = "." Then
        RenderText Font_Default, WordWrap(Trim$(Item(itemNum).Desc), Width - 10), X + 10, Y + 128, White
    End If
    ' work out name colour
    Select Case Item(itemNum).Rarity
        Case 0
            colour = White
        Case 1
            colour = Grey
        Case 2
            colour = White
        Case 3
            colour = Green
        Case 4
            colour = Yellow
        Case 5
            colour = Magenta
    End Select
    
    If Not soulBound Then
        theName = Trim$(Item(itemNum).name)
    Else
        theName = "(SB) " & Trim$(Item(itemNum).name)
    End If
    
    ' render name
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    ' class req
    If Item(itemNum).ClassReq(GetPlayerClass(MyIndex)) > 0 Then
        className = "Can´t use"
        ' do we match it?
        colour = BrightRed
    Else
        className = "Can use"
        colour = Green
    End If
    RenderText Font_Default, className, X + 48 - (EngineGetTextWidth(Font_Default, className) \ 2), Y + 92, colour
    
    ' level
    If Item(itemNum).LevelReq > 0 Then
        levelTxt = "Level " & Item(itemNum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(itemNum).LevelReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        levelTxt = "No level req."
        colour = Green
    End If
    RenderText Font_Default, levelTxt, X + 48 - (EngineGetTextWidth(Font_Default, levelTxt) \ 2), Y + 107, colour
    
    
    ' Weight
    If Item(itemNum).CarryWeight > 0 Then
        WeightTxt = "Weight: " & Item(itemNum).CarryWeight
        colour = Magenta
    End If
    RenderText Font_Default, WeightTxt, X + 95 - (EngineGetTextWidth(Font_Default, WeightTxt) \ 2), Y + 192, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE
            sInfo(I) = "No type"
        Case ITEM_TYPE_WEAPON
            sInfo(I) = "Weapon"
        Case ITEM_TYPE_ARMOR
            sInfo(I) = "Armor"
        Case ITEM_TYPE_HELMET
            sInfo(I) = "Helmet"
        Case ITEM_TYPE_BOOTS
            sInfo(I) = "Boots"
        Case ITEM_TYPE_ENCHANT
            sInfo(I) = "Enchant"
        Case ITEM_TYPE_CHARM
            sInfo(I) = "Charm"
        Case ITEM_TYPE_RING
            sInfo(I) = "Ring"
        Case ITEM_TYPE_WHETSTONE
            sInfo(I) = "Whetstone"
        Case ITEM_TYPE_SHIELD
            sInfo(I) = "Shield"
        Case ITEM_TYPE_CONSUME
            sInfo(I) = "Consume"
        Case ITEM_TYPE_KEY
            sInfo(I) = "Key"
        Case ITEM_TYPE_CURRENCY
            sInfo(I) = "Currency"
        Case ITEM_TYPE_SPELL
            sInfo(I) = "Spell"
        Case ITEM_TYPE_STAT_RESET
            sInfo(I) = "Stat Reset"
        Case ITEM_TYPE_RECIPE
            sInfo(I) = "Recipe"
    End Select
    
    ' more info
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
            ' binding
            If Item(itemNum).BindType = 1 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Pickup"
            ElseIf Item(itemNum).BindType = 2 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Equip"
            End If
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD, ITEM_TYPE_CHARM, ITEM_TYPE_ENCHANT, ITEM_TYPE_BOOTS, ITEM_TYPE_RING, ITEM_TYPE_WHETSTONE
            ' binding
            If Item(itemNum).BindType = 1 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Pickup"
            ElseIf Item(itemNum).BindType = 2 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Bind on Equip"
            End If
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price
            ' damage
            If Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Damage: " & Item(itemNum).Data2
                ' speed
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Speed: " & (Item(itemNum).speed / 1000) & "s"
            End If
            ' defense
            If Item(itemNum).Type <> ITEM_TYPE_WEAPON Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Defense: " & Item(itemNum).Data2
            End If
            ' stat bonuses
            If Item(itemNum).Add_Stat(Stats.strength) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.strength) & " Str"
            End If
            If Item(itemNum).Add_Stat(Stats.Endurance) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(itemNum).Add_Stat(Stats.Intelligence) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(itemNum).Add_Stat(Stats.Agility) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(itemNum).Add_Stat(Stats.Willpower) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
            
            ' stat requirements
            If Item(itemNum).Stat_Req(Stats.strength) > 0 Then
            If Item(itemNum).Stat_Req(Stats.strength) > Player(MyIndex).Stat(Stats.strength) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.strength) & " Str"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Endurance) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Endurance) > Player(MyIndex).Stat(Stats.Endurance) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Endurance) & " End"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Intelligence) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Intelligence) > Player(MyIndex).Stat(Stats.Intelligence) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Intelligence) & " Int"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Agility) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Agility) > Player(MyIndex).Stat(Stats.Agility) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Agility) & " Agi"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Willpower) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Willpower) > Player(MyIndex).Stat(Stats.Willpower) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Willpower) & " Will"
            End If
            End If
            
                
        Case ITEM_TYPE_CONSUME
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price
            If Item(itemNum).CastSpell > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Casts Spell"
            End If
            If Item(itemNum).AddHP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddHP & " HP"
            End If
            If Item(itemNum).AddMP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddMP & " SP"
            End If
            If Item(itemNum).AddEXP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddEXP & " EXP"
            End If
        Case ITEM_TYPE_SPELL
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Value: " & Item(itemNum).Price
            
            ' stat requirements
            If Item(itemNum).Stat_Req(Stats.strength) > 0 Then
            If Item(itemNum).Stat_Req(Stats.strength) > Player(MyIndex).Stat(Stats.strength) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.strength) & " Str"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Endurance) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Endurance) > Player(MyIndex).Stat(Stats.Endurance) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Endurance) & " End"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Intelligence) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Intelligence) > Player(MyIndex).Stat(Stats.Intelligence) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Intelligence) & " Int"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Agility) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Agility) > Player(MyIndex).Stat(Stats.Agility) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Agility) & " Agi"
            End If
            End If
            
            If Item(itemNum).Stat_Req(Stats.Willpower) > 0 Then
            If Item(itemNum).Stat_Req(Stats.Willpower) > Player(MyIndex).Stat(Stats.Willpower) Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Req: " & Item(itemNum).Stat_Req(Stats.Willpower) & " Will"
            End If
            End If
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For I = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(I), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(I)) \ 2), Y, White
    Next
End Sub
Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
    
    If mouseClicked Then Exit Sub
    
    If Not GUIWindow(GUI_SPELLS).Visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If PlayerSpells(spellSlot).Spell > 0 Then
            DrawSpellDesc PlayerSpells(spellSlot).Spell, GUIWindow(GUI_SPELLS).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_SPELLS).Y, spellSlot
        End If
    End If
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal spellSlot As Long = 0)
Dim colour As Long, theName As String, BuffType As String, sUse As String, sInfo() As String, I As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, Height As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If spellnum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not LenB(Trim$(Spell(spellnum).Desc)) = 0 Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(29), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(34), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Spell(spellnum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        RenderTexture Tex_SpellIcon(Spell(spellnum).Icon), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not LenB(Trim$(Spell(spellnum).Desc)) = 0 Then
        RenderText Font_Default, WordWrap(Trim$(Spell(spellnum).Desc), Width - 20), X + 10, Y + 128, White
    End If
    
    ' render name
    colour = White
    theName = Trim$(Spell(spellnum).name)
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    ' if it's a player spell then do the rank up message
    colour = White
    If spellSlot > 0 Then
        ' draw the rank bar
        barWidth = 78
        If Spell(spellnum).NextRank > 0 Then
            tmpWidth = ((PlayerSpells(spellSlot).Uses / barWidth) / (Spell(spellnum).NextUses / barWidth)) * barWidth
        Else
            tmpWidth = 78
        End If
        'EngineRenderRectangle Tex_GUI(35), x + 9, y + 99, 0, 0, tmpWidth, 16, tmpWidth, 16, tmpWidth, 16
        RenderTexture Tex_GUI(35), X + 9, Y + 99, 0, 0, tmpWidth, 16, tmpWidth, 16
        ' does it rank up?
        If Spell(spellnum).NextRank > 0 Then
            sUse = "Uses: " & PlayerSpells(spellSlot).Uses & "/" & Spell(spellnum).NextUses
            If PlayerSpells(spellSlot).Uses = Spell(spellnum).NextUses Then
                If Not GetPlayerLevel(MyIndex) >= Spell(Spell(spellnum).NextRank).LevelReq Then
                    colour = BrightRed
                    sUse = "Lvl " & Spell(Spell(spellnum).NextRank).LevelReq & " req."
                End If
            End If
        Else
            colour = White
            sUse = "Max Rank"
        End If
        RenderText Font_Default, sUse, X + 48 - (EngineGetTextWidth(Font_Default, sUse) \ 2), Y + 99, colour
    End If
    
    
    ' first we cache all information strings then loop through and render them

    ' item type
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP
            sInfo(I) = "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            sInfo(I) = "Damage SP"
        Case SPELL_TYPE_HEALHP
            sInfo(I) = "Heal HP"
        Case SPELL_TYPE_HEALMP
            sInfo(I) = "Heal SP"
        Case SPELL_TYPE_WARP
            sInfo(I) = "Warp"
    End Select
    
    Select Case Spell(spellnum).Type
            
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Vital: " & Spell(spellnum).Vital
            
            ' mp cost
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).AoE > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "AoE: " & Spell(spellnum).AoE
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
            End If
            
            ' blind
            If Spell(spellnum).BlindDuration > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Blind: " & Spell(spellnum).BlindDuration & "s"
            End If
            
            ' stealth
            If Spell(spellnum).StealthDuration > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Stealth: " & Spell(spellnum).StealthDuration & "s"
            End If
            
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For I = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(I), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(I)) \ 2), Y, White
    Next
End Sub

Public Sub DrawSkills()
Dim I As Long, X As Long, Y As Long, spellnum As Long, spellpic As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y, 0, 0, Width, Height, Width, Height
    
    ' render skills
    For I = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(I).Spell

        ' make sure not dragging it
        If DragSpell = I Then GoTo NextLoop
        
        ' actually render
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellpic = Spell(spellnum).Icon

            If spellpic > 0 And spellpic <= NumSpellIcons Then
                Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
                If SpellCD(I) > 0 Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawEquipment()
Dim X As Long, Y As Long, I As Long, Right As Long, Bottom As Long
Dim itemNum As Long, ItemPic As DX8TextureRec

    For I = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, I)

        
        Y = GUIWindow(GUI_CHARACTER).Y + EqTop + ((EqOffsetY + 32) * ((I - 1) \ EqColumns)) - 21
        X = GUIWindow(GUI_CHARACTER).X + EqLeft + ((EqOffsetX + 32) * (((I - 1) Mod EqColumns))) + 6
       ' Right = X + PIC_Y
        ' Bottom = Y + PIC_X + 32
        
         If itemNum > 0 Then
            ItemPic = Tex_Item(Item(itemNum).Pic)
            RenderTexture ItemPic, X, Y, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
        Else
            ' no item equiped - use blank image
            If I = 1 Then ItemPic = Tex_GUI(28)
            If I = 2 Then ItemPic = Tex_GUI(11)
            If I = 3 Then ItemPic = Tex_GUI(31)
            If I = 4 Then ItemPic = Tex_GUI(9)
            If I = 5 Then ItemPic = Tex_GUI(10)
            If I = 6 Then ItemPic = Tex_GUI(12)
            If I = 7 Then ItemPic = Tex_GUI(29)
            If I = 8 Then ItemPic = Tex_GUI(32)
            If I = 9 Then ItemPic = Tex_GUI(30)
            'ItemPic = Tex_GUI(I + 8)
            RenderTexture ItemPic, X, Y, 0, 0, 32, 32, 32, 32
        End If
    Next
    
 'With rec_pos
'.Top = EqTop + ((EqOffsetY + 32) * ((i - 1) \ EqColumns))
'.Bottom = .Top + PIC_Y
'.Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
'.Right = .Left + PIC_X
'End With

End Sub

Public Sub DrawCharacter()
Dim X As Long, Y As Long, I As Long, dX As Long, dY As Long, tmpString As String, buttonnum As Long
Dim Width As Long, Height As Long
Dim WeightLimit As String
Dim ClassString As String

GetPlayerCarryWeight (MyIndex)

    
    X = GUIWindow(GUI_CHARACTER).X
    Y = GUIWindow(GUI_CHARACTER).Y
    
    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(5), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(6), X, Y, 0, 0, Width, Height, Width, Height
    
    ' render name
    tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    RenderText Font_Default, tmpString, X + 7 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), Y + 9, White
    
    ' render Carry Limit
    WeightLimit = "Stamina: " & PlayerWeightLimit
    RenderText Font_Default, WeightLimit, X + 15, Y + 30, White
    
    ' render class
    ClassString = "Class: " & Trim$(Class(GetPlayerClass(MyIndex)).name)
    RenderText Font_Default, ClassString, X + 14, Y + 43, White
    
    ' render stats
    dX = X + 20
    dY = Y + 75
    RenderText Font_Default, "Str: " & GetPlayerStat(MyIndex, strength), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "End: " & GetPlayerStat(MyIndex, Endurance), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Int: " & GetPlayerStat(MyIndex, Intelligence), dX, dY, White
    dY = Y + 75
    dX = dX + 80
    RenderText Font_Default, "Agi: " & GetPlayerStat(MyIndex, Agility), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Will: " & GetPlayerStat(MyIndex, Willpower), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Pnts: " & GetPlayerPOINTS(MyIndex), dX, dY, White
    
    ' draw the face
    'If GetPlayerSprite(MyIndex) > 0 And GetPlayerSprite(MyIndex) <= NumFaces Then
        'EngineRenderRectangle Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
       ' RenderTexture Tex_Face(GetPlayerSprite(MyIndex)), X + 49, Y + 38, 0, 0, 96, 96, 96, 96
   ' End If
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        ' draw the buttons
        For buttonnum = 16 To 20
            X = GUIWindow(GUI_CHARACTER).X + Buttons(buttonnum).X
            Y = GUIWindow(GUI_CHARACTER).Y + Buttons(buttonnum).Y
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                Width = Buttons(buttonnum).Width
                Height = Buttons(buttonnum).Height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
    
    ' draw the equipment
    DrawEquipment
End Sub

Public Sub DrawOptions()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(24), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(21), GUIWindow(GUI_OPTIONS).X, GUIWindow(GUI_OPTIONS).Y, 0, 0, Width, Height, Width, Height
    
    
    
    ' draw buttons
    For I = 26 To 31
        ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    
End Sub

Public Sub DrawParty()
Dim I As Long, X As Long, Y As Long, Width As Long, playerNum As Long, theName As String
Dim Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(7), GUIWindow(GUI_PARTY).X, GUIWindow(GUI_PARTY).Y, 0, 0, Width, Height, Width, Height
    
    ' draw the bars
    If Party.Leader > 0 Then ' make sure we're in a party
        ' draw leader
        playerNum = Party.Leader
        ' name
        theName = Trim$(GetPlayerName(playerNum))
        ' draw name
        Y = GUIWindow(GUI_PARTY).Y + 12
        X = GUIWindow(GUI_PARTY).X + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
        RenderText Font_Default, theName, X, Y, White
        ' draw hp
        Y = GUIWindow(GUI_PARTY).Y + 29
        X = GUIWindow(GUI_PARTY).X + 6
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
        End If
        'EngineRenderRectangle Tex_GUI(13), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(13), X, Y, 0, 0, Width, 9, Width, 9
        ' draw mp
        Y = GUIWindow(GUI_PARTY).Y + 38
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
        End If
        'EngineRenderRectangle Tex_GUI(14), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(14), X, Y, 0, 0, Width, 9, Width, 9
        
        ' draw members
        For I = 1 To MAX_PARTY_MEMBERS
            If Party.Member(I) > 0 Then
                If Party.Member(I) <> Party.Leader Then
                    ' cache the index
                    playerNum = Party.Member(I)
                    ' name
                    theName = Trim$(GetPlayerName(playerNum))
                    ' draw name
                    Y = GUIWindow(GUI_PARTY).Y + 12 + ((I - 1) * 49)
                    X = GUIWindow(GUI_PARTY).X + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
                    RenderText Font_Default, theName, X, Y, White
                    ' draw hp
                    Y = GUIWindow(GUI_PARTY).Y + 29 + ((I - 1) * 49)
                    X = GUIWindow(GUI_PARTY).X + 6
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(13), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(13), X, Y, 0, 0, Width, 9, Width, 9
                    ' draw mp
                    Y = GUIWindow(GUI_PARTY).Y + 38 + ((I - 1) * 49)
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(14), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(14), X, Y, 0, 0, Width, 9, Width, 9
                End If
            End If
        Next
    End If
    
    ' draw buttons
    For I = 24 To 25
        ' set co-ordinate
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub
Public Sub DrawCurrency()
Dim X As Long, Y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_CURRENCY).X
    Y = GUIWindow(GUI_CURRENCY).Y
    ' render chatbox
    Width = GUIWindow(GUI_CURRENCY).Width
    Height = GUIWindow(GUI_CURRENCY).Height
    RenderTexture Tex_GUI(27), X, Y, 0, 0, Width, Height, Width, Height
    Width = EngineGetTextWidth(Font_Default, CurrencyText)
    RenderText Font_Default, CurrencyText, X + 87 + (123 - (Width / 2)), Y + 40, White
    RenderText Font_Default, sDialogue & chatShowLine, X + 90, Y + 65, White
    
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyAcceptState = 2 Then
        ' clicked
        RenderText Font_Default, "[Accept]", X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
            ' hover
            RenderText Font_Default, "[Accept]", X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 1 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 1
            End If
        Else
            ' normal
            RenderText Font_Default, "[Accept]", X, Y, Green
            ' reset sound if needed
            If lastNpcChatsound = 1 Then lastNpcChatsound = 0
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyCloseState = 2 Then
        ' clicked
        RenderText Font_Default, "[Close]", X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
            ' hover
            RenderText Font_Default, "[Close]", X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 2 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 2
            End If
        Else
            ' normal
            RenderText Font_Default, "[Close]", X, Y, Yellow
            ' reset sound if needed
            If lastNpcChatsound = 2 Then lastNpcChatsound = 0
        End If
    End If
End Sub
Public Sub DrawDialogue()
Dim I As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_DIALOGUE).X
    Y = GUIWindow(GUI_DIALOGUE).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_DIALOGUE).Width
    Height = GUIWindow(GUI_DIALOGUE).Height
    RenderTexture Tex_GUI(19), X, Y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(Dialogue_TitleCaption, 392), X + 10, Y + 10, White
    RenderText Font_Default, WordWrap(Dialogue_TextCaption, 392), X + 10, Y + 25, White
    
    If Dialogue_ButtonVisible(1) Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 90
            If Dialogue_ButtonState(1) = 2 Then
                ' clicked
                RenderText Font_Default, "[Accept]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
                    ' hover
                    RenderText Font_Default, "[Accept]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Accept]", X, Y, Green
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(2) Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
            If Dialogue_ButtonState(2) = 2 Then
                ' clicked
                RenderText Font_Default, "[Okay]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
                    ' hover
                    RenderText Font_Default, "[Okay]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Okay]", X, Y, BrightRed
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(3) Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If Dialogue_ButtonState(3) = 2 Then
            ' clicked
            RenderText Font_Default, "[Close]", X, Y, Grey
        Else
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
                ' hover
                RenderText Font_Default, "[Close]", X, Y, Cyan
                ' play sound if needed
                If Not lastNpcChatsound = 3 Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = 3
                End If
            Else
                ' normal
                RenderText Font_Default, "[Close]", X, Y, Yellow
                ' reset sound if needed
                If lastNpcChatsound = 3 Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub
Public Sub DrawShop()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = GUIWindow(GUI_SHOP).Width
    Height = GUIWindow(GUI_SHOP).Height
    'EngineRenderRectangle Tex_GUI(23), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(20), GUIWindow(GUI_SHOP).X, GUIWindow(GUI_SHOP).Y, 0, 0, Width, Height, Width, Height
    
    ' render the shop items
    For I = 1 To MAX_TRADES
        itemNum = Shop(InShop).TradeItem(I).Item
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= numitems Then
                
                Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(I).ItemValue > 1 Then
                    Y = Top + 22
                    X = Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(I).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                End If
            End If
        End If
    Next
    
    ' draw buttons
    For I = 23 To 23
        ' set co-ordinate
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    
    ' draw item descriptions
    DrawShopItemDesc
End Sub
Public Sub DrawMenu()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' draw background
    X = GUIWindow(GUI_MENU).X
    Y = GUIWindow(GUI_MENU).Y
    Width = GUIWindow(GUI_MENU).Width
    Height = GUIWindow(GUI_MENU).Height
    'EngineRenderRectangle Tex_GUI(3), x, y, 0, 0, width, height, width, height, width, height
    'RenderTexture Tex_GUI(3), X, Y, 0, 0, Width, Height, Width, Height
    
    ' draw buttons
    For I = 1 To 6
        If Buttons(I).Visible Then
            ' set co-ordinate
            X = GUIWindow(GUI_MENU).X + Buttons(I).X
            Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
            Width = Buttons(I).Width
            Height = Buttons(I).Height
            ' check for state
            If Buttons(I).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) And Not mouseClicked Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = I
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = I Then lastButtonSound = 0
            End If
        End If
    Next
    
    ' draw other buttons
    For I = 42 To 44
        If Buttons(I).Visible Then
            ' set co-ordinate
            X = GUIWindow(GUI_MENU).X + Buttons(I).X
            Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
            Width = Buttons(I).Width
            Height = Buttons(I).Height
            ' check for state
            If Buttons(I).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) And Not mouseClicked Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = I
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = I Then lastButtonSound = 0
            End If
        End If
    Next
End Sub


Public Sub DrawBank()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_BANK).Width
    Height = GUIWindow(GUI_BANK).Height
    
    RenderTexture Tex_GUI(26), GUIWindow(GUI_BANK).X, GUIWindow(GUI_BANK).Y, 0, 0, Width, Height, Width, Height
    
    ' render the bank items' are you serous? that is it??? maybe... one sec :D :Polol
        For I = 1 To MAX_BANK
            itemNum = GetBankItemNum(I)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                        
                     Top = GUIWindow(GUI_BANK).Y + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                     Left = GUIWindow(GUI_BANK).X + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))

                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
                       
                    ' If the bank item is in a stack, draw the amount...
                    If GetBankItemValue(I) > 1 Then
                        Y = Top + 22
                        X = Left - 4
                        Amount = CStr(GetBankItemValue(I))
                            
                        ' Draw the currency
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                    
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                    End If
                End If
            End If
    Next
    
        If mouseClicked Then Exit Sub
            
             DrawBankItemDesc
                            
                        
End Sub
Public Sub DrawBankItemDesc()
Dim bankNum As Long

    If mouseClicked Then Exit Sub

    If Not GUIWindow(GUI_BANK).Visible Then Exit Sub
        
        bankNum = IsBankItem(GlobalX, GlobalY)
     
        
    If bankNum > 0 Then
        If bankNum > 0 Then
            If GetBankItemNum(bankNum) > 0 Then
                DrawItemDesc GetBankItemNum(bankNum), GUIWindow(GUI_BANK).X + 480, GUIWindow(GUI_BANK).Y
           End If
        End If
    End If
            
End Sub

Public Sub DrawTrade()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_TRADE).Width
    Height = GUIWindow(GUI_TRADE).Width
    RenderTexture Tex_GUI(18), GUIWindow(GUI_TRADE).X, GUIWindow(GUI_TRADE).Y, 0, 0, Width, Height, Width, Height
        For I = 1 To MAX_INV
            ' render your offer
            itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                     RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(I).Value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                            
                        Amount = CStr(TradeYourOffer(I).Value)
                            
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                    End If
                End If
            End If
            
            ' draw their offer
            itemNum = TradeTheirOffer(I).num
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 257 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255 - Item(itemNum).a, 255 - Item(itemNum).R, 255 - Item(itemNum).G, 255 - Item(itemNum).B)
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(I).Value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                                
                        Amount = CStr(TradeTheirOffer(I).Value)
                                
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                    End If
                End If
            End If
        Next
        ' draw buttons
    For I = 40 To 41
        ' set co-ordinate
        X = Buttons(I).X
        Y = Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) And Not mouseClicked Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    RenderText Font_Default, "Your worth: " & YourWorth, GUIWindow(GUI_TRADE).X + 21, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, "Their worth: " & TheirWorth, GUIWindow(GUI_TRADE).X + 250, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, TradeStatus, (GUIWindow(GUI_TRADE).Width / 2) - (EngineGetTextWidth(Font_Default, TradeStatus) / 2), GUIWindow(GUI_TRADE).Y + 317, Yellow
    DrawTradeItemDesc
End Sub

Public Sub DrawTradeItemDesc()
Dim tradeNum As Long
Dim theirtradeNum As Long

    If mouseClicked Then Exit Sub

    If Not GUIWindow(GUI_TRADE).Visible Then Exit Sub
        
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    theirtradeNum = IsTradeItem(GlobalX, GlobalY, False)
    If tradeNum > 0 Then
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).num) > 0 Then
            DrawItemDesc GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).num), GUIWindow(GUI_TRADE).X + 480 + 10, GUIWindow(GUI_TRADE).Y
        End If
    End If
    If theirtradeNum > 0 Then
        If TradeTheirOffer(theirtradeNum).num > 0 Then
            DrawItemDesc (TradeTheirOffer(theirtradeNum).num), GUIWindow(GUI_TRADE).X + 480 + 10, GUIWindow(GUI_TRADE).Y
        End If
    End If
End Sub

Public Sub DrawGUIBars()
Dim tmpWidth As Long, barWidth As Long, X As Long, Y As Long, dX As Long, dY As Long, sString As String, tmpString As String
Dim Width As Long, Height As Long


    ' backwindow + empty bars
    X = GUIWindow(GUI_BARS).X
    Y = GUIWindow(GUI_BARS).Y
    Width = 254
    Height = 110
    
    
    'EngineRenderRectangle Tex_GUI(4), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(4), X, Y, 0, 0, Width, Height, Width, Height
    
            ' draw the face
   If GetPlayerSprite(MyIndex) > 0 And GetPlayerSprite(MyIndex) <= NumFaces Then
        'EngineRenderRectangle Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
       RenderTexture Tex_Face(GetPlayerSprite(MyIndex)), X + 15, Y + 30, 0, 0, 50, 50, 90, 90
   End If
    
    ' hardcoded for POT textures
    barWidth = 250
    
        ' render name
    tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    RenderText Font_Default, tmpString, X + 90 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), Y + 70, White

    ' health bar
    BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(13), X + 7, Y + 9, 0, 0, BarWidth_GuiHP, Tex_GUI(13).Height, BarWidth_GuiHP, Tex_GUI(13).Height
    ' render health
    sString = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    dX = X + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 9
    RenderText Font_Default, sString, dX, dY, White
    
    ' spirit bar
    BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(14), X + 7, Y + 31, 0, 0, BarWidth_GuiSP, Tex_GUI(14).Height, BarWidth_GuiSP, Tex_GUI(14).Height
    ' render spirit
    sString = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    dX = X + 80 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 31
    RenderText Font_Default, sString, dX, dY, White
    
    ' exp bar
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
    Else
        BarWidth_GuiEXP_Max = barWidth
    End If
    RenderTexture Tex_GUI(15), X + 7, Y + 53, 0, 0, BarWidth_GuiEXP, Tex_GUI(15).Height, BarWidth_GuiEXP, Tex_GUI(15).Height
    ' render exp
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        sString = Int(GetPlayerExp(MyIndex) / TNL * 100) & "%"
    Else
        sString = "Max Level"
    End If
    dX = X + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 53
    RenderText Font_Default, sString, dX, dY, White
End Sub
Public Sub DrawEventChat()
Dim I As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_EVENTCHAT).X
    Y = GUIWindow(GUI_EVENTCHAT).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_EVENTCHAT).Width
    Height = GUIWindow(GUI_EVENTCHAT).Height
    RenderTexture Tex_GUI(19), X, Y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(chatText, GUIWindow(GUI_EVENTCHAT).Width - 20), X + 10, Y + 22, White
    
    If chatOnlyContinue = False Then
        ' Draw replies
        For I = 1 To 4
            If Len(Trim$(chatOpt(I))) > 0 Then
                Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(I)) & "]")
               ' X = GUIWindow(GUI_CHAT).X + 95 + (155 - (Width / 2))
                'X = GUIWindow(GUI_CHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                Y = GUIWindow(GUI_CHAT).Y + 70 + ((I - 1) * 15)
                If chatOptState(I) = 2 Then
                    ' clicked
                    RenderText Font_Default, "[" & Trim$(chatOpt(I)) & "]", X, Y, Grey
                Else
                    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
                    'If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                        ' hover
                        RenderText Font_Default, "[" & Trim$(chatOpt(I)) & "]", X, Y, Yellow
                        ' play sound if needed
                        If Not lastNpcChatsound = I Then
                            PlaySound Sound_ButtonHover, -1, -1
                            lastNpcChatsound = I
                        End If
                    Else
                        ' normal
                        RenderText Font_Default, "[" & Trim$(chatOpt(I)) & "]", X, Y, BrightBlue
                        ' reset sound if needed
                        If lastNpcChatsound = I Then lastNpcChatsound = 0
                    End If
                End If
            End If
        Next
    Else
        Width = EngineGetTextWidth(Font_Default, "[Continue]")
        X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
        Y = GUIWindow(GUI_EVENTCHAT).Y + 100
        If chatContinueState = 2 Then
            ' clicked
            RenderText Font_Default, "[Continue]", X, Y, Grey
        Else
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) And Not mouseClicked Then
                ' hover
                RenderText Font_Default, "[Continue]", X, Y, Yellow
                ' play sound if needed
                If Not lastNpcChatsound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = I
                End If
            Else
                ' normal
                RenderText Font_Default, "[Continue]", X, Y, BrightBlue
                ' reset sound if needed
                If lastNpcChatsound = I Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub

Public Sub DrawGuildMenu()
Dim Width As Long, Height As Long, X As Long, Y As Long, I As Long
    Width = GUIWindow(GUI_GUILD).Width
    Height = GUIWindow(GUI_GUILD).Height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_GUILD).X, GUIWindow(GUI_GUILD).Y, 0, 0, Width, Height, Width, Height
    RenderText Font_Default, "Guild Menu", GUIWindow(GUI_GUILD).X + (GUIWindow(GUI_GUILD).Width / 2) - (EngineGetTextWidth(Font_Default, "Guild Menu") / 2), GUIWindow(GUI_GUILD).Y + 5, White
    
    If Len(Trim$(GuildData.Guild_Name)) > 0 Then
        RenderText Font_Georgia, "Guild Name: " & Trim$(GuildData.Guild_Name), GUIWindow(GUI_GUILD).X + (GUIWindow(GUI_GUILD).Width / 2) - (EngineGetTextWidth(Font_Default, "Guild Name: " & Trim$(GuildData.Guild_Name)) / 2), GUIWindow(GUI_GUILD).Y + 25, White
        
        RenderText Font_Georgia, "Guild Tag: " & Trim$(GuildData.Guild_Tag), GUIWindow(GUI_GUILD).X + (GUIWindow(GUI_GUILD).Width / 2) - (EngineGetTextWidth(Font_Default, "Guild Tag: " & Trim$(GuildData.Guild_Tag)) / 2), GUIWindow(GUI_GUILD).Y + 39, GuildData.Guild_Color
        
        RenderText Font_Georgia, "Message Of The Day:", GUIWindow(GUI_GUILD).X + (GUIWindow(GUI_GUILD).Width / 2) - (EngineGetTextWidth(Font_Default, "Message Of The Day:") / 2), GUIWindow(GUI_GUILD).Y + 56, White
        
        RenderText Font_Georgia, WordWrap(Trim$(GuildData.Guild_MOTD), GUIWindow(GUI_GUILD).Width - 20), GUIWindow(GUI_GUILD).X + 10, GUIWindow(GUI_GUILD).Y + 70, Yellow
        
        RenderText Font_Georgia, "- Guild Members -", GUIWindow(GUI_GUILD).X + (GUIWindow(GUI_GUILD).Width / 2) - (EngineGetTextWidth(Font_Default, "- Guild Members -") / 2), GUIWindow(GUI_GUILD).Y + 125, White
            
        If Not Player(MyIndex).GuildName = vbNullString Then
            For I = 1 To MAX_GUILD_MEMBERS
                If I > GuildScroll - (I - GuildScroll) - 2 And I < GuildScroll + 5 Then
                    If Not GuildData.Guild_Members(I).User_Name = vbNullString Then
                        If GuildData.Guild_Members(I).Online = True Then
                            RenderText Font_Default, "-  " & GuildData.Guild_Members(I).User_Name, GUIWindow(GUI_GUILD).X + 25, GUIWindow(GUI_GUILD).Y + 130 + ((I - GuildScroll) * 14), Green
                        Else
                            RenderText Font_Default, "-  " & GuildData.Guild_Members(I).User_Name, GUIWindow(GUI_GUILD).X + 25, GUIWindow(GUI_GUILD).Y + 130 + ((I - GuildScroll) * 14), Red
                        End If
                    End If
                End If
            Next I
        End If
        ' draw buttons
        For I = 45 To 46
            ' set co-ordinate
            X = GUIWindow(GUI_GUILD).X + Buttons(I).X
            Y = GUIWindow(GUI_GUILD).Y + Buttons(I).Y
            Width = Buttons(I).Width
            Height = Buttons(I).Height
            ' check for state
            If Buttons(I).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = I
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = I Then lastButtonSound = 0
            End If
        Next
    Else
        RenderText Font_Georgia, "You are not in a Guild.", GUIWindow(GUI_GUILD).X + (GUIWindow(GUI_GUILD).Width / 2) - (EngineGetTextWidth(Font_Default, "You are not in a Guild.") / 2), GUIWindow(GUI_GUILD).Y + 25, White
    End If
End Sub

Public Sub DrawQuestLog()
Dim buttonnum As Long, X As Long, Y As Long
Dim Width As Long, Height As Long
    ' render the window
    Width = GUIWindow(GUI_QUESTLOG).Width
    Height = GUIWindow(GUI_QUESTLOG).Height
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_QUESTLOG).X, GUIWindow(GUI_QUESTLOG).Y, 0, 0, Width, Height, Width, Height
    RenderText Font_Default, "Quest log", GUIWindow(GUI_QUESTLOG).X + (GUIWindow(GUI_QUESTLOG).Width / 2) - (EngineGetTextWidth(Font_Default, "Quest log") / 2), GUIWindow(GUI_QUESTLOG).Y + 5, White
    ' draw the buttons
        For buttonnum = 53 To 58
            X = GUIWindow(GUI_QUESTLOG).X + Buttons(buttonnum).X
            Y = GUIWindow(GUI_QUESTLOG).Y + Buttons(buttonnum).Y
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                Width = Buttons(buttonnum).Width
                Height = Buttons(buttonnum).Height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
End Sub
Public Sub DrawQuestDialogue()
Dim I As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_QUESTDIALOGUE).X
    Y = GUIWindow(GUI_QUESTDIALOGUE).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_QUESTDIALOGUE).Width
    Height = GUIWindow(GUI_QUESTDIALOGUE).Height
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(19), X, Y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(QuestName, Width - 20), X + (Width / 2) - (EngineGetTextWidth(Font_Default, QuestName) / 2), Y + 10, White
    RenderText Font_Georgia, WordWrap(QuestSubtitle, Width - 20), X + 10, Y + 25, White
    RenderText Font_Georgia, WordWrap(QuestSay, Width - 20), X + 10, Y + 40, White
    
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "Accept")
        X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        Y = GUIWindow(GUI_QUESTDIALOGUE).Y + 106
            If QuestAcceptState = 2 Then
                ' clicked
                RenderText Font_Georgia, ">>>Accept<<<", X - EngineGetTextWidth(Font_Georgia, ">>>"), Y, DarkGrey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Georgia, ">>>Accept<<<", X - EngineGetTextWidth(Font_Georgia, ">>>"), Y, White
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, "Accept", X, Y, White
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, QuestExtra)
        X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        Y = GUIWindow(GUI_QUESTDIALOGUE).Y + 107
            If QuestExtraState = 2 Then
                ' clicked
                RenderText Font_Georgia, ">>>" & QuestExtra & "<<<", X - EngineGetTextWidth(Font_Georgia, ">>>"), Y, DarkGrey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Georgia, ">>>" & QuestExtra & "<<<", X - EngineGetTextWidth(Font_Georgia, ">>>"), Y, White
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Georgia, QuestExtra, X, Y, White
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    Width = EngineGetTextWidth(Font_Georgia, "Close")
    X = (GUIWindow(GUI_QUESTDIALOGUE).X + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    Y = GUIWindow(GUI_QUESTDIALOGUE).Y + 120
    If QuestCloseState = 2 Then
        ' clicked
        RenderText Font_Georgia, ">>>Close<<<", X - EngineGetTextWidth(Font_Georgia, ">>>"), Y, DarkGrey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Georgia, ">>>Close<<<", X - EngineGetTextWidth(Font_Georgia, ">>>"), Y, White
            ' play sound if needed
            If Not lastNpcChatsound = 3 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 3
            End If
        Else
            ' normal
            RenderText Font_Georgia, "Close", X, Y, White
            ' reset sound if needed
            If lastNpcChatsound = 3 Then lastNpcChatsound = 0
        End If
    End If
End Sub

Public Sub EditorItem_DrawProjectile()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_Item.scrlProjectilePic.Value

    If itemNum < 1 Or itemNum > NumProjectiles Then
        frmEditor_Item.picProjectile.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    drect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Projectile(itemNum), sRECT, drect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picProjectile.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawProjectile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawProjectile()
Dim Angle As Long, X As Long, Y As Long, I As Long
    If LastProjectile > 0 Then
        
        ' ****** Create Particle ******
        For I = 1 To LastProjectile
            With ProjectileList(I)
                If .Graphic Then
                
                    ' ****** Update Position ******
                    Angle = DegreeToRadian * Engine_GetAngle(.X, .Y, .tx, .ty)
                    .X = .X + (Sin(Angle) * ElapsedTime * 0.6)
                    .Y = .Y - (Cos(Angle) * ElapsedTime * 0.6)
                    X = .X
                    Y = .Y
                    
                    ' ****** Update Rotation ******
                    If .RotateSpeed > 0 Then
                        .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.01)
                        Do While .Rotate > 360
                            .Rotate = .Rotate - 360
                        Loop
                    End If
                    
                    ' ****** Render Projectile ******
                    If .Rotate = 0 Then
                        Call RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(X), ConvertMapY(Y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y)
                    Else
                        Call RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(X), ConvertMapY(Y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y, , .Rotate)
                    End If
                    
                End If
            End With
        Next
        
        ' ****** Erase Projectile ******    Seperate Loop For Erasing
        For I = 1 To LastProjectile
            If ProjectileList(I).Graphic Then
                If Abs(ProjectileList(I).X - ProjectileList(I).tx) < 20 Then
                    If Abs(ProjectileList(I).Y - ProjectileList(I).ty) < 20 Then
                        Call ClearProjectile(I)
                    End If
                End If
            End If
        Next
        
    End If
End Sub

Sub DrawNight()
Dim Alpha As Byte, X As Long, Y As Long
If Map.DayNight = 2 Then Exit Sub
'If GameHours >= 20 Or GameHours < 6 Or Map.DayNight = 1 Then
   ' Alpha = 195
    RenderTexture Tex_Lightmap, ConvertMapX(GetPlayerX(MyIndex) * 32) + Player(MyIndex).xOffset + 16 - 1300, ConvertMapY(GetPlayerY(MyIndex) * 32) + Player(MyIndex).yOffset - 812.5, 0, 0, 2600, 1625, 2600, 1625, D3DColorARGB(Alpha, 0, 0, 0)
'End If

'Night/Day
    
    If Map.DayNight = 0 Then
   
    If GameHours >= 0 And GameHours < 1 Then
    Alpha = 245
    End If
    
    If GameHours >= 1 And GameHours < 2 Then
    Alpha = 225
    End If
    
    If GameHours >= 2 And GameHours < 3 Then
    Alpha = 205
    End If
    
    If GameHours >= 3 And GameHours < 4 Then
    Alpha = 185
    End If
    
    If GameHours >= 4 And GameHours < 5 Then
    Alpha = 155
    End If
    
    If GameHours >= 5 And GameHours < 6 Then
    Alpha = 95
    End If
    
    If GameHours >= 10 And GameHours < 18 Then
    Alpha = 0
    End If
    
    If GameHours >= 18 And GameHours < 20 Then
    Alpha = 95
    End If
    
    If GameHours >= 20 And GameHours < 22 Then
    Alpha = 155
    End If
    
    If GameHours >= 22 And GameHours < 24 Then
    Alpha = 185
    End If

    
   End If
    
    ' Always night maps
    If Map.DayNight = 1 Then
    Alpha = 255
    End If
    
     RenderTexture Tex_Lightmap, ConvertMapX(GetPlayerX(MyIndex) * 32) + Player(MyIndex).xOffset + 16 - 1300, ConvertMapY(GetPlayerY(MyIndex) * 32) + Player(MyIndex).yOffset - 812.5, 0, 0, 2600, 1625, 2600, 1625, D3DColorARGB(Alpha, 0, 0, 0)

End Sub

Public Sub EditorMap_DrawLight()
Dim Height As Long, Width As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT

Height = 128
Width = 128
    
    sRECT.Top = 0
    sRECT.Bottom = 128
    sRECT.Left = 0
    sRECT.Right = 128
    
    drect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Light, sRECT, drect, D3DColorARGB(frmEditor_Map.scrlA, frmEditor_Map.scrlR, frmEditor_Map.scrlG, frmEditor_Map.scrlB)
    'DirectX8.RenderTexture Tex_Light, 0, 0, 0, 0, Width, Height, Width, Height
    With destRect
        .X1 = 0
        .X2 = 128
        .Y1 = 0
        .Y2 = 128
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picLight.hWnd, ByVal (0)
    
End Sub

Public Sub DrawLight(ByVal X As Long, ByVal Y As Long, ByVal a As Long, ByVal R As Long, ByVal G As Long, ByVal B As Long)
    'engineRenderRectangle Tex_GUI(19), x, y, 0, 0, width, height, width, height, width, height
   If Options.Debug = 1 Then On Error GoTo errorhandler

    'DirectX8.RenderTexture Tex_Light, ConvertMapX(x) - 48, ConvertMapY(y) - 48, 0, 0, 128, 128, 128, 128, D3DColorARGB(Abs(Int(A) - Rand(0, 25)), Int(R), Int(G), Int(b))
    RenderTexture Tex_Light, ConvertMapX(X) - 48, ConvertMapY(Y) - 48, 0, 0, 128, 128, 128, 128, D3DColorARGB(Abs(Int(a) - Rand(0, 25)), Int(R), Int(G), Int(B))
    'RenderTexture Tex_Lightmap, ConvertMapX(GetPlayerX(MyIndex) * 32) + TempPlayer(MyIndex).xOffset + 16 - 1300, ConvertMapY(GetPlayerY(MyIndex) * 32) + TempPlayer(MyIndex).yOffset - 812.5, 0, 0, 2600, 1625, 2600, 1625, D3DColorARGB(Alpha, 0, 0, 0)
   ' Error handler
   Exit Sub
errorhandler:
    HandleError "DrawLight", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub DrawNewClass()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = GUIWindow(GUI_NEWCLASS).Width
    Height = GUIWindow(GUI_NEWCLASS).Height
    'EngineRenderRectangle Tex_GUI(23), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(33), GUIWindow(GUI_NEWCLASS).X, GUIWindow(GUI_NEWCLASS).Y, 0, 0, Width, Height, Width, Height
    
    LoadClasses
    
    ' draw buttons
    For I = 1 To 3
    If ClassData(I) <> 0 Then
        Width = EngineGetTextWidth(Font_Default, Trim$(Class(ClassData(I)).name))
        Select Case I
            Case 1
                X = (GUIWindow(GUI_NEWCLASS).X + (GUIWindow(GUI_NEWCLASS).Width / 2) - 130) - (Width / 2)
            Case 2
                X = (GUIWindow(GUI_NEWCLASS).X + (GUIWindow(GUI_NEWCLASS).Width / 2)) - (Width / 2)
            Case 3
                X = (GUIWindow(GUI_NEWCLASS).X + (GUIWindow(GUI_NEWCLASS).Width / 2) + 130) - (Width / 2)
        End Select
        Y = GUIWindow(GUI_NEWCLASS).Y + 28
        RenderText Font_Default, "Select class: (" & Trim$(Class(GetPlayerClass(MyIndex)).name) & ")", GUIWindow(GUI_NEWCLASS).X + (GUIWindow(GUI_NEWCLASS).Width / 2) - (EngineGetTextWidth(Font_Default, "Select class: (" & Trim$(Class(GetPlayerClass(MyIndex)).name) & ")") / 2), GUIWindow(GUI_NEWCLASS).Y + 10, Pink
        RenderText Font_Default, WordWrap(Trim$(Class(ClassData(I)).Desc), 120), GUIWindow(GUI_NEWCLASS).X + 10 + (GUIWindow(GUI_NEWCLASS).Width / 3 * (I - 1)), GUIWindow(GUI_NEWCLASS).Y + 46, White
        
        If ClassButtonState(I) = 2 Then
            ' clicked
            RenderText Font_Default, Trim$(Class(ClassData(I)).name), X, Y, Grey
        Else
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                ' hover
                RenderText Font_Default, Trim$(Class(ClassData(I)).name), X, Y, Cyan
                ' play sound if needed
                If Not lastNpcChatsound = I Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = I
                End If
            Else
                ' normal
                RenderText Font_Default, Trim$(Class(ClassData(I)).name), X, Y, Yellow
                ' reset sound if needed
                If lastNpcChatsound = I Then lastNpcChatsound = 0
            End If
        End If
        End If
    Next
End Sub

Public Sub DrawOverlay(Alpha As Byte, Red As Byte, Green As Byte, Blue As Byte)
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub RenderTextureRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    '12x12 tiles
    
    'Corners
    RenderTexture Tex_GUI(2), X, Y, 0, 0, 12, 12, 12, 12
    RenderTexture Tex_GUI(2), X + Width - 12, Y, 24, 0, 12, 12, 12, 12
    RenderTexture Tex_GUI(2), X, Y + Height - 12, 0, 24, 12, 12, 12, 12
    RenderTexture Tex_GUI(2), X + Width - 12, Y + Height - 12, 24, 24, 12, 12, 12, 12
    
    'Vertical Borders
    RenderTexture Tex_GUI(2), X, Y + 12, 0, 12, 12, Height - 24, 12, 12
    RenderTexture Tex_GUI(2), X + Width - 12, Y + 12, 24, 12, 12, Height - 24, 12, 12
    
    'Horizontal Borders
    RenderTexture Tex_GUI(2), X + 12, Y, 12, 0, Width - 24, 12, 12, 12
    RenderTexture Tex_GUI(2), X + 12, Y + Height - 12, 12, 24, Width - 24, 12, 12, 12
    
    'Center
    RenderTexture Tex_GUI(2), X + 12, Y + 12, 12, 12, Width - 24, Height - 24, 12, 12
End Sub

Public Function isAnimated(ByVal Sprite As Long) As Boolean
    isAnimated = False
    Select Case Sprite
        Case 31, 32, 33, 35, 36, 50, 53, 58, 59, 60, 62, 63, 69, 70, 80, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 166, 167, 171, 172, 173, 177, 178, 179, 180, 181, 187
            isAnimated = True
    End Select
End Function

Public Function hasShadow(ByVal Sprite As Long) As Boolean
    hasShadow = True
    Select Case Sprite
        Case 31, 33, 37, 38, 50, 66, 67, 182, 187
            hasShadow = False
    End Select
End Function

Public Sub DrawNews()


Dim I As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long
  ' draw background
  X = GUIWindow(GUI_NEWS).X
  Y = GUIWindow(GUI_NEWS).Y
 
  ' render chatbox
  Width = GUIWindow(GUI_NEWS).Width
  Height = GUIWindow(GUI_NEWS).Height
  'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
  RenderTexture Tex_GUI(37), X, Y, 0, 0, Width, Height, Width, Height
 
  Select Case PlayerInfo.CurrentWindow
      Case 0
             ' Draw the text
      RenderText Font_Default, "Recent News", X + 120, Y + 20, White
      RenderText Font_Default, WordWrap(NewsText, Width - 70), X + 35, Y + 45, Brown
      Case 1
      RenderText Font_Default, "Patch Notes", X + 120, Y + 20, White
      RenderText Font_Default, WordWrap(PatchNotes, Width - 70), X + 35, Y + 45, Brown
  End Select
 
  For I = 59 To 61
      X = GUIWindow(GUI_NEWS).X + Buttons(I).X
      Y = GUIWindow(GUI_NEWS).Y + Buttons(I).Y
      Width = Buttons(I).Width
      Height = Buttons(I).Height
      ' render accept button
      If Buttons(I).state = 2 Then
          ' we're clicked boyo
          'EngineRenderRectangle Tex_Buttons_c(Buttons(I).PicNum), x, y, 0, 0, width, height, width, height, width, height
          RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
      ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
          ' we're hoverin'
          'EngineRenderRectangle Tex_Buttons_h(Buttons(I).PicNum), x, y, 0, 0, width, height, width, height, width, height
          RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
          ' play sound if needed
          If Not lastButtonSound = I Then
              'PlaySound Sound_ButtonHover
              lastButtonSound = I
          End If
      Else
          ' we're normal
          'EngineRenderRectangle Tex_Buttons(Buttons(I).PicNum), x, y, 0, 0, width, height, width, height, width, height
          RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
          ' reset sound if needed
          If lastButtonSound = I Then lastButtonSound = 0
      End If
  Next
 
End Sub
