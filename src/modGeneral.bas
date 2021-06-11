Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed
    
    GameSecondsPerSecond = 30
    GameMinutesPerMinute = 1
    GameSeconds = Second(Now)
    GameMinutes = Minute(Now)
    GameHours = Hour(Now)

    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\Data\", "accounts"
    ChkDir App.path & "\Data\", "animations"
    ChkDir App.path & "\Data\", "banks"
    ChkDir App.path & "\Data\", "items"
    ChkDir App.path & "\Data\", "logs"
    ChkDir App.path & "\Data\", "maps"
    ChkDir App.path & "\Data\", "npcs"
    ChkDir App.path & "\Data\", "resources"
    ChkDir App.path & "\Data\", "shops"
    ChkDir App.path & "\Data\", "spells"
    ChkDir App.path & "\Data\", "guilds"
    ChkDir App.path & "\Data\", "quests"
    

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.path & "\data\options.ini", True) Then
        Options.Game_Name = "Skywardens"
        Options.Port = 7001
        Options.MOTD = "Welcome to Skywardens!"
        Options.Website = "http://www.skywyre.net"
        Options.Buy_Cost = 5000
        Options.Buy_Lvl = 20
        Options.Buy_Item = 1
        Options.Join_Cost = 1000
        Options.Join_Lvl = 20
        Options.Join_Item = 1
        Options.DayNight = 1
        SaveOptions
    Else
        LoadOptions
    End If
    
    If Options.PathfindingType = 1 Then
    PathfindingType = 1
    End If
    If Options.PathfindingType = 2 Then
    PathfindingType = 2
    End If
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Spawning global events...")
    Call SpawnAllMapGlobalEvents
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    ' Setup Guild ranks
    Call Set_Default_Guild_Ranks

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing guilds...")
    Call ClearGuilds
    Call SetStatus("Clearing quests...")
    Call ClearQuests
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    
    Call SetStatus("Loading rank...")
Call LoadRank
    Call SetStatus("Loading switches...")
    Call LoadSwitches
    Call SetStatus("Loading variables...")
    Call LoadVariables
    Call SetStatus("Loading quests...")
    Call LoadQuests
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function



