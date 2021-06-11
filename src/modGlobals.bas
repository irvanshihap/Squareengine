Attribute VB_Name = "modGlobals"
Option Explicit

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

' Game Time
Public GameSeconds As Byte
Public GameMinutes As Byte
Public GameHours As Byte
Public DayTime As Boolean
Public GameSecondsPerSecond As Byte
Public GameMinutesPerMinute As Byte

'- Pathfinding Constant -
'1 is the old method, faster but not smart at all
'2 is the new method, smart but can slow the server down if maps are huge and alot of npcs have targets.
Public PathfindingType As Long

' Used for Double Exp
Public DoubleExp As Boolean

Public AEditorPlayer As String
