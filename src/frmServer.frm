VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAccount 
      Caption         =   "Account Editor"
      Height          =   255
      Left            =   6600
      TabIndex        =   48
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame9 
      Caption         =   "Awesome Sh*t"
      Height          =   2655
      Left            =   6600
      TabIndex        =   46
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton btnDubExp 
         Caption         =   "  Activate    Double Exp"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Global Experience Modifier"
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   9120
      Width           =   4455
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         Left            =   120
         Max             =   10
         TabIndex        =   39
         Top             =   240
         Value           =   1
         Width           =   2535
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         Caption         =   "Exp Mod: x1"
         Height          =   195
         Left            =   2760
         TabIndex        =   40
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Guild Info"
      Height          =   4935
      Left            =   4680
      TabIndex        =   22
      Top             =   4800
      Width           =   1815
      Begin VB.CommandButton cmdGSave 
         Caption         =   "Save Config"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Join config"
         Height          =   2055
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   1575
         Begin VB.TextBox txtGJoinItem 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtGJoinLvl 
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtGJoinCost 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Item:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Level Req:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Cost:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buy Config"
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1575
         Begin VB.TextBox txtGBuyItem 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtGBuyLvl 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtGBuyCost 
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Item:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Level Req:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Cost:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Server Chat"
      Height          =   5175
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   4455
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   4680
         Width           =   4215
      End
      Begin VB.TextBox txtText 
         Height          =   4335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Players"
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   4455
      Begin VB.CommandButton cmdUnadmin 
         Caption         =   "Unadmin"
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdmin 
         Caption         =   "Admin"
         Height          =   255
         Left            =   2160
         TabIndex        =   44
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdBan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton cmdDisc 
         Caption         =   "Disconnect"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox lbPlayers 
         Height          =   2205
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server Cycles Per Second"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2535
      Begin VB.Label lblCPS 
         AutoSize        =   -1  'True
         Caption         =   "CPS: 0"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblCpsLock 
         AutoSize        =   -1  'True
         Caption         =   "[ Unlock ]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Game Time"
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   120
      Width           =   1815
      Begin VB.Label lblGameTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "xx:xx"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   185
         Width           =   1695
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server"
      Height          =   1335
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox chkServerLog 
         Caption         =   "Server Log"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdShutDown 
         Caption         =   "Shut Down"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Reload"
      Height          =   3135
      Left            =   4680
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton cmdReloadAnimations 
         Caption         =   "Animations"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdReloadResources 
         Caption         =   "Resources"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdReloadItems 
         Caption         =   "Items"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdReloadNPCs 
         Caption         =   "Npcs"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdReloadShops 
         Caption         =   "Shops"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton CmdReloadSpells 
         Caption         =   "Spells"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdReloadMaps 
         Caption         =   "Maps"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdReloadClasses 
         Caption         =   "Classes"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuLoadPLayer 
         Caption         =   "Load Player Data"
      End
      Begin VB.Menu mnuModPlayer 
         Caption         =   "Make Monitor"
      End
      Begin VB.Menu mnuMapPlayer 
         Caption         =   "Make Mapper"
      End
      Begin VB.Menu mnuDevPlayer 
         Caption         =   "Make Developer"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Player(1).Switches(1) = 1
End Sub

Private Sub Command2_Click()
    Player(1).Switches(1) = 0
End Sub

Private Sub btnDubExp_Click()
    DoubleExp = Not DoubleExp
    If DoubleExp Then
        Call GlobalMsg("Server: DOUBLE EXP has been activated. Enjoy.", Green)
        Call TextAdd("DOUBLE EXP activated.")
        btnDubExp.Caption = "Deactivate Double Exp"
    Else
        Call GlobalMsg("Server: DOUBLE EXP has been deactivated.", Green)
        Call TextAdd("DOUBLE EXP deactivated.")
        btnDubExp.Caption = "Activate Double Exp"
    End If
End Sub

Private Sub cmdAccount_Click()
LoadAEditor
frmAccount.Show
End Sub

Private Sub cmdAdmin_Click()

    If (lbPlayers.ListIndex > -1) Then
        If (Trim$(Player(lbPlayers.ListIndex + 1).Name) <> vbNullString) Then
            SetPlayerAccess (lbPlayers.ListIndex + 1), 4
            SendPlayerData (lbPlayers.ListIndex + 1)
            PlayerMsg (lbPlayers.ListIndex + 1), "You have been granted administrator access.", BrightCyan
        End If
    End If
End Sub

Private Sub cmdBan_Click()

    If (lbPlayers.ListIndex > -1) Then
        If (Trim$(Player(lbPlayers.ListIndex + 1).Name) <> vbNullString) Then
            ServerBanIndex (lbPlayers.ListIndex + 1)
        End If
    End If
End Sub

Private Sub cmdDisc_Click()
    If (lbPlayers.ListIndex > -1) Then
        If (Trim$(Player(lbPlayers.ListIndex + 1).Name) <> vbNullString) Then
            CloseSocket (lbPlayers.ListIndex + 1)
        End If
    End If
End Sub

Private Sub cmdGSave_Click()
    Options.Buy_Cost = frmServer.txtGBuyCost.Text
    Options.Buy_Lvl = frmServer.txtGBuyLvl.Text
    Options.Buy_Item = frmServer.txtGBuyItem.Text
    Options.Join_Cost = frmServer.txtGJoinCost.Text
    Options.Join_Lvl = frmServer.txtGJoinLvl.Text
    Options.Join_Item = frmServer.txtGJoinItem.Text
    SaveOptions
End Sub

Private Sub cmdUnadmin_Click()
    If (lbPlayers.ListIndex > -1) Then
        If (Trim$(Player(lbPlayers.ListIndex + 1).Name) <> vbNullString) Then
            SetPlayerAccess (lbPlayers.ListIndex + 1), 0
            SendPlayerData (lbPlayers.ListIndex + 1)
            PlayerMsg (lbPlayers.ListIndex + 1), "You have had your administrator access revoked.", BrightRed
        End If
    End If
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim i As Long
    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        frmServer.lbPlayers.AddItem vbNullString

        If i < 10 Then
            frmServer.lbPlayers.List(i) = "00" & i
        ElseIf i < 100 Then
            frmServer.lbPlayers.List(i) = "0" & i
        Else
            frmServer.lbPlayers.List(i) = i
        End If
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub

