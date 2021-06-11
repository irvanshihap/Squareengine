VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   13080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   24915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":014A
   MousePointer    =   99  'Custom
   ScaleHeight     =   872
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1661
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkAutoAttack 
      Caption         =   "Auto attack"
      Height          =   375
      Left            =   9120
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   0
      Left            =   10440
      MousePointer    =   1  'Arrow
      Picture         =   "frmMain.frx":021C
      ScaleHeight     =   1185
      ScaleWidth      =   945
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image cmdMall1 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   975
      End
      Begin VB.Image cmdMall3 
         Height          =   495
         Left            =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Image cmdMall2 
         Height          =   495
         Left            =   0
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      MousePointer    =   1  'Arrow
      Picture         =   "frmMain.frx":60C7
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   28
      Top             =   960
      Width           =   975
      Begin VB.Image cmdIM 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   495
      End
      Begin VB.Image cmdRank 
         Height          =   375
         Left            =   480
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picRank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   2880
      MousePointer    =   1  'Arrow
      Picture         =   "frmMain.frx":755B
      ScaleHeight     =   5535
      ScaleWidth      =   6735
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ListBox lstRank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   3480
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   6135
      End
      Begin VB.Image cmdRefresh 
         Height          =   495
         Left            =   2760
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Image cmdTutup 
         Height          =   495
         Left            =   5040
         Top             =   4560
         Width           =   975
      End
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         LargeChange     =   50
         Left            =   240
         Min             =   1
         TabIndex        =   11
         Top             =   3360
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   10
         Top             =   3960
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAName 
         Caption         =   "Set Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAHeal 
         Caption         =   "Full Heal"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKill 
         Caption         =   "Auto Kill"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Line Line4 
         X1              =   16
         X2              =   168
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   144
         Y2              =   144
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   17280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   9960
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstQuestLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      ItemData        =   "frmMain.frx":33AE3
      Left            =   840
      List            =   "frmMain.frx":33AEA
      MousePointer    =   1  'Arrow
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnuEditors 
      Caption         =   "&Editors"
      Visible         =   0   'False
      Begin VB.Menu mnuEditMap 
         Caption         =   "&Map"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Item"
      End
      Begin VB.Menu mnuEditNpc 
         Caption         =   "&NPC"
      End
      Begin VB.Menu mnuEditResource 
         Caption         =   "&Resource"
      End
      Begin VB.Menu mnuEditShop 
         Caption         =   "&Shop"
      End
      Begin VB.Menu mnuEditSpell 
         Caption         =   "&Spell"
      End
      Begin VB.Menu mnuEditQuest 
         Caption         =   "&Quest"
      End
      Begin VB.Menu mnuEditAnimation 
         Caption         =   "&Animation"
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "&Miscellaneous"
      Visible         =   0   'False
      Begin VB.Menu mnuMapReport 
         Caption         =   "&Map Report"
      End
      Begin VB.Menu mnuMapRespawn 
         Caption         =   "&Map Respawn"
      End
      Begin VB.Menu mnuDeleteBanlist 
         Caption         =   "&Delete Ban List"
      End
      Begin VB.Menu mnuLevelUp 
         Caption         =   "&Level Up"
      End
   End
   Begin VB.Menu mnuClientOnly 
      Caption         =   "&Client Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuGUI 
         Caption         =   "&GUI"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info"
      End
      Begin VB.Menu mnuVisibility 
         Caption         =   "&Visibility"
      End
      Begin VB.Menu mnuNight 
         Caption         =   "&Night"
      End
   End
   Begin VB.Menu mnuOtherCommands 
      Caption         =   "&Other Commands"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long
Private mouseClicked As Boolean

Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub

    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAHeal_Click()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then Exit Sub

    If IsNumeric(Trim$(txtAName.text)) Then Exit Sub

    SendHealPlayer Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAHeal_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAKill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub

    If Len(Trim$(txtAName.text)) < 2 Then Exit Sub

    If IsNumeric(Trim$(txtAName.text)) Then Exit Sub

    SendKillPlayer Trim$(txtAName.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAKill_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAName_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    
    Exit Sub
    End If
    
    If Len(Trim$(txtAName.text)) < 2 Then
    Exit Sub
    End If
    
    If IsNumeric(Trim$(txtAName.text)) Or IsNumeric(Trim$(txtAAccess.text)) Then
    Exit Sub
    End If
    
    SendSetName Trim$(txtAName.text), (Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAName_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAVisible_Click()
    SendVisibility
End Sub

Private Sub CmdHideGUI_Click()
hideGUI = Not hideGUI
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        SendRequestLevelUp GetPlayerName(MyIndex)
    Else
        SendRequestLevelUp Trim$(txtAName.text)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdQuests_Click()
If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditQuest
End Sub

Private Sub cmdSSMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' render the map temp
    ScreenshotMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdIM_Click()
If Picture1(0).Visible = False Then
Picture1(0).Visible = True
Else
Picture1(0).Visible = False
End If
End Sub

Private Sub cmdMall1_Click()
Call ClientOpenShop(19)
End Sub

Private Sub cmdMall2_Click()
Call ClientOpenShop(18)
End Sub

Private Sub cmdMall3_Click()
Call ClientOpenShop(20)
End Sub

Private Sub cmdRank_Click()
picRank.Visible = True
End Sub

Private Sub cmdRefresh_Click()

' If debug mode, handle error then exit out

If Options.Debug = 1 Then On Error GoTo errorhandler



SendRequestRank



' Error handler

Exit Sub

errorhandler:

HandleError "cmdRefresh_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext

Err.Clear

Exit Sub

End Sub

Private Sub cmdTutup_Click()
picRank.Visible = False
End Sub

Private Sub Form_DblClick()
    HandleDoubleClick
End Sub
Private Sub mnuNight_Click()
    NightDisabled = Not NightDisabled
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' move GUI
    picAdmin.Left = 0
    mouseClicked = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseUp Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleMouseMove CLng(X), CLng(Y), Button
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For I = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub menuClient_Click()

End Sub

Private Sub mnuEditMap_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditMap
End Sub

Private Sub mnuEditItem_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditItem
End Sub

Private Sub mnuEditNPC_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditNpc
End Sub

Private Sub mnuEditResource_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditResource
End Sub

Private Sub mnuEditShop_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditShop
End Sub

Private Sub mnuEditSpell_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditSpell
End Sub

Private Sub mnuEditQuest_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditQuest
End Sub

Private Sub mnuEditAnimation_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditAnimation
End Sub

Private Sub mnuGUI_Click()
    hideGUI = Not hideGUI
End Sub

Private Sub mnuInfo_Click()
 BLoc = Not BLoc
End Sub

Private Sub mnuMapReport_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapReport
End Sub

Private Sub mnuDeleteBanList_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendBanDestroy
End Sub

Private Sub mnuLevelUp_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestLevelUp (GetPlayerName(MyIndex))
End Sub


Private Sub mnuMapRespawn_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapRespawn
End Sub

Private Sub mnuOtherCommands_Click()
 picAdmin.Visible = Not picAdmin.Visible
End Sub

Private Sub mnuVisibility_Click()
SendVisibility
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.value).name)
    If Item(scrlAItem.value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.value).Stackable > 0 Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleKeyUp KeyCode
    
    ' Spinning
    If Not GetKeyState(vbKeyUp) < 0 And Not GetKeyState(vbKeyDown) < 0 And Not GetKeyState(vbKeyLeft) < 0 And Not GetKeyState(vbKeyRight) < 0 And Not GetKeyState(vbKeyW) < 0 And Not GetKeyState(vbKeyS) < 0 And Not GetKeyState(vbKeyD) < 0 And Not GetKeyState(vbKeyA) < 0 Then
    
    'Right
    If KeyCode = 33 Then
    Call SpinRight
    End If
    
    'Left
    If KeyCode = 34 Then
    Call SpinLeft
    End If
    
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
    
    Exit Sub
    End If
    
    If Len(Trim$(txtASprite.text)) < 1 Then
    Exit Sub
    End If
    
    If Not IsNumeric(Trim$(txtASprite.text)) Then
    Exit Sub
    End If
    
    If Len(Trim$(txtAName.text)) > 1 Then
    SendSetSprite CLng(Trim$(txtASprite.text)), txtAName.text
    Else
    SendSetSprite CLng(Trim$(txtASprite.text)), GetPlayerName(MyIndex)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.value, scrlAAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picAdmin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseClicked = True
End Sub

Private Sub picAdmin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseClicked = False
End Sub

Private Sub picAdmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mouseClicked Then
        picAdmin.Left = picAdmin.Left + X
        picAdmin.Top = picAdmin.Top + Y
    End If
End Sub
