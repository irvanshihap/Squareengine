VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   583
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   32
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   7935
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   7815
      Begin VB.Frame fraDrop 
         Caption         =   "Drop"
         Height          =   2295
         Left            =   120
         TabIndex        =   63
         Top             =   5520
         Width           =   4815
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   68
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.TextBox txtSpawnSecs 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   67
            Text            =   "0"
            Top             =   960
            Width           =   1815
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   500
            TabIndex        =   66
            Top             =   1920
            Width           =   3495
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            LargeChange     =   50
            Left            =   1200
            Max             =   1000
            TabIndex        =   65
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   64
            Text            =   "0"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spawn Rate (in seconds)"
            Height          =   180
            Left            =   120
            TabIndex        =   73
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   72
            Top             =   1920
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   71
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   70
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance 1 out of"
            Height          =   180
            Left            =   120
            TabIndex        =   69
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   1185
         End
      End
      Begin VB.Frame frmExtra 
         Caption         =   "Stuffs :D"
         Height          =   7940
         Left            =   5160
         TabIndex        =   43
         Top             =   0
         Width           =   2655
         Begin VB.Frame Frame7 
            Caption         =   "Color"
            Height          =   1245
            Left            =   240
            TabIndex        =   77
            Top             =   6600
            Width           =   2175
            Begin VB.HScrollBar scrlB 
               Height          =   255
               Left            =   960
               Max             =   255
               TabIndex        =   81
               Top             =   900
               Width           =   1095
            End
            Begin VB.HScrollBar scrlG 
               Height          =   255
               Left            =   960
               Max             =   255
               TabIndex        =   80
               Top             =   650
               Width           =   1095
            End
            Begin VB.HScrollBar scrlR 
               Height          =   255
               Left            =   960
               Max             =   255
               TabIndex        =   79
               Top             =   400
               Width           =   1095
            End
            Begin VB.HScrollBar scrlA 
               Height          =   255
               Left            =   960
               Max             =   255
               TabIndex        =   78
               Top             =   150
               Width           =   1095
            End
            Begin VB.Label lblB 
               Caption         =   "Blue: 255"
               Height          =   255
               Left            =   80
               TabIndex        =   85
               Top             =   930
               Width           =   1095
            End
            Begin VB.Label lblG 
               Caption         =   "Green: 255"
               Height          =   255
               Left            =   80
               TabIndex        =   84
               Top             =   680
               Width           =   1095
            End
            Begin VB.Label lblR 
               Caption         =   "Red: 255"
               Height          =   255
               Left            =   80
               TabIndex        =   83
               Top             =   430
               Width           =   1095
            End
            Begin VB.Label lblA 
               Caption         =   "Alpha: 255"
               Height          =   255
               Left            =   80
               TabIndex        =   82
               Top             =   180
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Element"
            Height          =   855
            Left            =   240
            TabIndex        =   74
            Top             =   5760
            Width           =   2175
            Begin VB.HScrollBar scrlElement 
               Height          =   255
               Left            =   120
               Max             =   6
               TabIndex        =   75
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label lblElement 
               Alignment       =   2  'Center
               Caption         =   "Element: None"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame fraSpell 
            Caption         =   "Spells"
            Height          =   1455
            Left            =   240
            TabIndex        =   58
            Top             =   4200
            Width           =   2175
            Begin VB.HScrollBar scrlSpell 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   60
               Top             =   1080
               Value           =   1
               Width           =   1935
            End
            Begin VB.HScrollBar scrlSpellNum 
               Height          =   255
               Left            =   120
               Max             =   5
               Min             =   1
               TabIndex        =   59
               Top             =   240
               Value           =   1
               Width           =   1935
            End
            Begin VB.Label lblSpellNum 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Num: 0"
               Height          =   180
               Left            =   120
               TabIndex        =   62
               Top             =   840
               Width           =   1875
            End
            Begin VB.Label lblSpellName 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Spell: None"
               Height          =   180
               Left            =   120
               TabIndex        =   61
               Top             =   600
               Width           =   1935
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Attack Speed"
            Height          =   1335
            Left            =   240
            TabIndex        =   55
            Top             =   2760
            Width           =   2175
            Begin VB.HScrollBar scrlAttackSpeed 
               Height          =   255
               LargeChange     =   100
               Left            =   120
               Max             =   30000
               Min             =   100
               TabIndex        =   57
               Top             =   960
               Value           =   100
               Width           =   1935
            End
            Begin VB.Label lblAttackSpeed 
               Alignment       =   2  'Center
               Caption         =   "Attack Speed:"
               Height          =   495
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Projectile"
            Height          =   1095
            Left            =   240
            TabIndex        =   48
            Top             =   1560
            Width           =   2175
            Begin VB.HScrollBar scrlProjectileRotation 
               Height          =   255
               LargeChange     =   10
               Left            =   1080
               Max             =   100
               TabIndex        =   51
               Top             =   720
               Value           =   1
               Width           =   975
            End
            Begin VB.HScrollBar scrlProjectileRange 
               Height          =   255
               Left            =   1080
               Max             =   255
               TabIndex        =   50
               Top             =   480
               Width           =   975
            End
            Begin VB.HScrollBar scrlProjectilePic 
               Height          =   255
               Left            =   1080
               TabIndex        =   49
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblProjectileRotation 
               Caption         =   "Rotation: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label lblProjectileRange 
               Caption         =   "Range: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblProjectilePic 
               Caption         =   "Pic: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame fraQuest 
            Caption         =   "Quest"
            Height          =   1335
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   2175
            Begin VB.CheckBox chkQuest 
               Caption         =   "Quest giver?"
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   240
               Width           =   1335
            End
            Begin VB.HScrollBar scrlQuest 
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label lblQuest 
               Caption         =   "Quest: None"
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   960
               Width           =   1935
            End
         End
      End
      Begin VB.HScrollBar scrlMoveSpeed 
         Height          =   255
         Left            =   2640
         Max             =   10
         Min             =   1
         TabIndex        =   42
         Top             =   3480
         Value           =   1
         Width           =   2175
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   2880
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   3240
         Width           =   2175
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   21
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1320
         List            =   "frmEditor_NPC.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   4815
         Begin VB.CheckBox chkDay 
            Caption         =   "Night Spawn?"
            Height          =   255
            Left            =   3360
            TabIndex        =   88
            Top             =   920
            Width           =   1335
         End
         Begin VB.CheckBox chkNight 
            Caption         =   "Day Spawn?"
            Height          =   255
            Left            =   3360
            TabIndex        =   87
            Top             =   1160
            Width           =   1335
         End
         Begin VB.CheckBox chkIsBoss 
            Caption         =   "Is Boss?"
            Height          =   255
            Left            =   3360
            TabIndex        =   86
            Top             =   680
            Width           =   975
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   120
            Max             =   30000
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   1680
            Max             =   30000
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   3240
            Max             =   30000
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            LargeChange     =   10
            Left            =   120
            Max             =   30000
            TabIndex        =   8
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            LargeChange     =   10
            Left            =   1680
            Max             =   30000
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   15
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   12
            Top             =   1080
            Width           =   480
         End
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblMoveSpeed 
         Caption         =   "Movement Speed: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2640
         TabIndex        =   37
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2640
         TabIndex        =   24
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7440
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   8040
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SpellIndex As Long
Private DropIndex As Byte


Private Sub chkDay_Click()
    NPC(EditorIndex).SpawnAtDay = chkDay.value
End Sub
Private Sub chkNight_Click()
    NPC(EditorIndex).SpawnAtNight = chkNight.value
End Sub

Private Sub chkIsBoss_Click()
    NPC(EditorIndex).isBoss = chkIsBoss.value
End Sub

Private Sub chkQuest_Click()
NPC(EditorIndex).Quest = chkQuest.value
End Sub

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub scrlA_Change()
    lblA.Caption = "Alpha: " & 255 - scrlA.value
    NPC(EditorIndex).a = scrlA.value
End Sub

Private Sub scrlR_Change()
    lblR.Caption = "Red: " & 255 - scrlR.value
    NPC(EditorIndex).r = scrlR.value
End Sub

Private Sub scrlG_Change()
    lblG.Caption = "Green: " & 255 - scrlG.value
    NPC(EditorIndex).G = scrlG.value
End Sub

Private Sub scrlB_Change()
    lblB.Caption = "Blue: " & 255 - scrlB.value
    NPC(EditorIndex).B = scrlB.value
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).name)
    lblAnimation.Caption = "Anim: " & sString
    NPC(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAttackSpeed_Change()
Dim intSpeed As Integer
Dim dblValue As Double
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

intSpeed = scrlAttackSpeed.value
NPC(EditorIndex).AttackSpeed = intSpeed

If intSpeed >= 100 And intSpeed <= 1000 Then
dblValue = Round(1000 / intSpeed, 3)
lblAttackSpeed.Caption = "Attack speed: " & dblValue & " attack(s) per 1 second."
ElseIf intSpeed > 1000 Then
dblValue = intSpeed / 1000
lblAttackSpeed.Caption = "Attack speed: 1 attack per " & dblValue & " second(s)."
Else
' lblAttackSpeed.Caption = "Attack speed: " & intSpeed
End If

' Error handler
Exit Sub
errorhandler:
HandleError "scrlAttackSpeed_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub scrlDrop_Change()
  DropIndex = scrlDrop.value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.text = NPC(EditorIndex).DropChance(DropIndex)
    scrlNum.value = NPC(EditorIndex).DropItem(DropIndex)
    scrlValue.value = NPC(EditorIndex).DropItemValue(DropIndex)
End Sub

Private Sub scrlElement_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub

If scrlElement.value = 0 Then
lblElement.Caption = "Element: None"
ElseIf scrlElement.value = 1 Then
lblElement.Caption = "Element: Fire"
ElseIf scrlElement.value = 2 Then
lblElement.Caption = "Element: Water"
ElseIf scrlElement.value = 3 Then
lblElement.Caption = "Element: Wind"
ElseIf scrlElement.value = 4 Then
lblElement.Caption = "Element: Earth"
ElseIf scrlElement.value = 5 Then
lblElement.Caption = "Element: Light"
ElseIf scrlElement.value = 6 Then
lblElement.Caption = "Element: Dark"
End If

NPC(EditorIndex).Element = scrlElement.value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlElement_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.value
    NPC(EditorIndex).Projectile = scrlProjectilePic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.value
    NPC(EditorIndex).ProjectileRange = scrlProjectileRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRotation_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
    lblProjectileRotation.Caption = "Rotation: " & scrlProjectileRotation.value / 2
    NPC(EditorIndex).Rotation = scrlProjectileRotation.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRotation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlQuest_Change()
    If scrlQuest.value > 0 Then
        lblQuest.Caption = "Quest: " & Quest(scrlQuest.value).name
    Else
        lblQuest.Caption = "Quest: None"
    End If
    NPC(EditorIndex).QuestNum = scrlQuest.value
End Sub

Private Sub scrlSpell_Change()
lblSpellNum.Caption = "Num: " & scrlSpell.value
    If scrlSpell.value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpell.value).name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    NPC(EditorIndex).Spell(SpellIndex) = scrlSpell.value
End Sub
Private Sub scrlSpellNum_Change()
SpellIndex = scrlSpellNum.value
    fraSpell.Caption = "Spell - " & SpellIndex
    scrlSpell.value = NPC(EditorIndex).Spell(SpellIndex)
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    NPC(EditorIndex).Sprite = scrlSprite.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.value
    NPC(EditorIndex).Range = scrlRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "Num: " & scrlNum.value

    If scrlNum.value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.value).name)
    Else
        lblItemName.Caption = "Item: None "
    End If
    
    NPC(EditorIndex).DropItem(scrlDrop.value) = scrlNum.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).value
    NPC(EditorIndex).Stat(Index) = scrlStat(Index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.value
    NPC(EditorIndex).DropItemValue(scrlDrop.value) = scrlValue.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).AttackSay = txtAttackSay.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Change()
    On Error GoTo chanceErr
    
    If Not IsNumeric(txtChance.text) And Not Right$(txtChance.text, 1) = "%" And Not InStr(1, txtChance.text, "/") > 0 And Not InStr(1, txtChance.text, ".") Then
        txtChance.text = "0"
        NPC(EditorIndex).DropChance(scrlDrop.value) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.text, 1) = "%" Then
        txtChance.text = Left(txtChance.text, Len(txtChance.text) - 1) / 100
    ElseIf InStr(1, txtChance.text, "/") > 0 Then
        Dim I() As String
        I = Split(txtChance.text, "/")
        txtChance.text = Int(I(0) / I(1) * 1000) / 1000
    End If
    
    If txtChance.text > 1 Or txtChance.text < 0 Then
        Err.Description = "Value must be between 0 and 1!"
        GoTo chanceErr
    End If
    
    NPC(EditorIndex).DropChance(scrlDrop.value) = txtChance.text
    Exit Sub
    
chanceErr:
    MsgBox "Invalid entry for chance! " & Err.Description
    txtChance.text = "0"
    NPC(EditorIndex).DropChance(scrlDrop.value) = 0
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtDamage.text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.text) Then NPC(EditorIndex).Damage = Val(txtDamage.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.text) Then NPC(EditorIndex).EXP = Val(txtEXP.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtHP.text) > 0 Then Exit Sub
    If IsNumeric(txtHP.text) Then NPC(EditorIndex).HP = Val(txtHP.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtLevel.text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.text) Then NPC(EditorIndex).Level = Val(txtLevel.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.text) > 0 Then Exit Sub
    NPC(EditorIndex).SpawnSecs = Val(txtSpawnSecs.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        NPC(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        NPC(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMoveSpeed_Change()
    lblMoveSpeed.Caption = "Movement Speed: " & scrlMoveSpeed.value
    NPC(EditorIndex).speed = scrlMoveSpeed.value
End Sub
