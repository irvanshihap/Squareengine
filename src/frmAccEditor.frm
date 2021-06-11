VERSION 5.00
Begin VB.Form frmAccEditor 
   Caption         =   "Account Editor"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Account Info"
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cmbAccess 
         Height          =   315
         ItemData        =   "frmAccEditor.frx":0000
         Left            =   960
         List            =   "frmAccEditor.frx":0013
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtCount 
         Height          =   285
         Left            =   3480
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkMember 
         Caption         =   "Member Access?"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save account"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   2415
      End
      Begin VB.HScrollBar scrlSwitch 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Value           =   1
         Width           =   2415
      End
      Begin VB.CheckBox chkSwitch 
         Caption         =   "Switch: 0 On/Off"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1650
         Width           =   2535
      End
      Begin VB.HScrollBar scrlVariable 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtVariable 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   2205
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3240
         Width           =   2415
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   3480
         Min             =   1
         TabIndex        =   1
         Top             =   960
         Value           =   1
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Login:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Access:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Day #:"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblVariable 
         Caption         =   "Variable: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2235
         Width           =   1095
      End
      Begin VB.Label lblLevel 
         Caption         =   "Level: 1"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAccEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MemberChanged As Boolean

Private Sub cmdClose_Click()
    AEditorPlayer = vbNullString
    MemberChanged = False
    Unload frmAccEditor
End Sub

Private Sub Form_Load()
On Error GoTo errorhandler
    txtLogin.Text = Trim$(AEditor.Login)
    txtName.Text = Trim$(AEditor.Name)
    txtPassword.Text = Trim$(AEditor.Password)
    scrlSwitch.Value = 0
    scrlVariable.Value = 0
    cmbAccess.ListIndex = AEditor.Access
    chkMember.Value = AEditor.IsMember
    If AEditor.IsMember = 0 Then
        txtCount.Text = Format(Date, "m/d/yyyy")
    Else
        txtCount.Text = Trim$(AEditor.DateCount)
    End If
    scrlSwitch.Max = MAX_SWITCHES
    scrlVariable.Max = MAX_VARIABLES
    scrlLevel.Max = MAX_LEVELS
    chkSwitch.Value = AEditor.Switches(0)
    txtVariable.Text = Trim$(AEditor.Variables(0))
    scrlLevel.Value = AEditor.Level
    
    MemberChanged = False
    ' Error handler
    Exit Sub
errorhandler:
    Err.Clear
    Exit Sub
End Sub
Private Sub cmdSave_Click()
Dim filename As String
Dim f As Long
Dim i As Long
Dim index As Long
    If Len(Trim$(AEditorPlayer)) > 0 Then
        AEditor.Name = Trim$(txtName.Text)
        AEditor.Login = Trim$(txtLogin.Text)
        AEditor.Password = Trim$(txtPassword.Text)
        AEditor.Access = cmbAccess.ListIndex
        If chkMember.Value = 1 Then
            AEditor.IsMember = 1
        Else
            AEditor.DateCount = "11/11/2011"
        End If
        AEditor.DateCount = Trim$(txtCount.Text)
        AEditor.IsMember = chkMember.Value
        AEditor.Level = scrlLevel.Value
        index = FindPlayer(Trim$(AEditor.Name))
        If index > 0 And index <= MAX_PLAYERS Then
            If IsPlaying(index) Then
                Player(index).Name = Trim$(txtName.Text)
                Player(index).Login = Trim$(txtLogin.Text)
                Player(index).Password = Trim$(txtPassword.Text)
                Player(index).Access = Trim$(cmbAccess.ListIndex)
                Player(index).DateCount = Trim$(txtCount.Text)
                If MemberChanged = True Then
                    If chkMember.Value = 1 Then
                        Player(index).IsMember = 1
                        Player(index).DateCount = Trim$(txtCount.Text)
                        PlayerMsg index, "You have been granted membership by the server.", Yellow
                    Else
                        PlayerMsg index, "Your membership has been expired.", BrightRed
                        Player(index).IsMember = 0
                        Player(index).DateCount = "11/11/2011"
                        MemberUnEquipItem index
                        If Map(GetPlayerMap(index)).IsMember > 0 Then
                            PlayerWarp index, Map(GetPlayerMap(index)).BootMap, Map(GetPlayerMap(index)).BootX, Map(GetPlayerMap(index)).BootY
                        End If
                    End If
                End If
                For i = 0 To MAX_SWITCHES
                    Player(index).Switches(i) = AEditor.Switches(i)
                Next
                For i = 0 To MAX_VARIABLES
                    Player(index).Variables(i) = AEditor.Variables(i)
                Next
                For i = 1 To Stats.Stat_Count - 1
                    Player(index).stat(i) = AEditor.stat(i)
                Next
                Player(index).Level = AEditor.Level
                For i = 1 To MAX_PLAYER_SPELLS
                    Player(index).SpellUses(i) = AEditor.SpellUses(i)
                Next
                SavePlayer index
                SendPlayerData index
                SendSwitchesAndVariables index
            End If
        Else
            filename = App.path & "\data\accounts\" & Trim$(AEditorPlayer) & ".bin"
            f = FreeFile
            Open filename For Binary As #f
            Put #f, , AEditor
            Close #f
        End If
    End If
    AEditorPlayer = vbNullString
    MemberChanged = False
    Unload frmAccEditor
End Sub

Private Sub chkMember_Click()
    MemberChanged = True
End Sub



Private Sub scrlLevel_Change()
    lblLevel.Caption = "Level: " & scrlLevel.Value
End Sub

Private Sub scrlSwitch_Change()
    chkSwitch.Caption = "Switch: " & scrlSwitch.Value & " On/Off"
    chkSwitch.Value = AEditor.Switches(scrlSwitch.Value)
End Sub

Private Sub chkSwitch_Click()
    AEditor.Switches(scrlSwitch.Value) = chkSwitch.Value
End Sub

Private Sub scrlVariable_Change()
    lblVariable.Caption = "Variable: " & scrlVariable.Value
    txtVariable.Text = Trim$(AEditor.Variables(scrlVariable.Value))
End Sub

Private Sub txtCount_Change()
    MemberChanged = True
End Sub

Private Sub txtVariable_Change()
    AEditor.Variables(scrlVariable.Value) = Val(txtVariable.Text)
End Sub

