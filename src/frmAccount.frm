VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmAccount 
   Caption         =   "Account Manager"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Account Manager"
      TabPicture(0)   =   "frmAccount.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAcctCount"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstAccounts"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAccEditor"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmAccount.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdAccEditor 
         Caption         =   "Edit Selected Account"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   3000
         Width           =   3255
      End
      Begin VB.ListBox lstAccounts 
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   7095
      End
      Begin VB.Label lblAcctCount 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Active accounts:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccEditor_Click()
Dim filename As String
Dim f As Long
    If Len(Trim$(lstAccounts.Text)) > 0 Then
        filename = App.path & "\data\accounts\" & Trim$(lstAccounts.Text) & ".bin"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , AEditor
        Close #f
        
        AEditorPlayer = Trim$(lstAccounts.Text)
        Load frmAccEditor
        frmAccEditor.Show
    End If
End Sub

Private Sub cmdClose_Click()
Unload frmAccount
End Sub
