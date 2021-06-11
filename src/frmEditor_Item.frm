VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FraSprite 
      Caption         =   "Frame7"
      Height          =   1815
      Left            =   9840
      TabIndex        =   122
      Top             =   4680
      Width           =   1575
      Begin VB.HScrollBar scrlEvolve 
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblEvolve 
         Caption         =   "Mount: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   8535
      Begin VB.HScrollBar scrlToolpower 
         Height          =   255
         Left            =   4440
         Max             =   500
         TabIndex        =   103
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox chkClassReq 
         Caption         =   "Not allowed to use?"
         Height          =   255
         Left            =   6240
         TabIndex        =   102
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Frame Frame6 
         Caption         =   "Carry Weight"
         Height          =   975
         Left            =   6240
         TabIndex        =   99
         Top             =   1320
         Width           =   2175
         Begin VB.HScrollBar scrlWeight 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   101
            Top             =   600
            Width           =   1940
         End
         Begin VB.Label lblWeight 
            Alignment       =   2  'Center
            Caption         =   "Weight: 0"
            Height          =   255
            Left            =   480
            TabIndex        =   100
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.HScrollBar scrlG 
         Height          =   255
         Left            =   7200
         Max             =   255
         TabIndex        =   94
         Top             =   720
         Width           =   1215
      End
      Begin VB.HScrollBar scrlB 
         Height          =   255
         Left            =   7200
         Max             =   255
         TabIndex        =   93
         Top             =   960
         Width           =   1215
      End
      Begin VB.HScrollBar scrlR 
         Height          =   255
         Left            =   7200
         Max             =   255
         TabIndex        =   92
         Top             =   480
         Width           =   1215
      End
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   7200
         Max             =   255
         TabIndex        =   91
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable"
         Height          =   255
         Left            =   2880
         TabIndex        =   76
         Top             =   2640
         Width           =   1335
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   74
         Top             =   2400
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   72
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   4200
         List            =   "frmEditor_Item.frx":333F
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3368
         Left            =   120
         List            =   "frmEditor_Item.frx":339F
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblToolpwr 
         Alignment       =   2  'Center
         Caption         =   "Tool Power: 0"
         Height          =   255
         Left            =   4440
         TabIndex        =   104
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblR 
         Caption         =   "Red: 255"
         Height          =   255
         Left            =   6240
         TabIndex        =   98
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblG 
         Caption         =   "Green: 255"
         Height          =   255
         Left            =   6240
         TabIndex        =   97
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblB 
         Caption         =   "Blue: 255"
         Height          =   255
         Left            =   6240
         TabIndex        =   96
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblA 
         Caption         =   "Alpha: 255"
         Height          =   255
         Left            =   6240
         TabIndex        =   95
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   75
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   73
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Requirement:"
         Height          =   180
         Left            =   6600
         TabIndex        =   71
         Top             =   2280
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stat Requirements"
      Height          =   975
      Left            =   3480
      TabIndex        =   6
      Top             =   6840
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   30000
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3120
         Max             =   30000
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   30000
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   30000
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3120
         Max             =   30000
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   4215
      Left            =   3360
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CheckBox ChkStaff 
         Caption         =   "Mage Projectile?"
         Height          =   255
         Left            =   6600
         TabIndex        =   117
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbCTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":342A
         Left            =   4560
         List            =   "frmEditor_Item.frx":3440
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Element"
         Height          =   855
         Left            =   3480
         TabIndex        =   88
         Top             =   2210
         Width           =   2655
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Left            =   120
            Max             =   6
            TabIndex        =   89
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblElement 
            Alignment       =   2  'Center
            Caption         =   "Element: None"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CheckBox ChkTwoh 
         Caption         =   "Two Handed?"
         Height          =   180
         Left            =   6600
         TabIndex        =   77
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5640
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   1080
         Width           =   480
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   3480
         TabIndex        =   57
         Top             =   1950
         Width           =   2655
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   40
         Top             =   720
         Value           =   100
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   30000
         TabIndex        =   39
         Top             =   1440
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   30000
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   4680
         Max             =   30000
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   30000
         TabIndex        =   36
         Top             =   1080
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   35
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3480
         Left            =   1320
         List            =   "frmEditor_Item.frx":3490
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   30000
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Caption         =   "Projectile"
         Height          =   1335
         Left            =   120
         TabIndex        =   78
         Top             =   1760
         Width           =   3255
         Begin VB.HScrollBar scrlProjectilePic 
            Height          =   255
            Left            =   1440
            TabIndex        =   83
            Top             =   240
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectileRange 
            Height          =   255
            Left            =   1440
            Max             =   255
            TabIndex        =   82
            Top             =   480
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectileRotation 
            Height          =   255
            LargeChange     =   10
            Left            =   1440
            Max             =   100
            TabIndex        =   81
            Top             =   720
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectileAmmo 
            Height          =   255
            Left            =   1440
            TabIndex        =   80
            Top             =   960
            Width           =   1095
         End
         Begin VB.PictureBox picProjectile 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   79
            Top             =   480
            Width           =   480
         End
         Begin VB.Label lblProjectilePic 
            Caption         =   "Pic: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblProjectileRange 
            Caption         =   "Range: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblProjectileRotation 
            Caption         =   "Rotation: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblProjectileAmmo 
            Caption         =   "Ammo: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.Label lblCTool 
         Caption         =   "Crafting Tool:"
         Height          =   255
         Left            =   3240
         TabIndex        =   116
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   4440
         TabIndex        =   56
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   48
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2040
         TabIndex        =   47
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   3840
         TabIndex        =   45
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2040
         TabIndex        =   44
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Dam/Def:"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlMultiplier 
         Height          =   255
         Left            =   3960
         Max             =   255
         TabIndex        =   119
         Top             =   1200
         Width           =   2055
      End
      Begin VB.HScrollBar scrlMultiplierTime 
         Height          =   255
         Left            =   3960
         TabIndex        =   118
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   65
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   63
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   61
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   50
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblExpMultiplier 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp Multiplier: 0"
         Height          =   255
         Left            =   3960
         TabIndex        =   121
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblExpTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplier Time: 0"
         Height          =   255
         Left            =   3960
         TabIndex        =   120
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   62
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   53
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraRecipe 
      Caption         =   "Recipe"
      Height          =   1365
      Left            =   3360
      TabIndex        =   105
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlResult 
         Height          =   255
         LargeChange     =   50
         Left            =   3000
         Max             =   1000
         TabIndex        =   109
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cmbCToolReq 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":34B1
         Left            =   3000
         List            =   "frmEditor_Item.frx":34C7
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrlItem1 
         Height          =   255
         LargeChange     =   50
         Left            =   120
         Max             =   1000
         TabIndex        =   107
         Top             =   360
         Width           =   2175
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   106
         Top             =   960
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "1."
         Height          =   255
         Left            =   720
         TabIndex        =   114
         Top             =   120
         Width           =   135
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         Caption         =   "Result: None"
         Height          =   180
         Left            =   3000
         TabIndex        =   113
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tool:"
         Height          =   180
         Left            =   2520
         TabIndex        =   112
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblItem1 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   960
         TabIndex        =   111
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblItemNum 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   840
         TabIndex        =   110
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecipeIndex As Long
Private LastIndex As Long

Private Sub chkClassReq_Click()
Item(EditorIndex).ClassReq(cmbClassReq.ListIndex + 1) = chkClassReq.Value
End Sub

Private Sub ChkStaff_Click()

    If ChkStaff.Value = 0 Then
        Item(EditorIndex).MageProjectile = False
    Else
        Item(EditorIndex).MageProjectile = True
    End If
    
End Sub

Private Sub ChkTwoh_Click()
 ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ChkTwoh.Value = 0 Then
        Item(EditorIndex).istwohander = False
    Else
        Item(EditorIndex).istwohander = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    chkClassReq.Value = Item(EditorIndex).ClassReq(cmbClassReq.ListIndex + 1)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.max = numitems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        chkStackable.Visible = False
        Item(EditorIndex).Stackable = 0
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
        chkStackable.Visible = True
    End If
    
     If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
        ChkTwoh.Visible = True
        Frame4.Visible = True
        Frame5.Visible = True
        scrlToolpower.Visible = True
        ChkStaff.Visible = True
            Else
        ChkTwoh.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
        scrlToolpower.Visible = False
        ChkStaff.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_TRANSFORM Then

FraSprite.Visible = True

'scrlVitalMod_Change

Else

FraSprite.Visible = False

End If
    
    If (cmbType.ListIndex = ITEM_TYPE_RECIPE) Then
        fraRecipe.Visible = True
    Else
        fraRecipe.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkStackable_Click()
    Item(EditorIndex).Stackable = chkStackable.Value
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub scrlA_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblA.Caption = "Alpha: " & 255 - scrlA.Value
    Item(EditorIndex).a = scrlA.Value
End Sub

Private Sub scrlMultiplier_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    lblExpMultiplier.Caption = "Exp Multiplier: " & scrlMultiplier.Value
    Item(EditorIndex).addExpMultiplier = scrlMultiplier.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMultiplier_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlMultiplierTime_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    lblExpTime.Caption = "Multiplier Time: " & scrlMultiplierTime.Value
    Item(EditorIndex).addExpMultiplierTime = scrlMultiplierTime.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMultiplierTime_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlR_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblR.Caption = "Red: " & 255 - scrlR.Value
    Item(EditorIndex).R = scrlR.Value
End Sub

Private Sub scrlG_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblG.Caption = "Green: " & 255 - scrlG.Value
    Item(EditorIndex).G = scrlG.Value
End Sub

Private Sub scrlB_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblB.Caption = "Blue: " & 255 - scrlB.Value
    Item(EditorIndex).B = scrlB.Value
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Else
    lblDamage.Caption = "Defense: " & scrlDamage.Value
    End If
    
    
    Item(EditorIndex).Data2 = scrlDamage.Value
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlElement_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

If scrlElement.Value = 0 Then
lblElement.Caption = "Element: None"
ElseIf scrlElement.Value = 1 Then
lblElement.Caption = "Element: Fire"
ElseIf scrlElement.Value = 2 Then
lblElement.Caption = "Element: Water"
ElseIf scrlElement.Value = 3 Then
lblElement.Caption = "Element: Wind"
ElseIf scrlElement.Value = 4 Then
lblElement.Caption = "Element: Earth"
ElseIf scrlElement.Value = 5 Then
lblElement.Caption = "Element: Light"
ElseIf scrlElement.Value = 6 Then
lblElement.Caption = "Element: Dark"
End If

Item(EditorIndex).Element = scrlElement.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlElement_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileAmmo_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileAmmo.Caption = "Ammo: " & scrlProjectileAmmo.Value
    Item(EditorIndex).Ammo = scrlProjectileAmmo.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileAmmo_Change", "frmEditor_Item", Err.Ammober, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.Value
    Item(EditorIndex).Projectile = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Item(EditorIndex).Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRotation_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRotation.Caption = "Rotation: " & scrlProjectileRotation.Value / 2
    Item(EditorIndex).Rotation = scrlProjectileRotation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRotation_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.Value).name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.Value
    
    Item(EditorIndex).data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlToolpower_Change()
If scrlToolpower.Value > 0 Then
        lblToolpwr.Caption = "Tool Power: " & scrlToolpower.Value
    Else
        lblToolpwr.Caption = "Tool Power: 0"
    End If
    
    Item(EditorIndex).Toolpower = scrlToolpower.Value
End Sub

Private Sub scrlWeight_Change()
lblWeight.Caption = "Weight: " & scrlWeight.Value
Item(EditorIndex).CarryWeight = scrlWeight.Value

End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' crafting
Private Sub scrlItem1_Change()

If scrlItem1.Value > 0 Then

lblItem1.Caption = "Item: " & Trim$(Item(scrlItem1.Value).name)

Else

lblItem1.Caption = "Item: None"

End If

If RecipeIndex = 0 Then
RecipeIndex = 1
End If

Item(EditorIndex).Recipe(RecipeIndex) = scrlItem1.Value

End Sub

Private Sub scrlItemNum_Change()

RecipeIndex = scrlItemNum.Value

lblItemNum.Caption = "Item: " & RecipeIndex

scrlItem1.Value = Item(EditorIndex).Recipe(RecipeIndex)

End Sub

Private Sub scrlResult_Change()

If scrlResult.Value > 0 Then

lblResult.Caption = "Result: " & Trim$(Item(scrlResult.Value).name)

Else

lblResult.Caption = "Result: None"

End If



Item(EditorIndex).Data3 = scrlResult.Value

End Sub


Private Sub cmbCTool_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Tool = cmbCTool.ListIndex
End Sub

Private Sub cmbCToolReq_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ToolReq = cmbCToolReq.ListIndex
End Sub
'/crafting
Private Sub scrlEvolve_Change()

lblEvolve.Caption = "Mount: " & scrlEvolve.Value

Item(EditorIndex).Sprite = scrlEvolve.Value

End Sub


