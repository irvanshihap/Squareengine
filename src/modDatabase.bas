Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim f As Long

    If ServerLog Then
        filename = App.path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If

        f = FreeFile
        Open filename For Append As #f
        Print #f, Time & ": " & Text
        Close #f
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.path & "\data\options.ini", "GUILDS", "Buy_Cost", STR(Options.Buy_Cost)
    PutVar App.path & "\data\options.ini", "GUILDS", "Buy_Lvl", STR(Options.Buy_Lvl)
    PutVar App.path & "\data\options.ini", "GUILDS", "Buy_Item", STR(Options.Buy_Item)
    PutVar App.path & "\data\options.ini", "GUILDS", "Join_Cost", STR(Options.Join_Cost)
    PutVar App.path & "\data\options.ini", "GUILDS", "Join_Lvl", STR(Options.Join_Lvl)
    PutVar App.path & "\data\options.ini", "GUILDS", "Join_Item", STR(Options.Join_Item)
    PutVar App.path & "\data\options.ini", "OPTIONS", "DayNight", STR(Options.DayNight)
    PutVar App.path & "\data\options.ini", "OPTIONS", "PathfindingType", STR(Options.PathfindingType)
    
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.path & "\data\options.ini", "OPTIONS", "Website")
    Options.Buy_Cost = GetVar(App.path & "\data\options.ini", "GUILDS", "Buy_Cost")
    Options.Buy_Lvl = GetVar(App.path & "\data\options.ini", "GUILDS", "Buy_Lvl")
    Options.Buy_Item = GetVar(App.path & "\data\options.ini", "GUILDS", "Buy_Item")
    Options.Join_Cost = GetVar(App.path & "\data\options.ini", "GUILDS", "Join_Cost")
    Options.Join_Lvl = GetVar(App.path & "\data\options.ini", "GUILDS", "Join_Lvl")
    Options.Join_Item = GetVar(App.path & "\data\options.ini", "GUILDS", "Join_Item")
    Options.DayNight = GetVar(App.path & "\data\options.ini", "OPTIONS", "DayNight")
    Options.PathfindingType = GetVar(App.path & "\data\options.ini", "OPTIONS", "PathfindingType")
    
    frmServer.txtGBuyCost.Text = Options.Buy_Cost
    frmServer.txtGBuyLvl.Text = Options.Buy_Lvl
    frmServer.txtGBuyItem.Text = Options.Buy_Item
    frmServer.txtGJoinCost.Text = Options.Join_Cost
    frmServer.txtGJoinLvl.Text = Options.Join_Lvl
    frmServer.txtGJoinItem.Text = Options.Join_Item
    
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim f As Long
    Dim i As Long
    filename = App.path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    f = FreeFile
    Open filename For Append As #f
    Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim f As Long
    Dim i As Long
    filename = App.path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    f = FreeFile
    Open filename For Append As #f
    Print #f, IP & "," & "Server"
    Close #f
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer index
    
    Player(index).Login = Name
    Player(index).Password = Password

    Call SavePlayer(index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.path & "\data\accounts\charlist.txt", App.path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean

    If LenB(Trim$(Player(index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim f As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(index).Name = Name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum
        
        
        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).stat(n) = Class(ClassNum).stat(n)
        Next n

        Player(index).Dir = DIR_DOWN
        Player(index).Map = START_MAP
        Player(index).x = START_X
        Player(index).y = START_Y
        Player(index).Dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Inv(n).num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(n)).Name)) > 0 Then
                        Player(index).Spell(n) = Class(ClassNum).StartSpell(n)
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        f = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal index As Long)
    Dim filename As String
    Dim f As Long

    filename = App.path & "\data\accounts\" & Trim$(Player(index).Login) & ".bin"
    
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , Player(index)
    Close #f
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
    Dim filename As String
    Dim f As Long
    Call ClearPlayer(index)
    filename = App.path & "\data\accounts\" & Trim(Name) & ".bin"
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Player(index)
    Close #f
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    ' whoever added this apparently does not understand what ZeroMemory does.
    ' love, Carim.
    Player(index).Login = vbNullString
    Player(index).Password = vbNullString
    Player(index).Name = vbNullString
    Player(index).Class = 1

    frmServer.lbPlayers.List(index - 1) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.path & "\data\classes.ini"
    Max_Classes = 16

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        Class(i).Next = Val(GetVar(filename, "CLASS" & i, "DoNotAdvance"))
        Class(i).Advancement(1) = Val(GetVar(filename, "CLASS" & i, "Advancement1"))
        Class(i).Advancement(2) = Val(GetVar(filename, "CLASS" & i, "Advancement2"))
        Class(i).Advancement(3) = Val(GetVar(filename, "CLASS" & i, "Advancement3"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    
    filename = App.path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal itemnum As Long)
    Dim filename As String
    Dim f  As Long
    filename = App.path & "\data\items\item" & itemnum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Item(itemnum)
    Close #f
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Item(i)
        Close #f
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data\shops\shop" & shopNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Shop(shopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Shop(i)
        Close #f
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data\spells\spells" & SpellNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NPCNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data\npcs\npc" & NPCNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , NPC(NPCNum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.path & "\data\npcs\npc" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , NPC(i)
        Close #f
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(index)), LenB(NPC(index)))
    NPC(index).Name = vbNullString
    NPC(index).AttackSay = vbNullString
    NPC(index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data\resources\resource" & ResourceNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Resource(ResourceNum)
    Close #f
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.path & "\data\resources\resource" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Resource(i)
        Close #f
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data\animations\animation" & AnimationNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Animation(AnimationNum)
    Close #f
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.path & "\data\animations\animation" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Animation(i)
        Close #f
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapnum As Long)
    Dim filename As String
    Dim f As Long
    Dim x As Long
    Dim y As Long, i As Long, z As Long, w As Long
    filename = App.path & "\data\maps\map" & mapnum & ".dat"
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , Map(mapnum).Name
    Put #f, , Map(mapnum).Music
    Put #f, , Map(mapnum).BGS
    Put #f, , Map(mapnum).Revision
    Put #f, , Map(mapnum).Moral
    Put #f, , Map(mapnum).Up
    Put #f, , Map(mapnum).Down
    Put #f, , Map(mapnum).Left
    Put #f, , Map(mapnum).Right
    Put #f, , Map(mapnum).BootMap
    Put #f, , Map(mapnum).BootX
    Put #f, , Map(mapnum).BootY
    
    Put #f, , Map(mapnum).Weather
    Put #f, , Map(mapnum).WeatherIntensity
    
    Put #f, , Map(mapnum).Fog
    Put #f, , Map(mapnum).FogSpeed
    Put #f, , Map(mapnum).FogOpacity
    
    Put #f, , Map(mapnum).Red
    Put #f, , Map(mapnum).Green
    Put #f, , Map(mapnum).Blue
    Put #f, , Map(mapnum).Alpha
    
    Put #f, , Map(mapnum).MaxX
    Put #f, , Map(mapnum).MaxY
    
    Call CacheMapBlocks(mapnum)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            Put #f, , Map(mapnum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map(mapnum).NPC(x)
        Put #f, , Map(mapnum).NpcSpawnType(x)
    Next
    
    Put #f, , Map(mapnum).IsMember
    Put #f, , Map(mapnum).DayNight
    Put #f, , Map(mapnum).Panorama
    
    Close #f
    
    'This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.
    filename = App.path & "\data\maps\map" & mapnum & "_eventdata.dat"
    PutVar filename, "Events", "EventCount", Val(Map(mapnum).EventCount)
    
    If Map(mapnum).EventCount > 0 Then
        For i = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "Global", Val(.Global)
                PutVar filename, "Event" & i, "x", Val(.x)
                PutVar filename, "Event" & i, "y", Val(.y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            If Map(mapnum).Events(i).PageCount > 0 Then
                For x = 1 To Map(mapnum).Events(i).PageCount
                    With Map(mapnum).Events(i).Pages(x)
                        PutVar filename, "Event" & i & "Page" & x, "chkVariable", Val(.chkVariable)
                        PutVar filename, "Event" & i & "Page" & x, "VariableIndex", Val(.VariableIndex)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCondition", Val(.VariableCondition)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCompare", Val(.VariableCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSwitch", Val(.chkSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchIndex", Val(.SwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchCompare", Val(.SwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemIndex", Val(.HasItemIndex)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchIndex", Val(.SelfSwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchCompare", Val(.SelfSwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & x, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX2", Val(.GraphicX2)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY2", Val(.GraphicY2)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & x, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & x, "MoveFreq", Val(.MoveFreq)
                        
                        PutVar filename, "Event" & i & "Page" & x, "IgnoreMoveRoute", Val(.IgnoreMoveRoute)
                        PutVar filename, "Event" & i & "Page" & x, "RepeatMoveRoute", Val(.RepeatMoveRoute)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveRouteCount", Val(.MoveRouteCount)
                        
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Index", Val(.MoveRoute(y).index)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data1", Val(.MoveRoute(y).Data1)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data2", Val(.MoveRoute(y).Data2)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data3", Val(.MoveRoute(y).Data3)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data4", Val(.MoveRoute(y).Data4)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data5", Val(.MoveRoute(y).data5)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data6", Val(.MoveRoute(y).data6)
                            Next
                        End If
                        
                        PutVar filename, "Event" & i & "Page" & x, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & x, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & x, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & x, "ShowName", Val(.ShowName)
                        PutVar filename, "Event" & i & "Page" & x, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & x, "CommandListCount", Val(.CommandListCount)
                        
                        PutVar filename, "Event" & i & "Page" & x, "Position", Val(.Position)
                    End With
                    
                    If Map(mapnum).Events(i).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map(mapnum).Events(i).Pages(x).CommandListCount
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "CommandCount", Val(Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount)
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "ParentList", Val(Map(mapnum).Events(i).Pages(x).CommandList(y).ParentList)
                            If Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                For z = 1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Index", Val(.index)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text1", .Text1
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text2", .Text2
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text3", .Text3
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text4", .Text4
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text5", .Text5
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data1", Val(.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data2", Val(.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data3", Val(.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data4", Val(.Data4)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data5", Val(.data5)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data6", Val(.data6)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchCommandList", Val(.ConditionalBranch.CommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchCondition", Val(.ConditionalBranch.Condition)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData1", Val(.ConditionalBranch.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData2", Val(.ConditionalBranch.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData3", Val(.ConditionalBranch.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchElseCommandList", Val(.ConditionalBranch.ElseCommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRouteCount", Val(.MoveRouteCount)
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Index", Val(.MoveRoute(w).index)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data1", Val(.MoveRoute(w).Data1)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data2", Val(.MoveRoute(w).Data2)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data3", Val(.MoveRoute(w).Data3)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data4", Val(.MoveRoute(w).Data4)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data5", Val(.MoveRoute(w).data5)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data6", Val(.MoveRoute(w).data6)
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim x As Long
    Dim y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Map(i).Name
        Get #f, , Map(i).Music
        Get #f, , Map(i).BGS
        Get #f, , Map(i).Revision
        Get #f, , Map(i).Moral
        Get #f, , Map(i).Up
        Get #f, , Map(i).Down
        Get #f, , Map(i).Left
        Get #f, , Map(i).Right
        Get #f, , Map(i).BootMap
        Get #f, , Map(i).BootX
        Get #f, , Map(i).BootY
        
        Get #f, , Map(i).Weather
        Get #f, , Map(i).WeatherIntensity
        
        Get #f, , Map(i).Fog
        Get #f, , Map(i).FogSpeed
        Get #f, , Map(i).FogOpacity
        
        Get #f, , Map(i).Red
        Get #f, , Map(i).Green
        Get #f, , Map(i).Blue
        Get #f, , Map(i).Alpha
        
        Get #f, , Map(i).MaxX
        Get #f, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #f, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #f, , Map(i).NPC(x)
            Get #f, , Map(i).NpcSpawnType(x)
            MapNpc(i).NPC(x).num = Map(i).NPC(x)
        Next
        
        Get #f, , Map(i).IsMember
        Get #f, , Map(i).DayNight
        Get #f, , Map(i).Panorama

        Close #f
        
        ClearTempTile i
        CacheResources i
        DoEvents
        CacheMapBlocks i
    Next
    
    For z = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & z & "_eventdata.dat"
        Map(z).EventCount = Val(GetVar(filename, "Events", "EventCount"))
        
        If Map(z).EventCount > 0 Then
            ReDim Map(z).Events(0 To Map(z).EventCount)
            For i = 1 To Map(z).EventCount
                With Map(z).Events(i)
                    .Name = GetVar(filename, "Event" & i, "Name")
                    .Global = Val(GetVar(filename, "Event" & i, "Global"))
                    .x = Val(GetVar(filename, "Event" & i, "x"))
                    .y = Val(GetVar(filename, "Event" & i, "y"))
                    .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
                End With
                If Map(z).Events(i).PageCount > 0 Then
                    ReDim Map(z).Events(i).Pages(0 To Map(z).Events(i).PageCount)
                    For x = 1 To Map(z).Events(i).PageCount
                        With Map(z).Events(i).Pages(x)
                            .chkVariable = Val(GetVar(filename, "Event" & i & "Page" & x, "chkVariable"))
                            .VariableIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableIndex"))
                            .VariableCondition = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCondition"))
                            .VariableCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCompare"))
                            
                            .chkSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSwitch"))
                            .SwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchIndex"))
                            .SwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchCompare"))
                            
                            .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & x, "chkHasItem"))
                            .HasItemIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemIndex"))
                            
                            .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSelfSwitch"))
                            .SelfSwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchIndex"))
                            .SelfSwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchCompare"))
                            
                            .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicType"))
                            .Graphic = Val(GetVar(filename, "Event" & i & "Page" & x, "Graphic"))
                            .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX"))
                            .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY"))
                            .GraphicX2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX2"))
                            .GraphicY2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY2"))
                            
                            .MoveType = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveType"))
                            .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveSpeed"))
                            .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveFreq"))
                            
                            .IgnoreMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "IgnoreMoveRoute"))
                            .RepeatMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "RepeatMoveRoute"))
                            
                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRouteCount"))
                            
                            If .MoveRouteCount > 0 Then
                                ReDim Map(z).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                                For y = 1 To .MoveRouteCount
                                    .MoveRoute(y).index = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Index"))
                                    .MoveRoute(y).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data1"))
                                    .MoveRoute(y).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data2"))
                                    .MoveRoute(y).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data3"))
                                    .MoveRoute(y).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data4"))
                                    .MoveRoute(y).data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data5"))
                                    .MoveRoute(y).data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data6"))
                                Next
                            End If
                            
                            .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkAnim"))
                            .DirFix = Val(GetVar(filename, "Event" & i & "Page" & x, "DirFix"))
                            .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkThrough"))
                            .ShowName = Val(GetVar(filename, "Event" & i & "Page" & x, "ShowName"))
                            .Trigger = Val(GetVar(filename, "Event" & i & "Page" & x, "Trigger"))
                            .CommandListCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandListCount"))
                         
                            .Position = Val(GetVar(filename, "Event" & i & "Page" & x, "Position"))
                        End With
                            
                        If Map(z).Events(i).Pages(x).CommandListCount > 0 Then
                            ReDim Map(z).Events(i).Pages(x).CommandList(0 To Map(z).Events(i).Pages(x).CommandListCount)
                            For y = 1 To Map(z).Events(i).Pages(x).CommandListCount
                                Map(z).Events(i).Pages(x).CommandList(y).CommandCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "CommandCount"))
                                Map(z).Events(i).Pages(x).CommandList(y).ParentList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "ParentList"))
                                If Map(z).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                    ReDim Map(z).Events(i).Pages(x).CommandList(y).Commands(Map(z).Events(i).Pages(x).CommandList(y).CommandCount)
                                    For p = 1 To Map(z).Events(i).Pages(x).CommandList(y).CommandCount
                                        With Map(z).Events(i).Pages(x).CommandList(y).Commands(p)
                                            .index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Index"))
                                            .Text1 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text1")
                                            .Text2 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text2")
                                            .Text3 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text3")
                                            .Text4 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text4")
                                            .Text5 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text5")
                                            .Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data1"))
                                            .Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data2"))
                                            .Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data3"))
                                            .Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data4"))
                                            .data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data5"))
                                            .data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data6"))
                                            .ConditionalBranch.CommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchCommandList"))
                                            .ConditionalBranch.Condition = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchCondition"))
                                            .ConditionalBranch.Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData1"))
                                            .ConditionalBranch.Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData2"))
                                            .ConditionalBranch.Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData3"))
                                            .ConditionalBranch.ElseCommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchElseCommandList"))
                                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRouteCount"))
                                            If .MoveRouteCount > 0 Then
                                                ReDim .MoveRoute(1 To .MoveRouteCount)
                                                For w = 1 To .MoveRouteCount
                                                    .MoveRoute(w).index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Index"))
                                                    .MoveRoute(w).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data1"))
                                                    .MoveRoute(w).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data2"))
                                                    .MoveRoute(w).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data3"))
                                                    .MoveRoute(w).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data4"))
                                                    .MoveRoute(w).data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data5"))
                                                    .MoveRoute(w).data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data6"))
                                                Next
                                            End If
                                        End With
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, index)), LenB(MapItem(mapnum, index)))
    MapItem(mapnum, index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal mapnum As Long)
    ReDim MapNpc(mapnum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).NPC(index)), LenB(MapNpc(mapnum).NPC(index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub ClearMap(ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(mapnum)), LenB(Map(mapnum)))
    Map(mapnum).Name = vbNullString
    Map(mapnum).MaxX = MAX_MAPX
    Map(mapnum).MaxY = MAX_MAPY
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal stat As Stats) As Long
    GetClassStat = Class(ClassNum).stat(stat)
End Function

Sub SaveBank(ByVal index As Long)
    Dim filename As String
    Dim f As Long
    
    filename = App.path & "\data\banks\" & Trim$(Player(index).Login) & ".bin"
    
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Bank(index)
    Close #f
End Sub

Public Sub LoadBank(ByVal index As Long, ByVal Name As String)
    Dim filename As String
    Dim f As Long

    Call ClearBank(index)

    filename = App.path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(index)
        Exit Sub
    End If

    f = FreeFile
    Open filename For Binary As #f
        Get #f, , Bank(index)
    Close #f

End Sub

Sub ClearBank(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(index)), LenB(Bank(index)))
End Sub

Sub ClearParty(ByVal partyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String
filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next

End Sub

Sub SaveVariables()
Dim i As Long, filename As String
filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String
filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
Next
End Sub

Sub LoadVariables()
Dim i As Long, filename As String
filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
Next
End Sub

Public Sub SaveRank()

Dim filename As String, i As Byte



filename = App.path & "\data\rank.ini"



For i = 1 To MAX_RANK

PutVar filename, "RANK", "Name" & i, Trim$(Rank(i).Name)

PutVar filename, "RANK", "Level" & i, Val(Rank(i).Level)

Next i

End Sub



Public Sub LoadRank()

Dim filename As String, i As Byte



filename = App.path & "\data\rank.ini"



If FileExist(filename, True) Then

For i = 1 To MAX_RANK

Rank(i).Name = GetVar(filename, "RANK", "Name" & i)

Rank(i).Level = Val(GetVar(filename, "RANK", "Level" & i))

Next i

Else

SaveRank

End If

End Sub

Sub LoadAEditor()
Dim strload As String
Dim i As Long
Dim TotalCount As Long
    frmAccount.lstAccounts.Clear
    strload = Dir(App.path & "\data\accounts\" & "*.bin")
    i = 1
    Do While strload > vbNullString
        frmAccount.lstAccounts.AddItem Mid(strload, 1, Len(strload) - 4)
        strload = Dir
        i = i + 1
    Loop
        
    TotalCount = (i - 1)
    frmAccount.lblAcctCount.Caption = TotalCount
End Sub
