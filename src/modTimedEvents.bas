Attribute VB_Name = "modTimedEvents"
Public Sub TimedEvents()
Dim i As Long
Dim MyDate As String
    If Hour(Now) = 24 Then
        For i = 1 To Player_HighIndex
            If Player(i).IsMember = 1 Then
                MyDate = Format(Date, "m/d/yyyy")
                If DateDiff("d", Player(i).DateCount, MyDate) >= 31 Then
                    PlayerMsg i, "Your membership has expired.", BrightRed
                    MemberUnEquipItem i
                    If Map(GetPlayerMap(i)).IsMember > 0 Then
                        PlayerWarp i, Map(GetPlayerMap(i)).BootMap, Map(GetPlayerMap(i)).BootX, Map(GetPlayerMap(i)).BootY
                    End If
                    Player(i).IsMember = 0
                    SavePlayer i
                Else
                    PlayerMsg i, "You have " & (31 - DateDiff("d", Player(i).DateCount, MyDate)) & " days remaining of your membership!", Yellow
                End If
            End If
        Next
    End If
End Sub

