Attribute VB_Name = "Module3"
Option Explicit


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' GENERAL DECLARATIONS
' -------------------------------------------------------------------
' ZOMBIE ANIMATIONS
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Function EnemyDie(W As Long, S As Long, L As Boolean) As Long
'Create Right 9 to 22
    With Enemy(CH)
        If W < 9 Or W > 22 Then W = 9
        If .AniDelay >= 2 Then
            .AniDelay = 0
            If W >= 9 And W < 22 Then
                S = IIf(L = True, DYING, DYING)
                W = W + 1
            ElseIf W >= 22 Then
                .HURT = False
                W = 0
                S = IIf(L = True, DEAD, DEAD)
                'Wipe Out Old Character
                If CH = TotlEn Then ' Delete
                    TotlEn = TotlEn - 1
                Else ' Replace With Last, Then Delete Last
                    .PX = Enemy(TotlEn).PX
                    .PY = Enemy(TotlEn).PY
                    .NY = Enemy(TotlEn).NY
                    .EType = Enemy(TotlEn).EType
                    .ANI = Enemy(TotlEn).ANI
                    .WEP = Enemy(TotlEn).WEP
                    .STA = Enemy(TotlEn).STA
                    .AP = Enemy(TotlEn).AP
                    .HP = Enemy(TotlEn).HP
                    .DP = Enemy(TotlEn).DP
                    .pLeft = Enemy(TotlEn).pLeft
                    .CEN = Enemy(TotlEn).CEN
                    .AniDelay = Enemy(TotlEn).AniDelay
                    .LXP = Enemy(TotlEn).LXP
                    '.SXP = Enemy(TotlEn).SXP
                    '.DrwWep = Enemy(TotlEn).DrwWep
                    pblnDrawSplash(CH) = pblnDrawSplash(TotlEn)
                    .WallR = Enemy(TotlEn).WallR
                    .WallL = Enemy(TotlEn).WallL
                    .WallU = Enemy(TotlEn).WallU
                    .WallD = Enemy(TotlEn).WallD
                    .StairR = Enemy(TotlEn).StairR
                    .StairL = Enemy(TotlEn).StairL
                    .BL = Enemy(TotlEn).BL
                    .BT = Enemy(TotlEn).BT
                    .BW = Enemy(TotlEn).BW
                    .BH = Enemy(TotlEn).BH
                    .HURT = Enemy(TotlEn).HURT
                    TotlEn = TotlEn - 1
                End If
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
    End With
End Function

Function EnemyWait(W As Long, S As Long, L As Boolean) As Long
'Wait 7 to 1
    With Enemy(CH)
        If W < 1 Or W > 7 Then W = 7: S = UNSPAWNING
        If .AniDelay >= 6 Then
            .AniDelay = 0
            If W >= 2 And W <= 7 Then
                W = W - 1
                S = UNSPAWNING
            ElseIf W <= 1 Then
                W = 1
                S = WAITING
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
    End With
End Function

Function CreateZombie(W As Long, S As Long, L As Boolean) As Long
'Create 1 to 7
    With Enemy(CH)
        If W < 1 Or W > 7 Then W = 1
        If .AniDelay >= 6 Then
            .AniDelay = 0
            If W >= 1 And W < 7 Then
                W = W + 1
                S = SPAWNING
            ElseIf W >= 7 Then
                W = 8
                S = IIf(L = True, WALKLEFT, WALKRIGHT)
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
    End With
End Function

Function EnemyWalk(W As Long, D As Long, S As Long, L As Boolean) As Long
'Walk Right 8 to 9
    With Enemy(CH)
        If W < 7 Or W > 9 Then W = 8
        If .pLeft = True Then
            EnemyMoveLeft 1, 1
        Else
            EnemyMoveRight 1, 1
        End If
        If D = 10 Then
            D = 0
            If W >= 7 And W < 9 Then
                S = IIf(L = True, WALKLEFT, WALKRIGHT)
                W = W + 1
            ElseIf W = 9 Then
                W = 8
            End If
        Else
            D = D + 1
        End If
    End With
End Function


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' ZOMBIE ANIMATIONS
' -------------------------------------------------------------------
' ZOMBIE MOVEMENT
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Function EnemyMoveRight(intDistance As Integer, intSpace As Integer) As Integer
    With Enemy(CH)
        ECollisionDetect .WallR, intSpace, 0, COLLISIONRIGHT
        If .WallR = False Or .StairR = True Then
            .PX = .PX + intDistance
        End If
        If .StairR = True Then '/up
            .PY = .PY - intDistance
        ElseIf .StairL = True Then  '\-down
            .PY = .PY + intDistance
        End If
        .StairL = False
        .StairR = False
    End With
End Function

Function EnemyMoveLeft(intDistance As Integer, intSpace As Integer) As Integer
    With Enemy(CH)
        ECollisionDetect .WallL, -intSpace, 0, COLLISIONLEFT
        If .WallL = False Or .StairL = True Then
            .PX = .PX - intDistance
        End If
        If .StairL = True Then ' \-up
            .PY = .PY - intDistance
        ElseIf .StairR = True Then '/-down
            .PY = .PY + intDistance
        End If
        .StairL = False
        .StairR = False
    End With
End Function

Function EnemyMoveDown(intDistance As Integer, intSpace As Integer) As Integer
    With Enemy(CH)
        ECollisionDetect .WallD, 0, intSpace + 3, COLLISIONDOWN
        If .WallD = False Then
            .PY = .PY + intDistance + 3
        ElseIf .WallD = True Then
            .PY = .PY + .NY
        End If
    End With
End Function

Function EnemyMoveUp(intDistance As Integer, intSpace As Integer) As Integer
    With Enemy(CH)
        ECollisionDetect .WallU, 0, -intSpace - 3, COLLISIONUP
        If .WallU = False Then
            .PY = .PY - intDistance - 3
        ElseIf .WallU = True Then
            .PY = .PY - .NY
        End If
    End With
End Function

