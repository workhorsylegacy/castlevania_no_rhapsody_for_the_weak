Attribute VB_Name = "Module2"
Option Explicit

Public Function CollisionDetect(C As Boolean, intNewX As Integer, intNewY As Integer, T As Integer) As Integer ' Alucard hit Enviornment
Dim DCwidth As Long, DCheight As Long
Dim W As Long, H As Long, P As Long ' Width, Height and Pixel
Dim L1 As Long, L2 As Long, L3 As Long ' Three loop variables
Dim LP As Long, V As Boolean, A As Long, B As Long ' Loop, Vertical GetPixel and Swap W and H
Dim L As Long, R As Long, U As Long, D As Long ' Four Directions
Dim cBlack As String, cGreen As String, cRed As String, cPurple As String, cYellow As String, cBlue As String
    With Player
        cBlack = "&H000000": cGreen = "&H00FF00": cRed = "&H0000FF": cPurple = "&HFF009C": cYellow = "&H00FFFF": cBlue = "&HFF0000"
        'Reset variables
        .Teeter = False
        .NY = 0
        .Soft = False
        ' SHOW FOR TEST
        frmCollision.picScreen.Cls
        .WallR = False: .WallL = False: .WallU = False: .WallD = False
        'Get box size
        DCwidth = .Widthp(.ANI) + .Xpos(.ANI)
        DCheight = .Heightp(.ANI) + .Ypos(.ANI)
        'Clean and Draw into DC
        BitBlt memCollEnv, 0, 0, DCwidth, DCheight, memBlank, 0, 0, &HCC0020 ' SRCCOPY
        BitBlt memCollEnv, 0, 0, DCwidth, DCheight, memBgCollision, (.PX + intNewX - BgX), (.PY + intNewY - BgY), &H660046 ' SRCINVERT
        ' SHOW FOR TEST
        BitBlt frmCollision.picScreen.hdc, 0, 0, DCwidth, DCheight, memCollEnv, 0, 0, &HCC0020 ' SRCCOPY
        'Examine DC
        For LP = 0 To 3
            Select Case LP
                Case 0: L1 = .BT: L2 = .BL: L3 = .BW: V = False
                Case 1: L1 = .BH: L2 = .BL: L3 = .BW: V = False
                Case 2: L1 = .BL: L2 = .BT: L3 = .BH: V = True
                Case 3: L1 = .BW: L2 = .BT: L3 = .BH: V = True
            End Select
            A = L1
            For B = L2 To L3
                Select Case V
                    Case True: W = A: H = B
                    Case False: W = B: H = A
                End Select
                P = GetPixel(memCollEnv, W, H)
                'Debug.Print PX(CH); PY(CH)
                ' SHOW FOR TEST
                SetPixel frmCollision.picScreen.hdc, W, H, 255
                Select Case P
                    Case cBlack ' Black - Collision
                        .mblnSink = False ' Turn off down collision ignorer
                        C = True
                        Select Case LP
                            Case 0: U = U + 1  ' Up
                            Case 1: D = D + 1  ' Down
                            Case 2: L = L + 1  ' Left
                            Case 3: R = R + 1  ' Right
                        End Select
                    Case cGreen ' Green - Soft Platform
                        If T = COLLISIONDOWN Then
                            .Soft = True
                            C = True
                            Select Case LP
                                Case 0: U = U + 1  ' Up
                                Case 1: D = D + 1  ' Down
                                Case 2: L = L + 1  ' Left
                                Case 3: R = R + 1  ' Right
                            End Select
                        End If
                    Case cRed: pblnDrawSplash(CH) = True ' Red - Water Splash
                    Case cPurple: .Water = True ' Purple - Water
                    Case cYellow: .Teeter = True: .StairL = True ' Yellow - Stairs left
                    Case cBlue: .Teeter = True: .StairR = True ' Blue - Stairs right
                End Select
            Next
        Next
        'SHOW FOR TEST
        If .pLeft = False Then ' Draws Mask
            BitBlt frmCollision.picScreen.hdc, .Xpos(.ANI), .Ypos(.ANI), .Widthp(.ANI), .Heightp(.ANI), .memMsk, .Xp(.ANI), .Yp(.ANI), &H8800C6  ' SRCAND
        Else
            StretchBlt frmCollision.picScreen.hdc, .LXP + .Widthp(.ANI), .Ypos(.ANI), -.Widthp(.ANI), .Heightp(.ANI), .memMsk, .Xp(.ANI), .Yp(.ANI), .Widthp(.ANI), .Heightp(.ANI), &H8800C6  ' SRCAND
        End If
        'Teetering
        If (L <> 0 And R = 0) Or (R <> 0 And L = 0) Then .Teeter = True
        'Splash effect
        If .Water = True And EfctDone(CH) = True Then pblnDrawSplash(CH) = False
        'Fix Position Sinkange
        Select Case T
            Case COLLISIONDOWN
                If .Soft = True And (L > 9 Or R > 9 Or U <> 0 Or D < 1) Then C = False ' Don't jump onto it until the top is reached. R & L for height before on platform. U for no headbuts. D for no getting stuck on the torso.
                If R >= L And D <> 0 Then ' Move down to the heighest side. D <> 0 for ledge sinking
                    .NY = intNewY - R
                ElseIf L > R And D <> 0 Then
                    .NY = intNewY - L
                End If
            Case COLLISIONUP
                Select Case R
                    Case Is > L: .NY = intNewY + R
                    Case Is >= R: .NY = intNewY + L
                End Select
        End Select
        If .mblnSink = True Then C = False ' Ignore collision detection
        'UnPin player from between wall and enemy
        If .HURT = True Then
            If .WallR = True Then
                .pLeft = False
            ElseIf .WallL = True Then
                .pLeft = True
            End If
        End If
    End With
End Function

Public Function ECollisionDetect(C As Boolean, intNewX As Integer, intNewY As Integer, T As Integer) As Integer ' Enemy hit Enviornment
Dim DCwidth As Long, DCheight As Long
Dim W As Long, H As Long, P As Long ' Width, Height and Pixel
Dim L1 As Long, L2 As Long, L3 As Long ' Three loop variables
Dim LP As Long, V As Boolean, A As Long, B As Long ' Loop, Vertical GetPixel and Swap W and H
Dim L As Long, R As Long, U As Long, D As Long ' Four Directions
Dim cBlack As String, cGreen As String, cRed As String, cPurple As String, cYellow As String, cBlue As String
    With Enemy(CH)
        cBlack = "&H000000": cGreen = "&H00FF00": cRed = "&H0000FF": cPurple = "&HFF009C": cYellow = "&H00FFFF": cBlue = "&HFF0000"
        'Reset variables
        .NY = 0
        ' SHOW FOR TEST
        'If CH = 1 Then frmCollision.picScreen.Cls
        .WallR = False: .WallL = False: .WallU = False: .WallD = False
        'Get box size
        DCwidth = EWidthp(.ANI, .EType) + EXpos(.ANI, .EType)
        DCheight = EHeightp(.ANI, .EType) + EYpos(.ANI, .EType)
        'Clean and Draw into DC
        BitBlt memCollEnv, 0, 0, DCwidth, DCheight, memBlank, 0, 0, &HCC0020 ' SRCCOPY
        BitBlt memCollEnv, 0, 0, DCwidth, DCheight, memBgCollision, (.PX + intNewX), (.PY + intNewY), &H660046 ' SRCINVERT
        ' SHOW FOR TEST
        'If CH = 1 Then BitBlt frmCollision.picScreen.hdc, 0, 0, DCwidth, DCheight, memCollEnv, 0, 0, &HCC0020 ' SRCCOPY
        'Examine DC
        For LP = 0 To 3 ' Loop once for each side
            Select Case LP
                Case 0: L1 = .BT: L2 = .BL: L3 = .BW: V = False
                Case 1: L1 = .BH: L2 = .BL: L3 = .BW: V = False
                Case 2: L1 = .BL: L2 = .BT: L3 = .BH: V = True
                Case 3: L1 = .BW: L2 = .BT: L3 = .BH: V = True
            End Select
            A = L1
            For B = L2 To L3
                Select Case V ' Swap for vertical or horizontal
                    Case True: W = A: H = B
                    Case False: W = B: H = A
                End Select
                P = GetPixel(memCollEnv, W, H)
                ' SHOW FOR TEST
                'If CH = 1 Then SetPixel frmCollision.picScreen.hdc, W, H, 255
                Select Case P
                    Case cBlack ' Black - Collision
                        C = True
                        Select Case LP
                            Case 0: U = U + 1  ' Up
                            Case 1: D = D + 1  ' Down
                            Case 2: L = L + 1  ' Left
                            Case 3: R = R + 1  ' Right
                        End Select
                    Case cGreen ' Green - Soft Platform
                        If T = COLLISIONDOWN Then
                            .Soft = True
                            C = True
                            Select Case LP
                                Case 0: U = U + 1  ' Up
                                Case 1: D = D + 1  ' Down
                                Case 2: L = L + 1  ' Left
                                Case 3: R = R + 1  ' Right
                            End Select
                        End If
                    Case cRed: pblnDrawSplash(CH) = True ' Red - Water Splash
                    Case cPurple: .HURT = True ' Purple - Water
                    Case cYellow: .StairL = True ' Yellow - Stairs left
                    Case cBlue: .StairR = True ' Blue - Stairs right
                End Select
            Next
        Next
        'SHOW FOR TEST
        'If CH = 1 Then BitBlt frmCollision.picScreen.hdc, EXpos(ANI(CH), EType(CH)), EYpos(ANI(CH), EType(CH)), EWidthp(ANI(CH), EType(CH)), EHeightp(ANI(CH), EType(CH)), memEMsk(EType(CH)), EXp(ANI(CH), EType(CH)), EYp(ANI(CH), EType(CH)), &H660046 ' SRCINVERT'&H8800C6 ' SRCAND
        'If CH = 1 Then StretchBlt frmCollision.picScreen.hdc, LXP(CH) + EWidthp(ANI(CH), EType(CH)), EYpos(ANI(CH), EType(CH)), -EWidthp(ANI(CH), EType(CH)), EHeightp(ANI(CH), EType(CH)), memEMsk(EType(CH)), EXp(ANI(CH), EType(CH)), EYp(ANI(CH), EType(CH)), EWidthp(ANI(CH), EType(CH)), EHeightp(ANI(CH), EType(CH)), &H660046 ' SRCINVERT'&H8800C6 ' SRCAND
        'Splash effect
        If .Water = True And EfctDone(CH) = True Then pblnDrawSplash(CH) = False
        Select Case T
            Case COLLISIONDOWN
                If .Soft = True And (L > 9 Or R > 9 Or U <> 0) Then C = False ' Don't jump onto it until the top is reached
                If R >= L And D <> 0 Then ' Move down to the heighest side. D <> 0 for ledge sinking
                    .NY = intNewY - R
                ElseIf L > R And D <> 0 Then
                    .NY = intNewY - L
                End If
            Case COLLISIONUP
                Select Case R
                    Case Is > L: .NY = intNewY + R
                    Case Is >= R: .NY = intNewY + L
                End Select
        End Select
        'Don't get your uper half stuck in the ceiling
        'If (R > 0 Or L > 0) And U > 0 And D = 0 Then C = False
    End With
End Function

Public Function SCollisionDetect(C As Boolean, H As Long, T As Long) As Boolean ' Enemy hit Alucard/Weapon hit Enemy
Dim P As Long
Dim W1 As Long, H1 As Long ' Bitblt width and height
'Dim BT As Long, BL As Long, BW As Long, BH As Long ' Box Dimentions
'Dim LW As Long, LH As Long ' Loop width and Height
Dim F As Long
Dim Pixel(34000) As Long, PS As Long, PE As Long
Dim Extra As Long
    'Weapon VS Enemy
    'right wall >= left wall, bottom wall >= top wall, left wall <= right wall, top wall <= bottom wall
    If Player.PX + Player.WepXpos(Player.WEP) + Player.WepW(Player.WEP) >= Enemy(H).PX + BgX + EXpos(Enemy(H).ANI, T) - EWidthp(Enemy(H).ANI, T) And _
       Player.PY + Player.WepYpos(Player.WEP) + Player.WepH(Player.WEP) >= Enemy(H).PY + BgY + EYpos(Enemy(H).ANI, T) And _
       Player.PX - Player.WepW(Player.WEP) - Player.SXP <= Enemy(H).PX + BgX + EXpos(Enemy(H).ANI, T) + EWidthp(Enemy(H).ANI, T) And _
       Player.PY + Player.WepYpos(Player.WEP) <= Enemy(H).PY + BgY + EYpos(Enemy(H).ANI, T) + EHeightp(Enemy(H).ANI, T) And _
       Player.DrwWep = True Then
        'Get box size
        With Player
            If Player.pLeft = True Then
                Extra = .CEN
                W1 = .WepW(.WEP) + .WepXpos(.WEP) + .WepXpos(.WEP)
                H1 = .WepH(.WEP) + .WepYpos(.WEP)
            Else
                W1 = .WepW(.WEP) + .WepXpos(.WEP)
                H1 = .WepH(.WEP) + .WepYpos(.WEP)
            End If
        End With
        'W1 = 150 ' remove
        'H1 = 150 ' remove
        ' SHOW FOR TEST
        With frmCollision.picScreen2: .Width = W1: .Height = H1: End With
        'Clean and Draw into DC - Subtract the other character's x and y positions
        BitBlt memCollCh, 0, 0, W1, H1, memBlank, 0, 0, &HCC0020 ' SRCCOPY
        With Enemy(H)
            If .pLeft = True Then
                StretchBlt memCollCh, .PX + .LXP + EWidthp(.ANI, T) + BgX - Player.PX + Extra, .PY + EYpos(.ANI, T) + BgY - Player.PY, -EWidthp(.ANI, T), EHeightp(.ANI, T), memEMsk(T), EXp(.ANI, T), EYp(.ANI, T), EWidthp(.ANI, T), EHeightp(.ANI, T), &HCC0020 ' SRCCOPY
            Else
                StretchBlt memCollCh, .PX + EXpos(.ANI, T) + BgX - Player.PX + Extra, .PY + EYpos(.ANI, T) + BgY - Player.PY, EWidthp(.ANI, T), EHeightp(.ANI, T), memEMsk(T), EXp(.ANI, T), EYp(.ANI, T), EWidthp(.ANI, T), EHeightp(.ANI, T), &HCC0020 ' SRCCOPY
            End If
        End With
        'Examine DC for collision
        F = Player.WEP
        PS = WepCollS(F)
        PE = WepCollE(F)
        With Player
            For Pixel(F) = PS To PE
                If .pLeft = True Then
                    P = GetPixel(memCollCh, .WepW(F) - (WepCollX(Pixel(F)) - .SXP) + Extra, WepCollY(Pixel(F)) + .WepYpos(F))
                    ' SHOW FOR TEST
                    SetPixel memCollCh, .WepW(F) - (WepCollX(Pixel(F)) - .SXP) + Extra, WepCollY(Pixel(F)) + .WepYpos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        'Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        SetPixel memCollCh, .WepW(F) - (WepCollX(Pixel(F)) - .SXP) + Extra, WepCollY(Pixel(F)) + .WepYpos(F), 255
                    End If
                Else
                    P = GetPixel(memCollCh, WepCollX(Pixel(F)) + .WepXpos(F), WepCollY(Pixel(F)) + .WepYpos(F))
                    ' SHOW FOR TEST
                    SetPixel memCollCh, WepCollX(Pixel(F)) + .WepXpos(F), WepCollY(Pixel(F)) + .WepYpos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        'Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        SetPixel memCollCh, WepCollX(Pixel(F)) + .WepXpos(F), WepCollY(Pixel(F)) + .WepYpos(F), 255
                    End If
                End If
            Next
        End With
        'Sword hit enemy
        If C = True Then Enemy(H).HrtingCh = CH
        If C = True And mblnAttack = True Then C = False: Enemy(H).HURT = True: Enemy(H).HP = Enemy(H).HP - Player.AP
        ' SHOW FOR TEST
        BitBlt frmCollision.picScreen2.hdc, 0, 0, frmCollision.picScreen2.Width, frmCollision.picScreen2.Height, memCollCh, 0, 0, &HCC0020 ' SRCCOPY
'Alucard's body VS Enemy
    ElseIf Player.PX + Player.Xpos(Player.ANI) + Player.Widthp(Player.ANI) >= Enemy(H).PX + BgX + EXpos(Enemy(H).ANI, T) - EWidthp(Enemy(H).ANI, T) And _
       Player.PY + Player.Ypos(Player.ANI) + Player.Heightp(Player.ANI) >= Enemy(H).PY + BgY + EYpos(Enemy(H).ANI, T) And _
       Player.PX - Player.Widthp(Player.ANI) + Player.Xpos(Player.ANI) <= Enemy(H).PX + BgX + EXpos(Enemy(H).ANI, T) + EWidthp(Enemy(H).ANI, T) And _
       Player.PY + Player.Ypos(Player.ANI) <= Enemy(H).PY + BgY + EYpos(Enemy(H).ANI, T) + EHeightp(Enemy(H).ANI, T) Then
        'Get box size
        With Player
            If .pLeft = True Then
                W1 = .Widthp(.ANI) + .LXP
                H1 = .Heightp(.ANI) + .Ypos(.ANI)
            Else
                W1 = .Widthp(.ANI) + .Xpos(.ANI)
                H1 = .Heightp(.ANI) + .Ypos(.ANI)
            End If
        End With
        ' SHOW FOR TEST
        With frmCollision.picScreen2: .Width = W1: .Height = H1: End With
        ' SHOW FOR TEST
        'frmCollision.Refresh
        'Clean and Draw into DC
        BitBlt memCollCh, 0, 0, W1, H1, memBlank, 0, 0, &HCC0020 ' SRCCOPY
        With Enemy(H)
            If .pLeft = True Then
                StretchBlt memCollCh, .PX + .LXP + EWidthp(.ANI, T) + BgX - Player.PX, .PY + EYpos(.ANI, T) + BgY - Player.PY, -EWidthp(.ANI, T), EHeightp(.ANI, T), memEMsk(T), EXp(.ANI, T), EYp(.ANI, T), EWidthp(.ANI, T), EHeightp(.ANI, T), &HCC0020 ' SRCCOPY
            Else
                StretchBlt memCollCh, .PX + EXpos(.ANI, T) + BgX - Player.PX, .PY + EYpos(.ANI, T) + BgY - Player.PY, EWidthp(.ANI, T), EHeightp(.ANI, T), memEMsk(T), EXp(.ANI, T), EYp(.ANI, T), EWidthp(.ANI, T), EHeightp(.ANI, T), &HCC0020 ' SRCCOPY
            End If
        End With
        'Examine DC for collision
        F = Player.ANI
        PS = CollS(F)
        PE = CollE(F)
        With Player
            For Pixel(F) = PS To PE
                If .pLeft = True Then
                    P = GetPixel(memCollCh, .Widthp(F) - (CollX(Pixel(F)) - .LXP), (CollY(Pixel(F)) + .Ypos(F)))
                    ' SHOW FOR TEST
                    SetPixel memCollCh, .Widthp(F) - (CollX(Pixel(F)) - .LXP), CollY(Pixel(F)) + .Ypos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        'Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        SetPixel memCollCh, .Widthp(F) - (CollX(Pixel(F)) - .LXP), CollY(Pixel(F)) + .Ypos(F), 255
                    End If
                Else
                    P = GetPixel(memCollCh, CollX(Pixel(F)) + .Xpos(F), CollY(Pixel(F)) + .Ypos(F))
                    ' SHOW FOR TEST
                    SetPixel memCollCh, CollX(Pixel(F)) + .Xpos(F), CollY(Pixel(F)) + .Ypos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        'Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        SetPixel memCollCh, CollX(Pixel(F)) + .Xpos(F), CollY(Pixel(F)) + .Ypos(F), 255
                    End If
                End If
            Next
        End With
        If C = True Then Player.HrtingCh = H ' Get Attack Value of Character Attacking to Measure Damage
        ' SHOW FOR TEST
        BitBlt frmCollision.picScreen2.hdc, 0, 0, frmCollision.picScreen2.Width, frmCollision.picScreen2.Height, memCollCh, 0, 0, &HCC0020 ' SRCCOPY
    End If
End Function

Public Function OCollisionDetect(H As Long) As Boolean ' Weapon hit Destroyable Object
Dim P As Long
Dim W1 As Long, H1 As Long ' Bitblt width and height
Dim BT As Long, BL As Long, BW As Long, BH As Long ' Box Dimentions
Dim LW As Long, LH As Long ' Loop width and Height
'Dim blnExit As Boolean ' Exit Loop
Dim F As Long
Dim Pixel(34000) As Long, PS As Long, PE As Long
Dim Extra As Long
Dim C As Boolean
    With Player
        ' SHOW FOR TEST
        ''frmCollision.picScreen2.Cls
        'Weapon Hit Object
        'right wall >= left wall, bottom wall >= top wall, left wall <= right wall, top wall <= bottom wall
        If .PX + .WepXpos(.WEP) + .WepW(.WEP) >= BgX + PsXpos(H) And _
           .PY + .WepYpos(.WEP) + .WepH(.WEP) >= BgY + PsYpos(H) And _
           .PX - .WepW(.WEP) - .SXP <= BgX + PsXpos(H) + PsWp(H) And _
           .PY + .WepYpos(.WEP) <= BgY + PsYpos(H) + PsHp(H) And _
           .DrwWep = True Then
            'Get box size
            If .pLeft = True Then
                Extra = .CEN
                W1 = .WepW(.WEP) + .WepXpos(.WEP) + .WepXpos(.WEP)
                H1 = .WepH(.WEP) + .WepYpos(.WEP)
            Else
                W1 = .WepW(.WEP) + .WepXpos(.WEP)
                H1 = .WepH(.WEP) + .WepYpos(.WEP)
            End If
            'W1 = 100 ' remove
            'H1 = 100 ' remove
            ' SHOW FOR TEST
            'With frmCollision.picScreen2: .Width = W1: .Height = H1: End With
            ' SHOW FOR TEST
            ''frmCollision.Refresh
            'Black Blt Source - object - scrcopy
            'Subtract the object's x and y positions
            'Clean and Draw into DC
            BitBlt memCollCh, 0, 0, W1, H1, memBlank, 0, 0, &HCC0020 ' SRCCOPY
            'StretchBlt memCollCh, PsXpos(H) + BgX - PX(CH) + CEN(CH), PsYpos(H) + BgY - PY(CH), PsWp(H), PsHp(H), memBgPsMsk, PsXp(mintCandle), PsYp(mintCandle), PsWp(H), PsHp(H), &HCC0020   ' SRCCOPY
            StretchBlt memCollCh, PsXpos(H) + BgX - .PX + Extra, PsYpos(H) + BgY - .PY, PsWp(H), PsHp(H), memBgPsMsk, PsXp(mintCandle), PsYp(mintCandle), PsWp(H), PsHp(H), &HCC0020 ' SRCCOPY
            'Examine DC for collision
            F = .WEP
            PS = WepCollS(F)
            PE = WepCollE(F)
            For Pixel(F) = PS To PE
                If .pLeft = True Then
                    P = GetPixel(memCollCh, .WepW(F) - (WepCollX(Pixel(F)) - .SXP) + Extra, (WepCollY(Pixel(F)) + .WepYpos(F)))
                    ' SHOW FOR TEST
                    'SetPixel memCollCh, WepW(F) - (WepCollX(Pixel(F)) - SXP(CH)) + Extra, WepCollY(Pixel(F)) + WepYpos(F), 16711680
                    If P = 0 Then
                        C = True
                        Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        'SetPixel memCollCh, WepW(F) - (WepCollX(Pixel(F)) - SXP(CH)) + Extra, WepCollY(Pixel(F)) + WepYpos(F), 255
                    End If
                Else
                    P = GetPixel(memCollCh, WepCollX(Pixel(F)) + .WepXpos(F), WepCollY(Pixel(F)) + .WepYpos(F))
                    ' SHOW FOR TEST
                    'SetPixel memCollCh, WepCollX(Pixel(F)) + WepXpos(F), WepCollY(Pixel(F)) + WepYpos(F), 16711680
                    If P = 0 Then
                        C = True
                        Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        'SetPixel memCollCh, WepCollX(Pixel(F)) + WepXpos(F), WepCollY(Pixel(F)) + WepYpos(F), 255
                    End If
                End If
            Next
            'Sword Hit Object
            If C = True And mblnAttack = True Then
                'Create New Item
                If PItm(H) > 0 Then
                    TotlItm = TotlItm + 1
                    ItmType(TotlItm) = PItm(H) ' Add As Next
                    IXpos(TotlItm) = PsXpos(H) ' Add X Coordinate
                    IYpos(TotlItm) = PsYpos(H) ' Add Y Coordinate
                    ItmFat(TotlItm) = PItmFat(H) ' Add Item's Weight
                End If
                'Wipe Out Old Object
                If H = TotlCand Then ' Delete
                    TotlCand = TotlCand - 1
                Else ' Replace With Last, Then Delete Last
                    PsXpos(H) = PsXpos(TotlCand)
                    PsYpos(H) = PsYpos(TotlCand)
                    PsXp(H) = PsXp(TotlCand)
                    PsYp(H) = PsYp(TotlCand)
                    PsWp(H) = PsWp(TotlCand)
                    PsHp(H) = PsHp(TotlCand)
                    PItm(H) = PItm(TotlCand)
                    PsDth(H) = PsDth(TotlCand)
                    TotlCand = TotlCand - 1
                End If
            End If
            ' SHOW FOR TEST
            'BitBlt frmCollision.picScreen2.hdc, 0, 0, frmCollision.picScreen2.Width, frmCollision.picScreen2.Height, memCollCh, 0, 0, &HCC0020 ' SRCCOPY
        End If
    End With
End Function

Public Function ICollisionDetect(C As Boolean, H As Long) As Boolean ' Item hit Enviornment
Dim P As Long
Dim F As Long
Dim W1 As Long, H1 As Long ' Width & Height
Dim W2 As Long, H2 As Long ' + 1
    W1 = IWp(ItmType(H)) ' Refered from original item
    H1 = IHp(ItmType(H))
    W2 = W1 + 1
    H2 = H1 + 1
    'Debug.Print W1; H1; H
    ' SHOW FOR TEST
    'frmCollision.picScreen2.Cls
        ' SHOW FOR TEST
        'With frmCollision.picScreen2: .Width = W2: .Height = H2: End With
        ' SHOW FOR TEST
        'frmCollision.Refresh
        'Clean and Draw into DC
        '''''____Temporary____'''' 100 & 100. Use W1 & H1
        BitBlt memCollEnv, 0, 0, W2, H2, memBlank, 0, 0, &HCC0020 ' SRCCOPY
        BitBlt memCollEnv, 0, 0, W2, H2, memBgCollision, IXpos(H), IYpos(H), &H660046 ' SRCINVERT
        'Examine DC for collision
        For F = 0 To W1
            P = GetPixel(memCollEnv, F, H1 - 1) ' -1 for space in image
            ' SHOW FOR TEST
            'SetPixel memCollEnv, F, H1, 16711680
            If P = 0 Then
                C = True
                Exit For  ' Black pixel
                ' SHOW FOR TEST
                'SetPixel memCollEnv, F, H1, 255
            End If
        Next
        ' SHOW FOR TEST
        'BitBlt frmCollision.picScreen2.hdc, 0, 0, W2, H2, memCollEnv, 0, 0, &HCC0020 ' SRCCOPY
End Function

Public Function CCollisionDetect(C As Boolean, H As Long) As Boolean ' Item hit Alucard
Dim P As Long
Dim W1 As Long, H1 As Long ' Bitblt width and height
Dim BT As Long, BL As Long, BW As Long, BH As Long ' Box Dimentions
Dim LW As Long, LH As Long ' Loop width and Height
'Dim blnExit As Boolean ' Exit Loop
Dim F As Long
Dim Pixel(34000) As Long, PS As Long, PE As Long
'Dim Extra As Long
    With Player
        ' SHOW FOR TEST
        ''frmCollision.picScreen2.Cls
        'Weapon Hit Object
        'right wall >= left wall, bottom wall >= top wall, left wall <= right wall, top wall <= bottom wall
        If .PX + .Xpos(.ANI) + .Widthp(.ANI) >= BgX + IXpos(H) And _
           .PY + .Ypos(.ANI) + .Heightp(.ANI) >= BgY + IYpos(H) And _
           .PX + .Xpos(.ANI) <= BgX + IXpos(H) + IWp(ItmType(H)) And _
           .PY + .Ypos(.ANI) <= BgY + IYpos(H) + IHp(ItmType(H)) Then
            'Get box size
            If .pLeft = True Then
                'Extra = CEN(CH)
                W1 = .Widthp(.ANI) + .LXP
                H1 = .Heightp(.ANI) + .Ypos(.ANI)
            Else
                W1 = .Widthp(.ANI) + .Xpos(.ANI)
                H1 = .Heightp(.ANI) + .Ypos(.ANI)
            End If
            
            'W1 = 100 ' remove
            'H1 = 100 ' remove
            ' SHOW FOR TEST
            'With frmCollision.picScreen2: .Width = W1: .Height = H1: End With
            ' SHOW FOR TEST
            ''frmCollision.Refresh
            'Black Blt Source - object - scrcopy
            'Subtract the object's x and y positions
            'Clean and Draw into DC
            BitBlt memCollCh, 0, 0, W1, H1, memBlank, 0, 0, &HCC0020 ' SRCCOPY
            'StretchBlt memCollCh, IXpos(H) + BgX - PX(CH) + CEN(CH), IYpos(H) + BgY - PY(CH), IWp(ItmType(H)), IHp(ItmType(H)), memItmMsk, IXp(ItmType(H)), IYp(ItmType(H)), IWp(ItmType(H)), IHp(ItmType(H)), &HCC0020 ' SRCCOPY
            'StretchBlt memCollCh, PsXpos(H) + BgX - PX(CH) + Extra, PsYpos(H) + BgY - PY(CH), PsWp(H), PsHp(H), memBgPsMsk, PsXp(mintCandle), PsYp(mintCandle), PsWp(H), PsHp(H), &HCC0020 ' SRCCOPY
            StretchBlt memCollCh, IXpos(H) + BgX - .PX, IYpos(H) + BgY - .PY, IWp(ItmType(H)), IHp(ItmType(H)), memItmMsk, IXp(ItmType(H)), IYp(ItmType(H)), IWp(ItmType(H)), IHp(ItmType(H)), &HCC0020  ' SRCCOPY
            'Examine DC for collision
            F = .ANI
            PS = CollS(F)
            PE = CollE(F)
            For Pixel(F) = PS To PE
                If .pLeft = True Then
                    P = GetPixel(memCollCh, .Widthp(F) - (CollX(Pixel(F)) - .LXP), (CollY(Pixel(F)) + .Ypos(F)))
                    ' SHOW FOR TEST
                    'SetPixel memCollCh, Widthp(F) - (CollX(Pixel(F)) - LXP(CH)), CollY(Pixel(F)) + Ypos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        'SetPixel memCollCh, Widthp(F) - (CollX(Pixel(F)) - LXP(CH)), CollY(Pixel(F)) + Ypos(F), 255
                    End If
                Else
                    P = GetPixel(memCollCh, CollX(Pixel(F)) + .Xpos(F), CollY(Pixel(F)) + .Ypos(F))
                    ' SHOW FOR TEST
                    'SetPixel memCollCh, CollX(Pixel(F)) + Xpos(F), CollY(Pixel(F)) + Ypos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        'SetPixel memCollCh, CollX(Pixel(F)) + Xpos(F), CollY(Pixel(F)) + Ypos(F), 255
                    End If
                End If
            Next
            ' Add Items To Inventory
            If C = True Then
                Inven(ItmType(H)) = Inven(ItmType(H)) + 1 ' Add One Of These Items To Inventory
                If H = TotlItm Then ' Just Delete
                    TotlItm = TotlItm - 1
                Else ' Replace With Last Item, Then Delete Last
                    ItmType(H) = ItmType(TotlItm)
                    IXpos(H) = IXpos(TotlItm)
                    IYpos(H) = IYpos(TotlItm)
                    ItmFat(H) = ItmFat(TotlItm)
                    TotlItm = TotlItm - 1
                End If
            End If
            ' SHOW FOR TEST
            'BitBlt frmCollision.picScreen2.hdc, 0, 0, frmCollision.picScreen2.Width, frmCollision.picScreen2.Height, memCollCh, 0, 0, &HCC0020 ' SRCCOPY
        End If
    End With
End Function

Public Function WCollisionDetect(C As Boolean, H As Integer) As Boolean ' Projectile Weapon hit Alucard
Dim P As Long
Dim W1 As Long, H1 As Long ' Bitblt width and height
Dim F As Long
Dim Pixel(34000) As Long, PS As Long, PE As Long
    'right wall >= left wall, bottom wall >= top wall, left wall <= right wall, top wall <= bottom wall
    If Player.PX + Player.Xpos(Player.ANI) + Player.Widthp(Player.ANI) >= Enemy(H).WX + BgX + EWepXpos(Enemy(H).SWEP) - EWepW(Enemy(H).SWEP) And _
       Player.PY + Player.Ypos(Player.ANI) + Player.Heightp(Player.ANI) >= Enemy(H).WY + BgY + EWepYpos(Enemy(H).SWEP) And _
       Player.PX - Player.Widthp(Player.ANI) + Player.Xpos(Player.ANI) <= Enemy(H).WX + BgX + EWepXpos(Enemy(H).SWEP) + EWepW(Enemy(H).SWEP) And _
       Player.PY + Player.Ypos(Player.ANI) <= Enemy(H).WY + BgY + EWepYpos(Enemy(H).SWEP) + EWepH(Enemy(H).SWEP) Then
        'Get box size
        With Player
            If .pLeft = True Then
                'Extra = CEN(CH)
                W1 = .Widthp(.ANI) + .LXP
                H1 = .Heightp(.ANI) + .Ypos(.ANI)
            Else
                W1 = .Widthp(.ANI) + .Xpos(.ANI)
                H1 = .Heightp(.ANI) + .Ypos(.ANI)
            End If
        End With
        ' SHOW FOR TEST
        'With frmCollision.picScreen2: .Width = W1: .Height = H1: End With
        'Clean and Draw into DC - Subtract the object's x and y positions
        BitBlt memCollCh, 0, 0, W1, H1, memBlank, 0, 0, &HCC0020 ' SRCCOPY
        With Enemy(H)
            If .wLeft = True Then
                StretchBlt memCollCh, .WX + .WXP + EWepW(.SWEP) + BgX - Player.PX, .WY + EWepY(.SWEP) + BgY - Player.PY, -EWepW(.SWEP), EWepH(.SWEP), memSubMsk, EWepX(.SWEP), EWepY(.SWEP), EWepW(.SWEP), EWepH(.SWEP), &HCC0020 ' SRCCOPY
            Else
                StretchBlt memCollCh, .WX + EWepXpos(.SWEP) + BgX - Player.PX, .WY + EWepYpos(.SWEP) + BgY - Player.PY, EWepW(.SWEP), EWepH(.SWEP), memSubMsk, EWepX(.SWEP), EWepY(.SWEP), EWepW(.SWEP), EWepH(.SWEP), &HCC0020 ' SRCCOPY
            End If
        End With
        'Examine DC for collision
        F = Player.ANI
        PS = CollS(F)
        PE = CollE(F)
        With Player
            For Pixel(F) = PS To PE
                If .pLeft = True Then
                    P = GetPixel(memCollCh, .Widthp(F) - (CollX(Pixel(F)) - .LXP), (CollY(Pixel(F)) + .Ypos(F)))
                    ' SHOW FOR TEST
                    'SetPixel memCollCh, Widthp(F) - (CollX(Pixel(F)) - LXP(CH)), CollY(Pixel(F)) + Ypos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        'SetPixel memCollCh, Widthp(F) - (CollX(Pixel(F)) - LXP(CH)), CollY(Pixel(F)) + Ypos(F), 255
                    End If
                Else
                    P = GetPixel(memCollCh, CollX(Pixel(F)) + .Xpos(F), CollY(Pixel(F)) + .Ypos(F))
                    ' SHOW FOR TEST
                    'SetPixel memCollCh, CollX(Pixel(F)) + Xpos(F), CollY(Pixel(F)) + Ypos(F), 16711680
                    If P = 0 Then
                        C = True
                        .Grav = 0
                        Exit For  ' Black pixel
                        ' SHOW FOR TEST
                        'SetPixel memCollCh, CollX(Pixel(F)) + Xpos(F), CollY(Pixel(F)) + Ypos(F), 255
                    End If
                End If
            Next
        End With
        ' Delete Projectile Weapons
        If C = True Then
            If H = TotlSub Then ' Just Delete
                TotlSub = TotlSub - 1
            Else ' Replace With Last Item, Then Delete Last
                Enemy(H).SWEP = Enemy(TotlSub).SWEP
                Enemy(H).wLeft = Enemy(TotlSub).wLeft
                Enemy(H).WX = Enemy(TotlSub).WX
                Enemy(H).WY = Enemy(TotlSub).WY
                TotlSub = TotlSub - 1
            End If
        End If
        If C = True Then Player.HrtingCh = H ' Get Attack Value of Character Attacking to Measure Damage
        ' SHOW FOR TEST
        'BitBlt frmCollision.picScreen2.hdc, 0, 0, frmCollision.picScreen2.Width, frmCollision.picScreen2.Height, memCollCh, 0, 0, &HCC0020 ' SRCCOPY
    End If
End Function

'Public Function DCollisionDetect(C As Boolean, H As Integer) As Boolean ' Weapon hit Projectiles
'Dim P As Long
'Dim W1 As Long, H1 As Long ' Bitblt width and height
'Dim F As Long
'Dim Pixel(34000) As Long, PS As Long, PE As Long
'Dim Extra As Long
'    'right wall >= left wall, bottom wall >= top wall, left wall <= right wall, top wall <= bottom wall
'    If PX(CH) + WepXpos(WEP(CH)) + WepW(WEP(CH)) >= WX(H) + BgX + EWepXpos(SWEP(H)) - EWepW(SWEP(H)) And _
'       PY(CH) + WepYpos(WEP(CH)) + WepH(WEP(CH)) >= WY(H) + BgY + EWepYpos(SWEP(H)) And _
'       PX(CH) - WepW(WEP(CH)) - SXP(CH) <= WX(H) + BgX + EWepXpos(SWEP(H)) + EWepW(SWEP(H)) And _
'       PY(CH) + WepYpos(WEP(CH)) <= WY(H) + BgY + EWepYpos(SWEP(H)) + EWepH(SWEP(H)) Then
'        'Get box size
'        If pLeft(CH) = True Then
'            Extra = CEN(CH)
'            W1 = WepW(WEP(CH)) + WepXpos(WEP(CH)) + WepXpos(WEP(CH))
'            H1 = WepH(WEP(CH)) + WepYpos(WEP(CH))
'        Else
'            W1 = WepW(WEP(CH)) + WepXpos(WEP(CH))
'            H1 = WepH(WEP(CH)) + WepYpos(WEP(CH))
'        End If
'        'W1 = 150 ' remove
'        'H1 = 150 ' remove
'        ' SHOW FOR TEST
'        'With frmCollision.picScreen2: .Width = W1: .Height = H1: End With
'        'Clean and Draw into DC - Subtract the object's x and y positions
'        BitBlt memCollCh, 0, 0, W1, H1, memBlank, 0, 0, &HCC0020 ' SRCCOPY
'        If wLeft(H) = True Then
'            StretchBlt memCollCh, WX(H) + WXP(H) + EWepW(SWEP(H)) + BgX - PX(CH) + Extra, WY(H) + EWepYpos(SWEP(H)) + BgY - PY(CH), -EWepW(SWEP(H)), EWepH(SWEP(H)), memSubMsk, EWepX(SWEP(H)), EWepY(SWEP(H)), EWepW(SWEP(H)), EWepH(SWEP(H)), &HCC0020 ' SRCCOPY
'        Else
'            StretchBlt memCollCh, WX(H) + EWepXpos(SWEP(H)) + BgX - PX(CH), WY(H) + EWepYpos(SWEP(H)) + BgY - PY(CH), EWepW(SWEP(H)), EWepH(SWEP(H)), memSubMsk, EWepX(SWEP(H)), EWepY(SWEP(H)), EWepW(SWEP(H)), EWepH(SWEP(H)), &HCC0020 ' SRCCOPY
'        End If
'        'Examine DC for collision
'        F = WEP(CH)
'        PS = WepCollS(F)
'        PE = WepCollE(F)
'        For Pixel(F) = PS To PE
'            If pLeft(CH) = True Then
'                P = GetPixel(memCollCh, WepW(F) - (WepCollX(Pixel(F)) - SXP(CH)) + Extra, WepCollY(Pixel(F)) + WepYpos(F))
'                ' SHOW FOR TEST
'                'SetPixel memCollCh, WepW(F) - (WepCollX(Pixel(F)) - SXP(CH)) + Extra, WepCollY(Pixel(F)) + WepYpos(F), 16711680
'                If P = 0 Then
'                    C = True
'                    Exit For  ' Black pixel
'                    ' SHOW FOR TEST
'                    'SetPixel memCollCh, WepW(F) - (WepCollX(Pixel(F)) - SXP(CH)) + Extra, WepCollY(Pixel(F)) + WepYpos(F), 255
'                End If
'            Else
'                P = GetPixel(memCollCh, WepCollX(Pixel(F)) + WepXpos(F), WepCollY(Pixel(F)) + WepYpos(F))
'                ' SHOW FOR TEST
'                'SetPixel memCollCh, WepCollX(Pixel(F)) + WepXpos(F), WepCollY(Pixel(F)) + WepYpos(F), 16711680
'               If P = 0 Then
'                    C = True
'                    Exit For  ' Black pixel
'                    ' SHOW FOR TEST
'                    'SetPixel memCollCh, WepCollX(Pixel(F)) + WepXpos(F), WepCollY(Pixel(F)) + WepYpos(F), 255
'                End If
'            End If
'        Next
'        'Sword Hit Projectile
'        'If C = True And mblnAttack = True Then
'        '    If H = TotlSub Then ' Just Delete
'        '        TotlSub = TotlSub - 1
'        '    Else ' Replace With Last Item, Then Delete Last
'        '        SWEP(H) = SWEP(TotlSub)
'        '        wLeft(H) = wLeft(TotlSub)
'        '        WX(H) = WX(TotlSub)
'        '        WY(H) = WY(TotlSub)
'        '        TotlSub = TotlSub - 1
'        '    End If
'        'End If
'        ' SHOW FOR TEST
'        'BitBlt frmCollision.picScreen2.hdc, 0, 0, frmCollision.picScreen2.Width, frmCollision.picScreen2.Height, memCollCh, 0, 0, &HCC0020 ' SRCCOPY
'    End If
'End Function
