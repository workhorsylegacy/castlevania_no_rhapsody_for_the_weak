Attribute VB_Name = "Module1"
'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' Castlevania : No Rhapsody for The Weak
' Programmed by Matt Jones
' -------------------------------------------------------------------
' API DECLARATIONS AND VARIABLES
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


'Public Type MEMORYSTATUS
'        dwLength As Long
'        dwMemoryLoad As Long
'        dwTotalPhys As Long
'        dwAvailPhys As Long
'        dwTotalPageFile As Long
'        dwAvailPageFile As Long
'        dwTotalVirtual As Long
'        dwAvailVirtual As Long
'End Type
'Public memInfo As MEMORYSTATUS
'Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'Bitblt and Constants
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Public Const SRCAND = &H8800C6 '......... (DWORD) dest = source AND dest
'Public Const SRCCOPY = &HCC0020 '.........(DWORD) dest = source
'Public Const SRCERASE = &H440328 '........(DWORD) dest = source AND (NOT dest )
'Public Const SRCINVERT = &H660046 '.......(DWORD) dest = source XOR dest
'Public Const SRCPAINT = &HEE0086 '........(DWORD) dest = source OR dest

'StretchBlt
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'GetTickCount
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'GetPixel
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long



'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' API DECLARATIONS AND VARIABLES
' -------------------------------------------------------------------
' PUBLIC VARIABLES
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Public CollX(12000) As Integer ' 12000 is the total number of lines in the Col.txt file.
Public CollY(12000) As Integer
Public CollS(12000) As Long
Public CollE(12000) As Long
Public WepCollX(1200) As Integer
Public WepCollY(1200) As Integer
Public WepCollS(1200) As Long
Public WepCollE(1200) As Long

Public Type PlayableCharacter
    '---CHARACTER ANIMATIONS
    Xpos(226) As Integer
    Ypos(226) As Integer
    Xp(226) As Integer
    Yp(226) As Integer
    Widthp(226) As Integer
    Heightp(226) As Integer
    LXP As Integer ' Left Xpos
    '---WEAPON ANIMATIONS
    WepXpos(40) As Integer
    WepYpos(40) As Integer
    WepX(40) As Integer
    WepY(40) As Integer
    WepW(40) As Integer
    WepH(40) As Integer
    SXP As Integer ' Left Weapon Xpos
    '---JUMPING AND FALLING
    JD As Currency ' Jump Distance
    JS As Currency ' Jump Space
    JT As Integer ' Jump Traveled
    FD As Currency ' Fall Distance
    FS As Currency ' Fall Space
    '---X AND Y COORDINATES
    PX As Integer ' X Position
    PY As Integer ' Y Position
    '---COORDINATE HELPERS
    NY As Integer ' New Y
    ANI As Long ' Frame of animation
    WEP As Integer ' Weapon frame of animation
    STA As Long ' Character State
    pLeft As Boolean ' Facing right or left?
    '---ATTRIBUTES
    AP As Integer ' Attack power
    DP As Integer ' Defence Points
    HP As Integer ' Health Power
    MaxHP As Integer ' Maximum Health
    HURT As Boolean
    CEN As Integer ' Character's center. used for LXP
    AniDelay As Long ' Animation Delay
    '---ENVIORNMENTAL ATTRIBUTES
    Grav As Integer
    HrtingCh As Integer ' Attacking Character
    Teeter As Boolean ' Ledge to teeter on
    Soft As Boolean ' Soft platform
    mblnSink As Boolean ' Ignore collision detection and fall
    Water As Boolean ' Water
    DrwWep As Boolean ' Draw Weapon?
    '---COLLISION DETECTION
    WallR As Boolean 'Collision with Right Wall
    WallL As Boolean 'Collision with Left Wall
    WallU As Boolean 'Collision with Top Wall
    WallD As Boolean 'Collision with Bottom Wall
    StairR As Boolean 'Collision with Right Stairs
    StairL As Boolean 'Collision with Left Stairs
    '---COLLISION RECTANGLE
    BL As Integer ' Box Left
    BT As Integer ' Box Top
    BW As Integer ' Box Width
    BH As Integer ' Box Height
    '---SPRITES AND MASKS
    memSpr As Long
    memMsk As Long
    memHurt As Long
    memWepSpr As Long
    memWepMsk As Long
    '---SUB-WEAPON ANIMATION
    'SWEP As Integer ' Sub Weapon Frame of Animation
    'WX As Integer ' Sub Weapon X
    'WY As Integer ' Sub Weapon Y
    'wLeft As Boolean ' Sub Weapon right or left?
    'WXP As Integer ' Projectile Sub Weapon Left Xpos
    'WCEN As Integer ' Projectile Weapon's center. Used for Left Xpos
End Type

Public Type NonPlayableCharacter
    '---CHARACTER ANIMATIONS
    LXP As Integer ' Left Xpos
    '---WEAPON ANIMATIONS
    'WepXpos(40) As Integer
    'WepYpos(40) As Integer
    'WepX(40) As Integer
    'WepY(40) As Integer
    'WepW(40) As Integer
    'WepH(40) As Integer
    'SXP As Integer ' Left Weapon Xpos
    '---JUMPING AND FALLING
    'JD As Currency ' Jump Distance
    'JS As Currency ' Jump Space
    'JT As Integer ' Jump Traveled
    'FD As Currency ' Fall Distance
    'FS As Currency ' Fall Space
    '---X AND Y COORDINATES
    PX As Integer ' X PositioN
    PY As Integer ' Y Position
    NY As Integer ' New Y
    '---COORDINATE HELPERS
    ANI As Long ' Frame of animation
    WEP As Integer ' Weapon frame of animation
    STA As Long ' Character State
    pLeft As Boolean ' Facing right or left?
    '---ATTRIBUTES
    AP As Integer ' Attack power
    HP As Integer ' Health Power
    DP As Integer ' Defence Points
    'MaxHP As Integer ' Maximum Health
    HURT As Boolean
    CEN As Integer ' Character's center. used for LXP
    AniDelay As Long ' Animation Delay
    '---ENVIORNMENTAL ATTRIBUTES
    HrtingCh As Integer ' Attacking Character
    'Teeter As Boolean ' Ledge to teeter on
    Soft As Boolean ' Soft platform
    mblnSink As Boolean ' Ignore collision detection and fall
    Water As Boolean ' Water
    'DrwWep As Boolean ' Draw Weapon?
    '---COLLISION DETECTION
    WallR As Boolean 'Collision with Right Wall
    WallL As Boolean 'Collision with Left Wall
    WallU As Boolean 'Collision with Top Wall
    WallD As Boolean 'Collision with Bottom Wall
    StairR As Boolean 'Collision with Right Stairs
    StairL As Boolean 'Collision with Left Stairs
    '---COLLISION RECTANGLE
    BL As Integer ' Box Left
    BT As Integer ' Box Top
    BW As Integer ' Box Width
    BH As Integer ' Box Height
    '---SUB-WEAPON ANIMATION
    SWEP As Integer ' Sub Weapon Frame of Animation
    WX As Integer ' Sub Weapon X
    WY As Integer ' Sub Weapon Y
    WXP As Integer ' Projectile Sub Weapon Left Xpos
    wLeft As Boolean ' Weapon right or left?
    EType As Long ' Type of Enemy. Decides what memDC to use.
    WCEN As Integer ' Projectile Weapon's center. Used for Left Xpos
End Type

Public Player As PlayableCharacter
Public Enemy(20) As NonPlayableCharacter

'---Program Misc
Public TotlEn As Long
Public TotlSub As Integer
Public mblnEndProgram As Boolean
Public mintCandle As Integer
Public TotlCand As Integer
Public TotlItm As Integer
Public TotlUp As Integer
Public DELAY As Long
Public Inven(20) As Long ' Inventory -- inven(item)=number of items
Public ResChanged As Boolean ' If resolution was changed
Public CH As Integer ' Currently selected character
Public CW As Integer ' Currently selected Sub Weapon
Public pblnDrawSplash(20) As Boolean ' Draw Water splash ?
'---Enemy Coordinates
Public EXpos(30, 1 To 5) As Integer
Public EYpos(30, 1 To 5) As Integer
Public EXp(30, 1 To 5) As Integer
Public EYp(30, 1 To 5) As Integer
Public EWidthp(30, 1 To 5) As Integer
Public EHeightp(30, 1 To 5) As Integer
'---HeadsUp Display positions
Public HXpos(40) As Integer
Public HYpos(40) As Integer
Public HXp(40) As Integer
Public HYp(40) As Integer
Public HWidth(40) As Integer
Public HHeight(40) As Integer
Public HDelay(40) As Integer
'---Background Position
Public BgX As Long
Public BgY As Long
Public BgX2 As Long
Public BgY2 As Long
Public BgX3 As Long
Public BgY3 As Long
Public ScrlL As Boolean ' Scroll Left
Public ScrlR As Boolean ' Scroll Right
Public ScrlU As Boolean ' Scroll Up
Public ScrlD As Boolean ' Scroll Down
Public BgCurrent As String
Public BgScroll As Integer
Public BgPath(10) As String
Public BgCurExit(10) As Integer
Public BgSpr(10) As String
Public BgMsk(10) As String
Public BgColl(10) As String
Public BgExitTxt As String
Public BgExitNxt(10) As String
Public BgExitType(10) As Integer
Public BgExitTotal As Integer
Public BgSpr2(10) As String
Public BgMsk2(10) As String
Public BgSpr3(10) As String
Public BgMsk3(10) As String
Public BgPsSpr(10) As String
Public BgPsMsk(10) As String
Public BgUpSpr(10) As String
Public BgUpMsk(10) As String
Public BgUpColl(10) As String
Public BgChTxt(10) As String
Public BgMuTxt(10) As String
Public BgExitX(10) As Integer
Public BgExitY(10) As Integer
Public BgStartX(10) As Integer
Public BgStartY(10) As Integer
Public BgNewX(10) As Integer
Public BgNewY(10) As Integer
Public BgType(10) As Integer
Public BgPsTxt(10) As String
Public BgUpTxt(10) As String
Public BgLayer(2 To 4) As Boolean
Public BgWidth As Integer
Public BgHeight As Integer
'---Headsup Display
Public Head(10) As Integer
Public HeadCount As Integer
'---Effects
Public EfctX(20) As Integer ' Effect X
Public EfctY(20) As Integer ' Effect Y
Public AniEfct(20) As Integer ' Effect frame of animation
Public fX As Integer ' Currently Selected Effect
Public EfctDly(20) As Integer ' Effect Delay
Public EfctDone(20) As Boolean
'---Key Press
Public mblnRight As Boolean
Public mblnLeft As Boolean
Public mblnUp As Boolean
Public mblnDown As Boolean
Public mblnAttack As Boolean
'---Previous screen settings
Public pintScreenWidth As Integer
Public pintScreenHeight As Integer
Public pintScreenColor As Integer
Public pintScreenRefresh As Integer
'---Passable Object Coordinates
Public PsXpos(20) As Integer
Public PsYpos(20) As Integer
Public PsXp(20) As Integer
Public PsYp(20) As Integer
Public PsWp(20) As Integer
Public PsHp(20) As Integer
Public PsDth(20) As Boolean ' Object Death for destruction and effects
Public PItm(20) As Integer ' Item that's produced when destroyed
Public PItmFat(20) As Integer ' Weight of item thats produced (for gravity) passes to ItmFat()
'---UnPassable Object Coordinates
Public UpXpos(20) As Integer
Public UpYpos(20) As Integer
Public UpXp(20) As Integer
Public UpYp(20) As Integer
Public UpWp(20) As Integer
Public UpHp(20) As Integer
'---Item Coordinates
Public IXpos(20) As Integer ' X, set when passable destroyed
Public IYpos(20) As Integer ' Y, set when passable destroyed
Public IXp(20) As Integer
Public IYp(20) As Integer
Public IWp(20) As Integer
Public IHp(20) As Integer
Public ItmType(20) As Integer ' Type of Item to Display
Public ItmFat(20) As Integer ' Weight of Item For Gravity
'---Sub Weapons & Projectile Coordinates
Public EWepXpos(20) As Integer ' Index is not ANI(CH). Use SWEP(Each Wep)
Public EWepYpos(20) As Integer
Public EWepX(20) As Integer
Public EWepY(20) As Integer
Public EWepW(20) As Integer
Public EWepH(20) As Integer

'---Memory DC
Public memScreen As Long
Public memBuffer As Long
Public memBgColl As Long
Public memHeadSpr As Long
Public memHeadMsk As Long
Public memESpr(1 To 100) As Long ' For Each Type of Enemy
Public memEMsk(1 To 100) As Long
Public memSubSpr As Long
Public memSubMsk As Long
Public memBgSpr As Long
Public memBgCollision As Long
Public memBgPsSpr As Long
Public memBgPsMsk As Long
Public memBgUpSpr As Long
Public memBgUpMsk As Long
Public memBgUpColl As Long
Public memBgSpr2 As Long
Public memBgMsk2 As Long
Public memBgSpr3 As Long ''''''''''''
Public memBgMsk3 As Long ''''''''''''
Public memCollEnv As Long
Public memCollCh As Long
Public memBlank As Long
Public memItmSpr As Long
Public memItmMsk As Long

'---Sound variables
Public DX As DirectX8 ' Call DIrect-X 8
Public DS As DirectSound8 ' Call Direct Sound
Public SndEft(20) As DirectSoundSecondaryBuffer8 ' Each Alucard Sound Effect
Public blnSndFX(20) As Boolean ' Is Alucard Sound Effect Present?
Public ESndEft(20) As DirectSoundSecondaryBuffer8 ' Each Enemy Sound Effect
Public EblnSndFX(20) As Boolean ' Is Enamy Sound Effect Present?
Public BgMsic As DirectSoundSecondaryBuffer8 ' Background Music
Public blnMsic As Boolean ' Is Music Present?
Public SndDsc As DSBUFFERDESC ' Describes Sound Buffer

'---State Machine Constants
Public Const DEAD = 0
Public Const DYING = 1
Public Const STANDRIGHT = 2
Public Const STANDLEFT = 3
Public Const ATTACKRIGHT = 4
Public Const ATTACKLEFT = 5
Public Const ATTACKOUTRIGHT = 6
Public Const ATTACKOUTLEFT = 7
Public Const ATTACKDOWNRIGHT = 8
Public Const ATTACKDOWNLEFT = 9
Public Const JUMPATTACKRIGHT = 10
Public Const JUMPATTACKLEFT = 11
Public Const DUCKRIGHT = 12
Public Const DUCKLEFT = 13
Public Const STANDUPRIGHT = 14
Public Const STANDUPLEFT = 15
Public Const DASHRIGHT = 16
Public Const DASHLEFT = 17
Public Const WALKRIGHT = 18
Public Const WALKLEFT = 19
Public Const TURNRIGHT = 20
Public Const TURNLEFT = 21
Public Const JUMPRIGHT = 22
Public Const JUMPLEFT = 23
Public Const FALLRIGHT = 24
Public Const FALLLEFT = 25
Public Const FALLMOVERIGHT = 26
Public Const FALLMOVELEFT = 27
Public Const FALLATTACKRIGHT = 28
Public Const FALLATTACKLEFT = 29
Public Const FALLATTACKRIGHTDOWN = 30
Public Const FALLATTACKLEFTDOWN = 31
Public Const LANDRIGHT = 32
Public Const LANDLEFT = 33
Public Const HURTRIGHT = 34
Public Const HURTLEFT = 35
Public Const STOPRIGHT = 36
Public Const STOPLEFT = 37
Public Const SPAWNING = 38
Public Const UNSPAWNING = 39
Public Const WAITING = 40
'---Misc Constants
Public Const SCREENWIDTH = 320
Public Const SCREENHEIGHT = 240
Public Const DISTANCE = 3
Public Const SPACE = 4
Public Const JUMPHEIGHT = 70
Public Const COLLISIONUP = 1
Public Const COLLISIONDOWN = 2
Public Const COLLISIONRIGHT = 3
Public Const COLLISIONLEFT = 4
