VERSION 5.00
Begin VB.Form frmChangeRez 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form4"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picScreenBox 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmChangeRez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' Castlevania : No Rhapsody for The Weak
' Resolution Changer
' Programmed by Matt Jones
' -------------------------------------------------------------------
' GENERAL DECLARATIONS
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/

Option Explicit

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const DISP_CHANGE_SUCCESSFUL = 0

'Const HWND_BROADCAST = &HFFFF&
'Const WM_DISPLAYCHANGE = &H7E&
'Const SPI_SETNONCLIENTMETRICS = 42

Const SCREENWIDTH = 320
Const SCREENHEIGHT = 240

Const STARTMESSAGE = 0
Const ENDMESSAGE = 1

Private Type DEVMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' GENERAL DECLARATIONS
' -------------------------------------------------------------------
' RESOLUTION CHANGE
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Sub StartRez()
Dim RTN As Long
Dim lResult As Long
Dim DevM As DEVMODE ' Refresh Rate
    ResChanged = True
    'Record Width & Height
    pintScreenWidth = Screen.Width \ Screen.TwipsPerPixelX
    pintScreenHeight = Screen.Height \ Screen.TwipsPerPixelY
    'Get Display Info
    lResult = EnumDisplaySettings(0, 0, DevM)
    pintScreenRefresh = DevM.dmDisplayFrequency
    pintScreenColor = DevM.dmBitsPerPel
    'Hide Cursor & Task Bar
    ShowCursor 0
    RTN = FindWindow("Shell_traywnd", "") 'get the Window
    'Change Resolution
    ChangeResolution SCREENWIDTH, SCREENHEIGHT, 16, 0, lResult, DevM, STARTMESSAGE
End Sub

Sub EndRez()
Dim RTN As Long
Dim lResult As Long
Dim DevM As DEVMODE ' Refresh Rate
    If ResChanged = True Then
        'Change Resolution
        lResult = EnumDisplaySettings(0, 0, DevM)
        ChangeResolution pintScreenWidth, pintScreenHeight, pintScreenColor, pintScreenRefresh, lResult, DevM, ENDMESSAGE
        'Show Cursor & Taskbar
        RTN = FindWindow("Shell_traywnd", "") 'get the Window
        SetWindowPos RTN, 0, 0, 0, 0, 0, SWP_SHOWWINDOW 'show the Taskbar
        ShowCursor 1
    End If
End Sub

Private Sub TryAgain()
Dim DevM    As DEVMODE
Dim lResult As Long
    lResult = EnumDisplaySettings(0, 0, DevM)
    With DevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = 640
        .dmPelsHeight = 480
    End With
    lResult = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
    Select Case lResult
        Case Is <> DISP_CHANGE_SUCCESSFUL
            MsgBox "The program could not change the resolution to 640 x 480."
    End Select
End Sub

Private Function ChangeResolution(intWidth As Integer, intHeight As Integer, intColor As Integer, intFrequency As Integer, lResult As Long, DevM As DEVMODE, intMessage As Integer) As Long
Dim lngReply As Long
    With DevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = intWidth
        .dmPelsHeight = intHeight
        .dmBitsPerPel = intColor
        If intFrequency <> 0 Then .dmDisplayFrequency = intFrequency
    End With
    lResult = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
    Select Case lResult
        Case Is <> DISP_CHANGE_SUCCESSFUL
            If intMessage = STARTMESSAGE Then
                MsgBox "Resolution not supported", vbExclamation, "Program Error"
                Exit Function
            ElseIf intMessage = ENDMESSAGE Then
                lngReply = MsgBox("The program wan unable to return you to your previous resolution." & _
                " Would you like to try 640 X 480?", vbExclamation Or vbYesNo, "Program Error")
                If lngReply = vbYes Then
                    TryAgain
                End If
                Exit Function
            End If
    End Select
End Function
