VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFullScreen 
      Caption         =   "Full Screen"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.HScrollBar hscrDelay 
      Height          =   255
      LargeChange     =   10
      Left            =   1680
      Max             =   1000
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Windows 32 API Version"
      Height          =   615
      Left            =   3720
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "1. View ReadMe.txt"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "  Esc = end game"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "  Arrows = Movement"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "  Z = Jump"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "  X = Attack"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Build: 07/30/03"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "3. Keys: "
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "No Rhapsody for The Weak"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   11
      Top             =   435
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "2. Use 16-bit Color"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "CASTLEVANIA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   960
      TabIndex        =   8
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Delay:"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblDelay 
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "1000 (or 1 second)"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Note: The less the delay, the faster the game will run."
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Redraw Delay"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mintDelay As Integer

Private Sub cmdStart_Click()
    DELAY = mintDelay
    Me.Hide
    If chkFullScreen.Value = Checked Then frmChangeRez.StartRez
    Load frmMain
    Load frmCollision
    frmMain.GameLoop
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Show
    mintDelay = 35
    hscrDelay.Value = mintDelay
    lblDelay.Caption = mintDelay
End Sub

Private Sub hscrDelay_Change()
    mintDelay = hscrDelay.Value
    lblDelay.Caption = mintDelay
End Sub
