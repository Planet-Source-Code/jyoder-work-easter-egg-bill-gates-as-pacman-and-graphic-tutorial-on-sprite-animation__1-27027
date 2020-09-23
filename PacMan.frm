VERSION 5.00
Begin VB.Form frmPacMan 
   Caption         =   "PacMan"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   754
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Caption         =   "Game Options"
      ClipControls    =   0   'False
      Height          =   5415
      Left            =   6000
      TabIndex        =   19
      Top             =   720
      Width           =   6495
      Begin VB.CheckBox chkSound 
         Alignment       =   1  'Right Justify
         Caption         =   "Sound"
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame fraGraphics 
         Caption         =   "Graphics"
         ClipControls    =   0   'False
         Height          =   1575
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   2655
         Begin VB.OptionButton optGraphics 
            Alignment       =   1  'Right Justify
            Caption         =   "Normal PacMan"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optGraphics 
            Alignment       =   1  'Right Justify
            Caption         =   "Evil PacMan"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   27
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optGraphics 
            Alignment       =   1  'Right Justify
            Caption         =   "Bill Gates Graphics"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   26
            Top             =   1080
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CommandButton pbOK 
         Caption         =   "Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Frame fraMazeSize 
         Caption         =   "Maze Size"
         Height          =   2055
         Left            =   720
         TabIndex        =   20
         Top             =   2280
         Width           =   4935
         Begin VB.OptionButton optMazeSize 
            Caption         =   "Large  (Available at Resolution 1024 x 768 and higher)"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   720
            Width           =   4335
         End
         Begin VB.OptionButton optMazeSize 
            Caption         =   "Small   (Available at Resolution 800 x 600 and higher)"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.Label Label2 
            Caption         =   $"PacMan.frx":0000
            Height          =   615
            Left            =   360
            TabIndex        =   23
            Top             =   1200
            Width           =   4215
         End
      End
   End
   Begin VB.PictureBox pWallTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1800
      Picture         =   "PacMan.frx":008D
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox pPowerUpSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   960
      Picture         =   "PacMan.frx":3CCF
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pGhostScared 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3480
      Picture         =   "PacMan.frx":4911
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pPowerUp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2640
      Picture         =   "PacMan.frx":7953
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pPacSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "PacMan.frx":8595
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pPacSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "PacMan.frx":91D7
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pPelletSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      Picture         =   "PacMan.frx":9E19
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pFang 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   3960
      Picture         =   "PacMan.frx":AA5B
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pGhostSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2640
      Picture         =   "PacMan.frx":B69D
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pGhostSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   240
      Picture         =   "PacMan.frx":E6DF
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pGhostSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   2640
      Picture         =   "PacMan.frx":11721
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pGhostSub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "PacMan.frx":14763
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pFang 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   3960
      Picture         =   "PacMan.frx":177A5
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      Picture         =   "PacMan.frx":183E7
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pPellet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2040
      Picture         =   "PacMan.frx":19029
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pWall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      Picture         =   "PacMan.frx":19C6B
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pPac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   2280
      Picture         =   "PacMan.frx":1A8AD
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pPac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   2280
      Picture         =   "PacMan.frx":1B4EF
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pGhost 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   3480
      Picture         =   "PacMan.frx":1C131
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "frmPacMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************
' After including this form in one of your projects, you must somehow call
' the InitializeProgram proc on this form. One example of how to do this is to
' set one of your other form's KeyPreview property to TRUE, and paste the
' following code within that form. Then when your program is running and that
' form is displayed, just type in the required EGG_KEY from anywhere and
' PRESTO your Easter Egg is activated! But you can call this however you want
' -- just use your imagination and be somewhat sneaky!
' ************************************************************************
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Const EGG_KEY = "ILOVECOMPUTERGAMES"
'Static strKey As String
'Dim frm As frmPacMan
'    If Len(strKey) = Len(EGG_KEY) Then strKey = Mid$(strKey, 2) ' Chop off first letter
'    strKey = strKey & UCase$(Chr$(KeyCode))
'    If strKey = EGG_KEY Then
'        strKey = ""
'        Set frm = New frmPacMan
'        frm.InitializeProgram
'        Set frm = Nothing
'    End If
'End Sub


Const W = 32
Const H = 32
Const PACMAN_IMAGES = 2
Const KEY_DOWN = &H1000
Const KEY_TOGGLED = &H1
Const MOVEINCR = 8      ' Must be evenly divisible in W & H
Const SCOREBOARD_SPACING = 25

Private Enum eMaze
    Wall
    Pellet
    OpenSpace
    PowerUp
End Enum

Private Type tMaze
    GridType As eMaze
    Intersection As Boolean
End Type

Const PELLET_POINTS = 10
Const POWERUP_POINTS = 50
Const BONUS_PACMAN = 10000
Const MAX_LEVEL = 5
Const ANIMATE_WHEN_NOT_MOVING = False
Const POWERUP_MOVES = 120
Const FIRST_GHOST_WORTH = 200

Dim mbAlternateGraphics As Boolean
Dim mbSound As Boolean
Dim mbLargeMaze As Boolean
Dim mbEvilPacMan As Boolean
Dim miEatGhostWorth As Long
Dim mbUserResize As Boolean
Dim miFPS As Integer    ' Framerates per second
Dim mbPaused As Boolean

Dim Rows As Integer
Dim Columns As Integer
Dim Maze() As tMaze
Dim miPacIndex As Integer
Dim miLevel As Integer
Dim mbLevelCleared As Boolean
Dim mbExitLoop As Boolean
Dim mPacMan As tCharacter
Dim mGhost(0 To 3) As tCharacter
Dim miTotalPellets As Integer
Dim miEatenPellets As Integer
Dim miTotalPoints As Long
Dim mbDied As Boolean
Dim miLives As Integer
Dim miPowerUpMoves As Integer
Dim mbScaredBlinkToggle As Boolean
Dim miStartX As Integer
Dim miStartY As Integer


Public Enum Direction
    East = 0
    South = 1
    West = 2
    North = 3
End Enum

Private Type tCharacter
    CoorX As Long
    CoorY As Long
    OldCoorX As Long
    OldCoorY As Long
    Facing As Direction
    Moving As Boolean
    Turn As Direction
    HuntFactor As Integer
    Scared As Boolean
End Type

' HuntFactor is a value from 1 to 10, with 10 being the highest (impossible
' to shake a ghost if at level 10). It's the chance in 10 that every time a
' particular ghost reaches a maze intersection, that he moves in a direction
' to "hunt down" PacMan. Otherwise, he selects a random direction.
' The values 0 - 3 are each randomly assigned to a different ghost each game.
' Then for each of the four levels to complete the games, it increases by one.
' So by the time you reach the final level (5th), one ghost will have a
' HuntFactor of 7 on the last maze while the others will be 6, 5 & 4. And trust
' me -- by that time, you'll have noticed the AI getting progressively better
' with each level. But if a ghost is "scared" then have them choose random path.

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Sub LoadUpForm()
    DisplayHelpScreen
    If ScreenWidth < 800 Then
        MsgBox "Cannot run in resolutions under 800 x 640." & vbCrLf & vbCrLf & "By the way, why are you running in anything less than a resolution of 800 x 640!?", vbExclamation, "Too Small!"
        Unload Me
        Exit Sub
    End If
    Me.Show
    fraOptions.Left = 0
    fraOptions.Top = 0
    Me.Width = Me.ScaleX(fraOptions.Width, vbPixels, vbTwips)
    Me.Height = Me.ScaleY(fraOptions.Height + 20, vbPixels, vbTwips)
    If ScreenWidth < 1024 Then
        optMazeSize(1).Enabled = False
        optMazeSize(0).Value = True
    End If
    CenterForm
End Sub


Private Sub InitializeProgram()
Dim i As Integer
Dim iGhost As Integer
Dim iHuntFactor As Integer

    Randomize
    miFPS = 14
    
    CreateNewGhost vbGreen, 1
    CreateNewGhost vbCyan, 2
    CreateNewGhost vbMagenta, 3
    
    If mbAlternateGraphics Then
        ' Substitute alternate images before generating other images and masks
        pPac(0).Picture = pPacSub(0).Picture
        pPac(1).Picture = pPacSub(1).Picture
        For i = 0 To 3: pGhost(i).Picture = pGhostSub(i).Picture: Next i
        pPellet.Picture = pPelletSub.Picture
        pPowerUp.Picture = pPowerUpSub.Picture
    End If
    
    If mbEvilPacMan Then
        ' I switched index assignment only because when PaMan is at
        ' rest, would like to see pic with open mouthed fang instead
        pPac(0) = pFang(1)
        pPac(1) = pFang(0)
    End If
    
    ' Create east, south and north frame for PacMan
    For i = 0 To PACMAN_IMAGES - 1: CreateFrames pPac(i): Next i
    
    ' Create masks only for enemies since pre-drew them
    For i = 0 To 3: CreateMasks pGhost(i): Next i
    
    ' Create the images for when ghosts are scared in a 3rd row in each ghost image
    For i = 0 To 3: CreateScaredGhosts pGhost(i): Next i
    
    miPacIndex = 0
    miLevel = 1
    miTotalPoints = 0
    miLives = 3
    mbDied = False
    
    For iHuntFactor = 0 To 3
        Do
            ' Start with these value, different per ghost per game
            iGhost = Rand(0, 3)
        Loop Until mGhost(iGhost).HuntFactor = 0
        mGhost(iGhost).HuntFactor = iHuntFactor
    Next iHuntFactor
    
    
    Me.Show
    DoEvents
    
    ' If ever want to change the background color, use this...
    'ChangeBackgroundColor vbWhite
    
    Do
        
        RefreshScoreBoard
        
        If Not mbDied Then
            miEatenPellets = 0
            If mbLargeMaze Then LoadLargeMaze Else LoadSmallMaze
            ' Change Map Tile
            BitBlt pWall.hdc, 0, 0, W, H, pWallTiles.hdc, (miLevel - 1) * W, 0, vbSrcCopy
            Me.Width = Me.ScaleX(W * Columns, vbPixels, vbTwips) + 100
            Me.Height = Me.ScaleY(H * Rows, vbPixels, vbTwips) + 400
            CenterForm
            DoEvents
        End If
        
        ' ************** START OF INIT STUFF BEFORE EACH LEVEL ***************
        mbUserResize = True
        mbDied = False
        mbScaredBlinkToggle = False
        miPowerUpMoves = 0
        For i = 0 To 3: mGhost(i).Scared = False: Next i
        miEatGhostWorth = FIRST_GHOST_WORTH
        mPacMan.CoorX = (miStartX - 1) * W
        mPacMan.CoorY = (miStartY - 1) * H
        miPacIndex = 0
        mPacMan.Facing = East
        mPacMan.Turn = mPacMan.Facing
        mPacMan.Moving = False
        AssignGhostPositions
        mbLevelCleared = False
        DrawMaze
        DrawChar mPacMan
        For i = 0 To 3: DrawChar mGhost(i), i: Next i ' Draw their first pics
        DoEvents
        If mbSound Then
            PlaySound App.Path & "\Data1.dat", True
        Else
            TickCountPause 2000
        End If
        mPacMan.Facing = East
        mPacMan.Turn = mPacMan.Facing
        mPacMan.Moving = False
        ' ************** END OF INIT STUFF BEFORE EACH LEVEL ***************
        
        GameLoop
        
        If mbLevelCleared Then
            miLevel = miLevel + 1
            miFPS = miFPS + 1
            ' Increase HuntFactor of Ghosts
            For i = 0 To 3: mGhost(i).HuntFactor = mGhost(i).HuntFactor + 1: Next i
        End If
        
    Loop Until (miLevel > MAX_LEVEL) Or mbExitLoop Or (mbDied And miLives = 0)
    
    'If Not mbDied And mbLevelCleared Then
    If miLevel > MAX_LEVEL Then
        MsgBox "Congratulations! You defeated Pac-MANager!", , "Pac-MANager"
        ValidateTopScores
    End If
    
    ' Player lost
    If mbDied Then
        ValidateTopScores
    End If
    
    Unload Me
    
End Sub


Private Sub AssignGhostPositions()
    ' Start ghosts in the corners
    mGhost(0).CoorX = W
    mGhost(0).CoorY = H
    mGhost(0).Facing = East
    mGhost(1).CoorX = (Columns - 2) * W
    mGhost(1).CoorY = H
    mGhost(1).Facing = South
    mGhost(2).CoorX = W
    mGhost(2).CoorY = (Rows - 2) * H
    mGhost(2).Facing = North
    mGhost(3).CoorX = (Columns - 2) * W
    mGhost(3).CoorY = (Rows - 2) * H
    mGhost(3).Facing = West
End Sub


Public Sub GameLoop()
Dim i As Integer
Dim x As Direction
Dim y As Direction
Dim x2 As Direction
Dim y2 As Direction
Dim iNewDir As Direction
Dim bTurn As Boolean
Dim xDist As Integer
Dim yDist As Integer
Dim iHunt As Integer
Dim iGhostCenterX As Long, iGhostCenterY As Long
Dim iPacManCenterX As Long, iPacManCenterY As Long
Dim xPacMan As Integer, yPacMan As Integer
Dim bGridSpot As Boolean
Dim lCurrentTick As Long, lNextTick As Long
Static iBonus As Long
    
    
    If iBonus = 0 Then iBonus = 1
    lCurrentTick = GetTickCount
    
    Do
        
        While mbPaused
            lCurrentTick = GetTickCount
            DoEvents
        Wend
        If mbExitLoop Then Exit Sub ' So can close form when close on pause
        
        While Me.WindowState = vbMinimized: DoEvents: Wend
        lNextTick = lCurrentTick + (1000 / miFPS)
        
        If KeyPressed(vbKeyUp) Then
            mPacMan.Turn = North
            mPacMan.Moving = True
        End If
        If KeyPressed(vbKeyDown) Then
            mPacMan.Turn = South
            mPacMan.Moving = True
        End If
        If KeyPressed(vbKeyRight) Then
            mPacMan.Turn = East
            mPacMan.Moving = True
        End If
        If KeyPressed(vbKeyLeft) Then
            mPacMan.Turn = West
            mPacMan.Moving = True
        End If
        
        
        If miPowerUpMoves > 0 Then
            miPowerUpMoves = miPowerUpMoves - 1
            If miPowerUpMoves <= 40 Then
                If miPowerUpMoves Mod 5 = 0 Then
                    mbScaredBlinkToggle = Not mbScaredBlinkToggle
                End If
            End If
            If miPowerUpMoves = 0 Then
                For i = 0 To 3: mGhost(i).Scared = False: Next i
                mbScaredBlinkToggle = False
                miEatGhostWorth = FIRST_GHOST_WORTH
            End If
        End If
        
        
        ' Check if coordinates are at a "block" area (intersection or corner)
        If mPacMan.CoorX Mod W = 0 And mPacMan.CoorY Mod H = 0 Then
            
            x = mPacMan.CoorX / W + 1
            y = mPacMan.CoorY / H + 1
            
            ' Assess the maze array to see if can turn in desired direction
            bTurn = False
            Select Case mPacMan.Turn
                Case North
                    If Maze(x, y - 1).GridType <> Wall Then bTurn = True
                Case South
                    If Maze(x, y + 1).GridType <> Wall Then bTurn = True
                Case East
                    If Maze(x + 1, y).GridType <> Wall Then bTurn = True
                Case West
                    If Maze(x - 1, y).GridType <> Wall Then bTurn = True
            End Select
            
            If bTurn Then
                mPacMan.Facing = mPacMan.Turn
            Else
                ' Can't turn -- now check if there's a barrier in before you
                Select Case mPacMan.Facing
                    Case North
                        If Maze(x, y - 1).GridType = Wall Then mPacMan.Moving = False
                    Case South
                        If Maze(x, y + 1).GridType = Wall Then mPacMan.Moving = False
                    Case East
                        If Maze(x + 1, y).GridType = Wall Then mPacMan.Moving = False
                    Case West
                        If Maze(x - 1, y).GridType = Wall Then mPacMan.Moving = False
                End Select
            End If
        
        Else
            
            ' If NOT at block intersection, should be allowed to turn back
            ' in opposite direction before having to reach a crossroads
            If mPacMan.Turn = OppositeDir(mPacMan.Facing) Then
                mPacMan.Facing = mPacMan.Turn
                x = Int(mPacMan.CoorX / W) + 1
                y = Int(mPacMan.CoorY / H) + 1
            End If
            
        End If
        
        xPacMan = x
        yPacMan = y
        
        If mPacMan.Moving Then
            mPacMan.OldCoorX = mPacMan.CoorX
            mPacMan.OldCoorY = mPacMan.CoorY
            Select Case mPacMan.Facing
                Case North
                    mPacMan.CoorY = mPacMan.CoorY - MOVEINCR
                Case South
                    mPacMan.CoorY = mPacMan.CoorY + MOVEINCR
                Case East
                    mPacMan.CoorX = mPacMan.CoorX + MOVEINCR
                Case West
                    mPacMan.CoorX = mPacMan.CoorX - MOVEINCR
            End Select
        End If
        
        ' Loop thru ghosts
        For i = 0 To 3
            
            bGridSpot = False   ' Init
            
            ' Check if coordinates of ghost are at a "block" area
            ' and if the ghost is, then select a different path
            If mGhost(i).CoorX Mod W = 0 And mGhost(i).CoorY Mod H = 0 Then
                x = mGhost(i).CoorX / W + 1
                y = mGhost(i).CoorY / H + 1
                bGridSpot = True
            End If
            
            ' If Ghost is at a block area which is an intersection, then must
            ' now decide where to turn them. This is where our AI resides.
            If bGridSpot And (Maze(x, y).Intersection) Then
                
                ' Init new map coors to move to
                x2 = x
                y2 = y
                
                ' Randomly select iHunt between 0 and 10. If the value is greater
                ' then or equal to the ghost's HuntFactor, then head for the PacMan!!
                ' Unless "scared" if PacMan is PoweredUp, in which case make it random.
                iHunt = Rand(1, 10)
                
                If mGhost(i).HuntFactor >= iHunt And Not mGhost(i).Scared Then
                    
                    ' GO HUNTING AFTER THE PACMAN!!
                    ' This routine will take a fairly direct route toward PacMan.
                    ' If Ghost's hunt value is ever 10, then it's possible (but rare)
                    ' that the ghost will get "stuck" in a repetitive path in trying
                    ' to get to PacMan. However, if they ever get near PacMan, he'll
                    ' never shake them. So NEVER assign value of 10 since never want
                    ' ghosts to be stuck indefinitely, and always want PacMan to have
                    ' a chance at shaking them.
                    
                    xDist = xPacMan - x
                    yDist = yPacMan - y
                    
                    If Abs(xDist) > Abs(yDist) Then
                        ' More distance to PacMan in the X direction
                        ChooseLeftOrRight xDist, x, y, x2, iNewDir
                        If x2 = x Then
                            ' Not accessible, choose closer up/down direction instead
                            ChooseUpOrDown yDist, x, y, y2, iNewDir
                        End If
                    Else
                        ' More distance to PacMan in the Y direction
                        ChooseUpOrDown yDist, x, y, y2, iNewDir
                        If y2 = y Then
                            ' Not accessible, choose closer left/right direction instead
                            ChooseLeftOrRight xDist, x, y, x2, iNewDir
                        End If
                    End If
                    
                End If
                
                
                ' If Hunt was never called (or failed) then randomly select a direction
                If (x2 = x) And (y2 = y) Then
                    
                    Do
                        x2 = x
                        y2 = y
                        iNewDir = Rand(0, 3)
                        Select Case iNewDir
                            Case North: y2 = y - 1
                            Case South: y2 = y + 1
                            Case East: x2 = x + 1
                            Case West: x2 = x - 1
                        End Select
                    ' Use this loop if never want them to turn directly around
                    Loop Until Maze(x2, y2).GridType <> Wall And mGhost(i).Facing <> OppositeDir(iNewDir)
                    ' Use this loop if want to allow them to turn directly around (but looks stupid)
                    'Loop Until Maze(x2, y2) <> Wall
                    
                End If
                
                mGhost(i).Facing = iNewDir
                
            End If
            
            mGhost(i).OldCoorX = mGhost(i).CoorX
            mGhost(i).OldCoorY = mGhost(i).CoorY
            
            Select Case mGhost(i).Facing
                Case North: mGhost(i).CoorY = mGhost(i).CoorY - MOVEINCR
                Case South: mGhost(i).CoorY = mGhost(i).CoorY + MOVEINCR
                Case East: mGhost(i).CoorX = mGhost(i).CoorX + MOVEINCR
                Case West: mGhost(i).CoorX = mGhost(i).CoorX - MOVEINCR
            End Select
            
        Next i
        
        
        ' Erase chars at current locations
        EraseChar mPacMan
        For i = 0 To 3
            EraseChar mGhost(i)
        Next i
        
        
        If (miEatenPellets = miTotalPellets) Then mbLevelCleared = True
        
        
        ' Draw chars at new locations
        DrawChar mPacMan
        For i = 0 To 3
            DrawChar mGhost(i), i
        Next i
        
        
        ' Collision detection from middle of chars
        For i = 0 To 3
            iGhostCenterX = mGhost(i).CoorX + W \ 2
            iGhostCenterY = mGhost(i).CoorY + H \ 2
            iPacManCenterX = mPacMan.CoorX + W \ 2
            iPacManCenterY = mPacMan.CoorY + H \ 2
            If Diff(iGhostCenterX, iPacManCenterX) <= W \ 2 And _
                Diff(iGhostCenterY, iPacManCenterY) <= H \ 2 Then
                If mGhost(i).Scared Then
                    ' PacMan eats ghost
                    miTotalPoints = miTotalPoints + miEatGhostWorth
                    If mbSound Then
                        PlaySound App.Path & "\Data4.dat", True
                    Else
                        TickCountPause 600
                    End If
                    miEatGhostWorth = miEatGhostWorth * 2
                    EraseChar mGhost(i)
                    SendToFarthestCorner mGhost(i)
                Else
                    ' Ghost kills PacMan
                    If Not mbDied Then  ' This is so two ghosts at same time don't take off two lives
                        miLives = miLives - 1
                        mbDied = True
                    End If
                End If
            End If
        Next i
        
        If miTotalPoints > BONUS_PACMAN * iBonus Then
            miLives = miLives + 1
            iBonus = iBonus + 1
            If mbSound Then PlaySound App.Path & "\Data5.dat", True
        End If
        RefreshScoreBoard
        
        ' Play dying sound
        If mbDied Then
            If mbSound Then
                PlaySound App.Path & "\Data3.dat", True
            Else
                If miLives > 0 Then TickCountPause 2000
            End If
        End If
        
        ' Pause for so many milliseconds (using framerate) otherwise game too fast
        lCurrentTick = GetTickCount
        While lCurrentTick < lNextTick
            lCurrentTick = GetTickCount
            DoEvents
        Wend
        DoEvents
        
    Loop Until mbExitLoop Or mbLevelCleared Or mbDied
    
End Sub


Private Function KeyPressed(iKeyCode As Integer) As Boolean
    KeyPressed = GetKeyState(iKeyCode) And KEY_DOWN
End Function


Private Function Diff(v1 As Long, v2 As Long) As Long
    Diff = Abs(v1 - v2)
End Function


Private Function Dist(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Double
    Dist = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function


Private Sub SendToFarthestCorner(mChar As tCharacter)
Dim lMax As Double
Dim lDistNW As Double, lDistNE As Double
Dim lDistSW As Double, lDistSE As Double
    
    lDistNE = Dist(mChar.CoorX, mChar.CoorY, W, H)
    lDistNW = Dist(mChar.CoorX, mChar.CoorY, W * (Columns - 2), H)
    lDistSE = Dist(mChar.CoorX, mChar.CoorY, W, H * (Rows - 2))
    lDistSW = Dist(mChar.CoorX, mChar.CoorY, W * (Columns - 2), H * (Rows - 2))
    lMax = Max(lDistNW, lDistNE)
    lMax = Max(lMax, lDistSW)
    lMax = Max(lMax, lDistSE)
    Select Case lMax
        Case lDistNE
            mChar.CoorX = W
            mChar.CoorY = H
        Case lDistNW
            mChar.CoorX = W * (Columns - 2)
            mChar.CoorY = H
        Case lDistSE
            mChar.CoorX = W
            mChar.CoorY = H * (Rows - 2)
        Case lDistSW
            mChar.CoorX = W * (Columns - 2)
            mChar.CoorY = H * (Rows - 2)
    End Select
    mChar.OldCoorX = mChar.CoorX
    mChar.OldCoorY = mChar.CoorY
    mChar.Scared = False    ' No longer scared
End Sub


Private Function Max(v1 As Double, v2 As Double) As Double
    Max = IIf(v1 >= v2, v1, v2)
End Function


Private Sub DrawChar(mChar As tCharacter, Optional Index As Integer = -1)
Dim x As Integer, y As Integer, i As Integer
Dim bPacMan As Boolean
Dim SourceX As Long
    
    If Index = -1 Then bPacMan = True
    
    x = Int(mChar.CoorX / W) + 1
    y = Int(mChar.CoorY / H) + 1
    
    If mChar.CoorX Mod W = 0 And mChar.CoorY Mod H = 0 Then
        If bPacMan Then
            Select Case Maze(x, y).GridType
                Case Pellet
                    Maze(x, y).GridType = OpenSpace
                    miEatenPellets = miEatenPellets + 1
                    miTotalPoints = miTotalPoints + PELLET_POINTS
                    If mbSound Then PlaySound App.Path & "\Data2.dat", False
                Case PowerUp
                    Maze(x, y).GridType = OpenSpace
                    miEatenPellets = miEatenPellets + 1
                    miTotalPoints = miTotalPoints + POWERUP_POINTS
                    miPowerUpMoves = POWERUP_MOVES
                    For i = 0 To 3: mGhost(i).Scared = True: Next i
                    mbScaredBlinkToggle = False ' Reset
            End Select
        End If
    End If
    
    ' First paint the mask with vbSrcAnd, then paint the pic with vbSrcPaint
    SourceX = W * mChar.Facing
    
    ' Draw the new coors for the character
    If bPacMan Then
        BitBlt Me.hdc, mChar.CoorX, mChar.CoorY, W, H, pPac(miPacIndex).hdc, SourceX, H, vbSrcAnd
        BitBlt Me.hdc, mChar.CoorX, mChar.CoorY, W, H, pPac(miPacIndex).hdc, SourceX, 0, vbSrcPaint
    Else
        BitBlt Me.hdc, mChar.CoorX, mChar.CoorY, W, H, pGhost(Index).hdc, SourceX, H, vbSrcAnd
        If Not mChar.Scared Or mbScaredBlinkToggle Then
            ' Normal pic of ghost -- used when not scared or if scared part if flickering
            BitBlt Me.hdc, mChar.CoorX, mChar.CoorY, W, H, pGhost(Index).hdc, SourceX, 0, vbSrcPaint
        Else
            ' Scared pic of ghost
            BitBlt Me.hdc, mChar.CoorX, mChar.CoorY, W, H, pGhost(Index).hdc, SourceX, H * 2, vbSrcPaint
        End If
    End If
    
    ' Change the index of the pic to use for next time
    If bPacMan And (mPacMan.Moving Or ANIMATE_WHEN_NOT_MOVING) Then
        miPacIndex = miPacIndex + 1
        If miPacIndex = PACMAN_IMAGES Then miPacIndex = 0
    End If
    
    If Me.AutoRedraw Then Me.Refresh
    
End Sub


' Returns both new map coor (x2) and New Direction (iNewDir)
Private Sub ChooseLeftOrRight(xDist As Integer, x As Direction, y As Direction, _
                                ByRef x2 As Direction, ByRef iNewDir As Direction)
    If xDist > 0 Then
        ' PacMan is to the right
        If Maze(x + 1, y).GridType <> Wall Then x2 = x2 + 1
        iNewDir = East
    Else
        ' PacMan is to the left
        If Maze(x - 1, y).GridType <> Wall Then x2 = x2 - 1
        iNewDir = West
    End If
End Sub


' Returns both new map coor (y2) and New Direction (iNewDir)
Private Sub ChooseUpOrDown(yDist As Integer, x As Direction, y As Direction, _
                                ByRef y2 As Direction, ByRef iNewDir As Direction)
    If yDist > 0 Then
        ' PacMan is to the South
        If Maze(x, y + 1).GridType <> Wall Then y2 = y + 1
        iNewDir = South
    Else
        ' PacMan is to the North
        If Maze(x, y - 1).GridType <> Wall Then y2 = y - 1
        iNewDir = North
    End If
End Sub


Private Function OppositeDir(iDir As Direction) As Direction
    Select Case iDir
        Case North
            OppositeDir = South
        Case South
            OppositeDir = North
        Case East
            OppositeDir = West
        Case West
            OppositeDir = East
    End Select
End Function


Private Function PaintGrid(x As Integer, y As Integer)
    Select Case Maze(x, y).GridType
        Case Wall
            BitBlt Me.hdc, (x - 1) * W, (y - 1) * H, W, H, pWall.hdc, 0, 0, vbSrcCopy
        Case Pellet
            BitBlt Me.hdc, (x - 1) * W, (y - 1) * H, W, H, pPellet.hdc, 0, 0, vbSrcCopy
        Case OpenSpace
            BitBlt Me.hdc, (x - 1) * W, (y - 1) * H, W, H, pBlank.hdc, 0, 0, vbSrcCopy
        Case PowerUp
            BitBlt Me.hdc, (x - 1) * W, (y - 1) * H, W, H, pPowerUp.hdc, 0, 0, vbSrcCopy
    End Select
End Function


Private Function EraseChar(mChar As tCharacter)
Dim x As Integer, y As Integer
Dim ToX As Integer, ToY As Integer
Dim FromX As Integer, FromY As Integer
    
    x = Int(mChar.CoorX / W) + 1
    y = Int(mChar.CoorY / H) + 1
    ToX = x
    ToY = y
    
    ' Check if coordinates are at a "block" area (intersection or corner)
    If mChar.CoorX Mod W = 0 And mChar.CoorY Mod H = 0 Then
        ' Do nothing if they are
    Else
        If mChar.Facing = East Then
            ToX = ToX + 1
        ElseIf mChar.Facing = South Then
            ToY = ToY + 1
        End If
    End If
    
    ' Get direction opposite from where facing and check if pellet is behind
    ' you. If so, then repaint pellet since may have partially BitBlted over.
    ' (This will happen when you initially turn around, but can also happen
    ' after you've already turned around which is why must check for it here).
    FromX = ToX ' Init
    FromY = ToY ' Init
    Select Case mChar.Facing
        Case North: FromY = ToY + 1
        Case South: FromY = ToY - 1
        Case East: FromX = ToX - 1
        Case West: FromX = ToX + 1
    End Select
    
    ' Repaint To and From squares
    PaintGrid ToX, ToY
    PaintGrid FromX, FromY
    
End Function


Private Sub CreateScaredGhosts(p As PictureBox)
Dim x As Integer, y As Integer
Dim lBackgroundColor As Long
Dim lColor As Long
    
    p.Height = H * 3
    
    If Not mbAlternateGraphics Then
        ' For just normal ghosts, copy over pre-made image
        BitBlt p.hdc, 0, H * 2, W * 4, H, pGhostScared.hdc, 0, 0, vbSrcCopy
    Else
        ' The alternate pics use their own inverse
        BitBlt p.hdc, 0, H * 2, W * 4, H, p.hdc, 0, 0, vbNotSrcCopy
    End If
    
    lBackgroundColor = GetPixel(p.hdc, 0, H * 2)    ' Assume for our sprites that background color is at 0,0
    For x = 0 To (W * 4) - 1: For y = H * 2 To (H * 3) - 1
        lColor = GetPixel(p.hdc, x, y)
        If lColor = lBackgroundColor Then
            SetPixel p.hdc, x, y, vbBlack
        End If
    Next y: Next x
    
End Sub


Private Sub CreateNewGhost(lNewColor As Long, iIndex As Integer)
Dim x As Long, y As Long
Dim lColor As Long, lGhostColor As Long
    
    Load pGhost(iIndex)
    With pGhost(iIndex)
        .Picture = pGhost(iIndex - 1)
        .Left = pGhost(iIndex - 1).Left
        .Top = pGhost(iIndex - 1).Top + H * 3
        lGhostColor = GetPixel(.hdc, 6, 6)  ' The known coors of color to change
        For x = 0 To .ScaleWidth - 1
            For y = 0 To H - 1
                lColor = GetPixel(.hdc, x, y)
                If lColor = lGhostColor Then SetPixel .hdc, x, y, lNewColor
            Next y
        Next x
        .Visible = pGhost(iIndex - 1).Visible
    End With
    
End Sub


Private Sub CreateFrames(p As PictureBox)
Dim x As Long
Dim y As Long
Dim lColor As Long
Dim xStart As Long
Dim yStart As Long
Dim yCoor As Long
    
    p.Width = W * 4
    
    ' Draw down face (90 degrees from right face)
    xStart = W
    yStart = 0
    
    For x = 0 To W - 1
        yCoor = H
        For y = 0 To H - 1
            lColor = GetPixel(p.hdc, x, y)
            yCoor = yCoor - 1
            SetPixel p.hdc, xStart + yCoor, yStart + x, lColor
        Next y
    Next x
    
    ' Draw left face (opposite of left)
    StretchBlt p.hdc, W * 2, 0, W, H, p.hdc, W - 1, 0, -W, H, vbSrcCopy
    
    ' Draw up face (flipped horizontally and then vertically of down)
    StretchBlt p.hdc, W * 3, 0, W, H, p.hdc, (W * 2) - 1, 0, -W, H, vbSrcCopy
    StretchBlt p.hdc, W * 3, 0, W, H, p.hdc, (W * 3), H - 1, W, -H, vbSrcCopy
    
    ' Create the masks beneath it in same pic
    CreateMasks p
    
    p.Refresh
    
End Sub


Private Sub CreateMasks(p As PictureBox)
Dim x As Long
Dim y As Long
Dim lColor As Long
Dim lBackgroundColor As Long

    p.Height = H * 2
    lBackgroundColor = GetPixel(p.hdc, 0, 0)    ' Assume for our sprites that background color is at 0,0
    
    For x = 0 To (W * 4) - 1
        For y = 0 To H - 1
            lColor = GetPixel(p.hdc, x, y)
            If lColor <> lBackgroundColor Then
                ' Create mask in second row
                SetPixel p.hdc, x, y + H, vbBlack
            Else
                ' Black out white background in main (first) row
                SetPixel p.hdc, x, y, vbBlack
            End If
        Next y
    Next x
    
End Sub


Private Sub LoadSmallMaze()
Dim m(1 To 18) As String
     m(1) = "XXXXXXXXXXXXXXXXXXXXXXXX"
     m(2) = "X.....X...X............X"
     m(3) = "X.XXX.X.X.X.XXX.X.X.XX.X"
     m(4) = "X..P..X.X.....X.X.XP...X"
     m(5) = "X.X.X.....XXX.X...X.XX.X"
     m(6) = "X.X.X.XXX...X.X.X....X.X"
     m(7) = "X.X...X...X.....X.XX.X.X"
     m(8) = "X.X.X.X.XXXXX.XXX......X"
     m(9) = "X...X....... .X...XX.X.X"
    m(10) = "XXX.X.XXX.X.X.X.XXXX.X.X"
    m(11) = "X.......X.X.X........X.X"
    m(12) = "X.XXXXX.X...X.XXX.X.XX.X"
    m(13) = "X....X....X.......X....X"
    m(14) = "X.XX.X.XX.XXX.XXX.X.XXXX"
    m(15) = "X...P..X....X.....P....X"
    m(16) = "X.XXXXXX.XX.XX.X.XXX.X.X"
    m(17) = "X..............X.......X"
    m(18) = "XXXXXXXXXXXXXXXXXXXXXXXX"
    
    ReadMaze m
    
End Sub


Private Sub LoadLargeMaze()
Dim m(1 To 22) As String
    
     m(1) = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
     m(2) = "X.....X...X........X..........X"
     m(3) = "X.XXX.X.X.X.XXXX.X...XXX.X.XX.X"
     m(4) = "X.....X.X......X.X.X.X...X....X"
     m(5) = "X.X.X..PX.XX.X.X.X.X.X.X.X.XX.X"
     m(6) = "X.X.X.X......X.X...X...XP...X.X"
     m(7) = "X.X.X.XXX.X.X....X.XXX.X.XX.X.X"
     m(8) = "X.X...X...X.X.XXXX...X.X......X"
     m(9) = "X.X.X.X.XXX.X......X.....XX.X.X"
    m(10) = "X...X.....X.X.X.XX.XX.XXXXX.X.X"
    m(11) = "X.XXX.XXX.X...X ...X..........X"
    m(12) = "X.......X.XXX.X.XXXX.XXX.X.XXXX"
    m(13) = "XXX.XXX.X.....X......X...X....X"
    m(14) = "X.......X.XXX.XXXX.XXX.XXX.XX.X"
    m(15) = "X.XXXXX...X..........X......X.X"
    m(16) = "X.....X.XXX.XX.X.XXX...XXXX...X"
    m(17) = "X.X.X..P.......X.X...X...P..X.X"
    m(18) = "X.X.XXX.X.XXXX.X.X.X.XXXX.XXX.X"
    m(19) = "X.X...X.X.X......X.X........X.X"
    m(20) = "X.XXX.X.X.X.XXXX.X.XXX.XXXX.X.X"
    m(21) = "X.........X...................X"
    m(22) = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    
    ReadMaze m
    
End Sub


Private Sub ReadMaze(m() As String)
Dim x As Integer, y As Integer
Dim Char As String
Dim bEastWest As Boolean, bNorthSouth As Boolean
    
    Rows = UBound(m)
    Columns = Len(m(1))
    ReDim Maze(1 To Columns, 1 To Rows)
    
    miTotalPellets = 0
    For y = 1 To Rows: For x = 1 To Columns
        Char = Mid$(m(y), x, 1)
        Select Case Char
            Case "X"
                Maze(x, y).GridType = Wall
            Case "."
                Maze(x, y).GridType = Pellet
                miTotalPellets = miTotalPellets + 1
            Case " "
                Maze(x, y).GridType = OpenSpace
            Case "P"
                Maze(x, y).GridType = PowerUp
                miTotalPellets = miTotalPellets + 1
        End Select
    Next x: Next y
    
    ' Assign intersection values as needed (ignore edges since assumed always blocked)
    For y = 2 To Rows - 1: For x = 2 To Columns - 1
        If (Maze(x, y).GridType <> Wall) Then
            bEastWest = False: bNorthSouth = False
            If (Maze(x + 1, y).GridType <> Wall) Or (Maze(x - 1, y).GridType <> Wall) Then
                ' There an open grid east, west, or both
                bEastWest = True
            End If
            If (Maze(x, y + 1).GridType <> Wall) Or (Maze(x, y - 1).GridType <> Wall) Then
                ' There an open grid south, north, or both
                bNorthSouth = True
            End If
            If bEastWest And bNorthSouth Then Maze(x, y).Intersection = True
        End If
    Next x: Next y
    
    AssignPacPos
    
End Sub


Private Sub AssignPacPos()
    Dim y As Integer, x As Integer
    For y = 1 To Rows: For x = 1 To Columns
        If Maze(x, y).GridType = OpenSpace Then
            ' This is where to start PacMan
            miStartX = x
            miStartY = y
        End If
    Next x: Next y
End Sub


Private Function ScreenWidth() As Long
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
End Function


Private Sub DrawMaze()
Dim x As Integer, y As Integer
    Me.BackColor = vbBlack
    For y = 1 To Rows
        For x = 1 To Columns
            PaintGrid x, y
        Next x
    Next y
    If Me.AutoRedraw Then Me.Refresh
    DoEvents
End Sub


Private Sub CenterForm()
    On Error Resume Next
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub


Private Function PlaySound(ByRef szFileToPlay As String, Optional bolPauseUntilDone As Boolean = True) As Long
On Error Resume Next    ' Don't want sound to mess us up
Dim iFlag As Integer
    If ("" = Dir(szFileToPlay)) Then Exit Function
    iFlag = IIf(bolPauseUntilDone, 2, 1)
    PlaySound = sndPlaySound(szFileToPlay, iFlag)
End Function


Public Function Rand(lLowerBound As Long, lUpperBound As Long) As Long
    Rand = Int((lUpperBound - lLowerBound + 1) * Rnd + lLowerBound)
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If miFPS = 0 Then Exit Sub  ' Game hasn't started yet
    
    ' If hit pause or "P", pause game -- any other except F1, F3, and Escape unpauses
    If (KeyCode = vbKeyPause Or KeyCode = vbKeyP) Then
        mbPaused = Not mbPaused
    ElseIf KeyCode = vbKeyF1 Then
        DisplayHelpScreen
    ElseIf KeyCode = vbKeyF3 Then
        DisplayHighScores
    ElseIf KeyCode = vbKeyEscape Then
        mbPaused = False
        mbExitLoop = True
    Else
        mbPaused = False
    End If
    
End Sub


Private Sub Form_Paint()
Dim i As Integer
    ' Since form has AutoRedraw set to True, we redraw Maze if Paint event is
    ' called. (If another window overlapped it, or is restored from being minimized)
    DrawMaze
    If mbPaused Then
        DrawChar mPacMan
        For i = 0 To 3: DrawChar mGhost(i), i: Next i
    End If
    DoEvents
End Sub

Private Sub Form_Resize()
    If Not mbUserResize Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.Width = Me.ScaleX(W * Columns, vbPixels, vbTwips) + 100
    Me.Height = Me.ScaleY(H * Rows, vbPixels, vbTwips) + 400
    CenterForm
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mbExitLoop = True
    mbPaused = False
End Sub


Private Sub ChangeBackgroundColor(lColor As Long)
    Me.BackColor = lColor
    ChangePicBGColor pBlank, lColor
    ChangePicBGColor pPellet, lColor
    ChangePicBGColor pPowerUp, lColor
    DrawMaze
End Sub


Private Sub ChangePicBGColor(p As PictureBox, lColor As Long)
Dim x As Integer, y As Integer
Dim lOldBG As Long
    lOldBG = GetPixel(p.hdc, 0, 0)
    For x = 0 To p.Width - 1: For y = 0 To p.Height - 1
        If GetPixel(p.hdc, x, y) = lOldBG Then
            SetPixel p.hdc, x, y, lColor
        End If
    Next y: Next x
    If p.AutoRedraw Then p.Refresh
End Sub


Private Sub DisplayHelpScreen()
Dim s As String
    
    s = s & "Welcome to Pac-MANager Version " & App.Major & "." & App.Minor & " by Jeremy Yoder"
    s = s & vbCrLf & vbCrLf & "All graphics and code (except for sound files named DataX.dat) were encapsulated into one" & vbCrLf & "Form in order to make this easy to incorporate as an Easter Egg into your own work projects."
    s = s & vbCrLf & vbCrLf & "You'll notice on the startup screen the option to choose ""Alternate Graphics"" are for you to" & vbCrLf & "customize to tell your own work story/frustrations about your managers and co-workers. (For" & vbCrLf & "this upload to PSC, I used Bill Gates being chased by Apple, Linux, Netscape and RedHat)."
    
    s = s & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbTab & vbTab & "This Help Screen" & vbTab & vbTab & "F1" & vbCrLf & vbTab & vbTab & "Move PacMan" & vbTab & vbTab & "Arrow Keys" & vbCrLf & vbTab & vbTab & "Quit" & vbTab & vbTab & vbTab & "Escape" & vbCrLf & vbTab & vbTab & "Pause" & vbTab & vbTab & vbTab & "Pause or P" & vbCrLf & vbTab & vbTab & "View High Scores" & vbTab & vbTab & "F3" & vbCrLf & vbCrLf
    MsgBox s, , "Pac-MANager"
    
End Sub


Private Sub RefreshScoreBoard()
    Me.Caption = "Pac-MANager Version " & App.Major & "." & App.Minor & Space$(SCOREBOARD_SPACING) & "Lives: " & miLives & Space$(SCOREBOARD_SPACING) & "Level " & miLevel & " of " & MAX_LEVEL & Space$(SCOREBOARD_SPACING) & "Score: " & Format$(miTotalPoints, "###,###,###,##0")
End Sub


Private Sub pbOK_Click()
    fraOptions.Visible = False
    If chkSound.Value = vbChecked Then mbSound = True
    If optGraphics(1).Value Then mbEvilPacMan = True
    If optGraphics(2).Value Then mbAlternateGraphics = True
    If optMazeSize(1).Value Then mbLargeMaze = True
    InitializeProgram
End Sub


Private Sub ReadTopScores(ByRef aScores() As Long, ByRef aNames() As String)
Dim i As Integer
    For i = 1 To 10
        aScores(i) = GetSetting("Pac-MANager", "Top Scores", "Score " & i, 0)
        aNames(i) = GetSetting("Pac-MANager", "Top Names", "Name " & i, "")
    Next i
End Sub


Private Sub ValidateTopScores()
Dim i As Integer, j As Integer
Dim strName As String
Dim aScores(1 To 10) As Long
Dim aNames(1 To 10) As String
Dim k As Integer

    ReadTopScores aScores, aNames
    If miTotalPoints > aScores(10) Then
        strName = InputBox$("You made the Top 10 Scoring List! Enter your name for recording...", "Pac-MANager", "")
        strName = Trim$(strName)
        For i = 1 To 10
            If miTotalPoints > aScores(i) Then
                ' Bump the rest down
                For j = 10 To i + 1 Step -1
                    aNames(j) = aNames(j - 1)
                    aScores(j) = aScores(j - 1)
                Next j
                aNames(i) = strName
                aScores(i) = miTotalPoints
                
                ' Write the new scores
                For k = 1 To 10
                    SaveSetting "Pac-MANager", "Top Scores", "Score " & k, aScores(k)
                    SaveSetting "Pac-MANager", "Top Names", "Name " & k, aNames(k)
                Next k
                
                ReadTopScores aScores, aNames
                DisplayHighScores
                Exit For
            End If
        Next i
    End If
End Sub


Private Sub DisplayHighScores()
Dim i As Integer
Dim aScores(1 To 10) As Long
Dim aNames(1 To 10) As String
Dim str As String
    ReadTopScores aScores, aNames
    If aScores(1) <> 0 Then
        For i = 1 To 10
            str = str & i & "." & vbTab & vbTab & Format$(aScores(i), "###,###,###,##0") & vbTab & vbTab & aNames(i) & vbCrLf
        Next i
        MsgBox str, , "Top 10 Scores"
    Else
        MsgBox "There are no High Scores at this time.", , "Pac-MANager"
    End If
End Sub


' NOTE: 1000 Milliseconds = 1 second
Public Sub TickCountPause(lMilliseconds As Long)
Dim lStart As Long
    lStart = GetTickCount
    While (GetTickCount - lStart < lMilliseconds): DoEvents: Wend
End Sub
