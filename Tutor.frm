VERSION 5.00
Begin VB.Form frmTutor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pGhosts 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   7440
      Picture         =   "Tutor.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   240
      Width           =   8415
   End
   Begin VB.PictureBox pBackgrnd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   4320
      Picture         =   "Tutor.frx":3042
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox pMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   7560
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox pFang 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8880
      Picture         =   "Tutor.frx":6F84
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pNormal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      Picture         =   "Tutor.frx":7BC6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1200
      Picture         =   "Tutor.frx":8808
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame fraDummy 
      Caption         =   "Dummy Images"
      Height          =   2055
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   15
         Left            =   3600
         Picture         =   "Tutor.frx":944A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   21
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   14
         Left            =   4200
         Picture         =   "Tutor.frx":A08C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   13
         Left            =   4800
         Picture         =   "Tutor.frx":ACCE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   19
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   12
         Left            =   5400
         Picture         =   "Tutor.frx":B910
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   11
         Left            =   3600
         Picture         =   "Tutor.frx":C552
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   4200
         Picture         =   "Tutor.frx":D194
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   16
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   4800
         Picture         =   "Tutor.frx":DDD6
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   5400
         Picture         =   "Tutor.frx":EA18
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   2040
         Picture         =   "Tutor.frx":F65A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   1440
         Picture         =   "Tutor.frx":1029C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   840
         Picture         =   "Tutor.frx":10EDE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   240
         Picture         =   "Tutor.frx":11B20
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   2040
         Picture         =   "Tutor.frx":12762
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1440
         Picture         =   "Tutor.frx":133A4
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   840
         Picture         =   "Tutor.frx":13FE6
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "Tutor.frx":14C28
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Opening Mouths"
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Closing Mouths"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
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
      Left            =   600
      Picture         =   "Tutor.frx":1586A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   4200
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
      Left            =   3120
      Picture         =   "Tutor.frx":164AC
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton pbNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmTutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const W = 32
Const H = 32

Private Enum Direction
    East = 0
    South = 1
    West = 2
    North = 3
End Enum

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Tutor(str As String, Optional x As Long = -1, Optional y As Long = -1)
    txtInfo.Text = str
End Sub


Private Sub PointTo(ctl As Control)
    If TypeOf ctl Is Frame Then
        Me.Line (txtInfo.Width / 2, txtInfo.Top + txtInfo.Height + 10)- _
            (ctl.Left + (ctl.Width / 2), ctl.Top - 10)
    Else
        Me.Line (txtInfo.Width / 2, txtInfo.Top + txtInfo.Height + 10)- _
                (ctl.Left + (ctl.ScaleWidth / 2), ctl.Top - 10)
    End If
    ctl.Visible = True
End Sub


Private Sub Form_Load()
    Me.DrawWidth = 2
    Me.ForeColor = vbMagenta
    txtInfo.BackColor = Me.BackColor
    pbNext_Click
End Sub


Private Sub CreateNewGhost(lNewColor As Long, iNextIndex As Integer)
Dim x As Long, y As Long
Dim lColor As Long, lGhostColor As Long

    Load pGhosts(iNextIndex)
    With pGhosts(iNextIndex)
        .Picture = pGhosts(iNextIndex - 1)
        .Left = pGhosts(iNextIndex - 1).Left
        .Top = pGhosts(iNextIndex - 1).Top + pGhosts(iNextIndex - 1).ScaleHeight * 3
        lGhostColor = GetPixel(.hdc, 6, 6)  ' The known coors of color to change
        For x = 0 To .ScaleWidth - 1
            For y = 0 To H - 1
                lColor = GetPixel(.hdc, x, y)
                If lColor = lGhostColor Then SetPixel .hdc, x, y, lNewColor
            Next y
        Next x
        .Visible = True
    End With
    iNextIndex = iNextIndex + 1
    
End Sub


Private Sub CreatePacManFrames(p As PictureBox)
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
    
    p.Refresh
    
End Sub


Private Sub CreateMasks(p As PictureBox)
Dim x As Long
Dim y As Long
Dim lColor As Long
Dim lBackgroundColor As Long

    p.Height = H * 2
    ' Assume for our sprites that background color is at 0,0
    lBackgroundColor = GetPixel(p.hdc, 0, 0)
    
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


Private Sub MakeInvisIfExists(p As PictureBox)
On Error Resume Next
    p.Visible = False
End Sub


Private Sub pbNext_Click()
Static iStep As Integer
Dim x As Long, y As Long
Dim str As String

' Init by making controls invisible
Me.Cls
pNormal.Visible = False
pMask.Visible = False
pPac(0).Visible = False
pPac(1).Visible = False
pFang.Visible = False
pGhosts(0).Visible = False
fraDummy.Visible = False
pBackgrnd.Visible = False
pMap.Visible = False
MakeInvisIfExists pGhosts(1)

iStep = iStep + 1: x = -1: y = -1
'If iStep = 1 Then iStep = 50

Select Case iStep
Case 1
    'str = "This Graphical Tutorial will take you step by step thru the basics of BitBlt sprite animation. While apps based on DirectX are faster than BitBlt, it's good to cut your teeth on understanding the concepts presented in this tutorial regardless of which technique you use. (Click the next button to begin and to contine on to each next step)." & vbCrLf & vbCrLf & "NOTE: I'm assuming you know what BitBlt and API's are. If not, learn those first before continuing!"
    str = "This Graphical Tutorial will take you step by step thru the basics of BitBlt sprite animation. While apps based on DirectX are faster than BitBlt, it's good to cut your teeth on understanding the concepts presented in this tutorial regardless of which technique you use. (Click the next button to begin and to contine on to each next step)." & vbCrLf & vbCrLf & "NOTE: I'm assuming you know what BitBlt and API's are. If not, learn those first before continuing!"
    x = 9: y = 10
Case 2
    str = "First things first: When dealing with graphics in VB, set the ScaleMode properties of your Forms and PictureBoxes to Pixels, and their AutoRedraw properties to true. (You don't have in your own projects since there are purposes to the other settings which I'll talk about at the end, but for simplicity, this is what we'll do)." & vbCrLf & vbCrLf & "Next, include some commonly used graphic API's in your project, such as BitBlt & StretchBlt (both for painting entire images or parts of images), GetPixel & SetPixel (for pixel by pixel manipulation) as are needed." & vbCrLf & vbCrLf & "You may need some, all, or more depending on your needs. See the code within this form for the declaractions used in this tutorial."
    iStep = iStep + 1
Case 4
    str = "Where to start? Well, when dealing with images, you first need your images! And you'll want to create/load them all on your form (as I've done in this tutorial) or in memory once your program starts." & vbCrLf & vbCrLf & "Note that you can keep them stored out in a file, but once your program starts, it's best to load them all into Pictureboxes or memory for speed purposes when drawing later. (What do you think commercial PC games are doing while you're waiting for them to begin!?)"
Case 5
    str = "For our tutorial, we'll start with the PacMan image in a PictureBox named ""pNormal""." & vbCrLf & vbCrLf & "Note that the actual image has a BLACK background which is where we'll want images we ""paint over top of"" to remain intact. (Like a background)." ', and the MASK portion is a blackened image $$$ However, we need But we'll need "
    PointTo pNormal
Case 6
    str = "However, in order to ""paint over top of"" other images (like backgrounds) we also need to have a MASK of our image." & vbCrLf & vbCrLf & "We'll get to why and how in a bit, but for now, note that the MASK of an image is the same size, and has the actual image BLACKENED OUT, with its background set to WHITE. The name of this Picturebox is ""pMask""." '& vbCrLf & vbCrLf & "Note that the actual image has a BLACK background, and the MASK portion is a blackened image $$$ However, we need But we'll need "
    'Set ctl = pPac(0)
    PointTo pNormal
    PointTo pMask
Case 7
    str = "Next, we need to have a background of some sort." & vbCrLf & vbCrLf & "How you work your background bitmap(s) is up to you. It can be it's own bitmap (as we'll do) or if you're doing a tile map, it can be several small bitmaps (tiles) which you'll refresh as needed on your map (as I do in my Pac-Manager game)."
    PointTo pBackgrnd
Case 8
    str = "What map am I talking about!? THIS map." & vbCrLf & vbCrLf & "Think of your Map as your Painting Canvas that the end user sees and interacts with. It can be a PictureBox (like we're doing) or the Form itself like I did with my PacMan game. This is where you ""paste"" the background, and refresh it with images as needed, whether they be the background itself, or our sprites."
    PointTo pMap
Case 9
    str = "OK, enough talk. Down to the nitty gritty. First, we paint our required background onto the Map. For us at this point, that'd be the entire background..." & vbCrLf & vbCrLf & "BitBlt pMap.hdc, 0, 0, pMap.Width, pMap.Height, pBackgrnd.hdc, 0, 0, vbSrcCopy" & vbCrLf & vbCrLf & "Notice the last parameter (the ""paint type"" or ""paint mode"") is simply a direct copy when you use vbSrcCopy." '& vbCrLf & vbCrLf & "Let's break that down..."
    PointTo pBackgrnd
    PointTo pMap
    BitBlt pMap.hdc, 0, 0, pMap.Width, pMap.Height, pBackgrnd.hdc, 0, 0, vbSrcCopy
Case 10
    str = "Next, select where the Sprite should be displayed. We'll say coordinates (10,10). We now BitBlt the desired image's MASK with the following..." & vbCrLf & vbCrLf & "BitBlt pMap.hdc, 10, 10, pMask.Width, pMask.Height, pMask.hdc, 0, 0, vbSrcAnd" & vbCrLf & vbCrLf & "1) Instead of using vbSrcCopy, we used vbSrcAnd as our paint type with our MASK" & vbCrLf & "2) The result is that the WHITE portions of our picture did NOT overwrite the background!"  '& vbCrLf & vbCrLf & "Click NEXT to understand what just happened..."
    PointTo pMap
    PointTo pMask
    BitBlt pMap.hdc, 10, 10, pMask.Width, pMask.Height, pMask.hdc, 0, 0, vbSrcAnd
    iStep = iStep + 1
'Case 11
'    str = "BitBlt pMap.hdc, 10, 10, pMask.Width, pMask.Height, pMask.hdc, 0, 0, vbSrcAnd" & vbCrLf & vbCrLf & "1) Instead of using vbSrcCopy, we used vbSrcAnd as our paint type with our MASK" & vbCrLf & "2) The result is that the WHITE portions of our picture did NOT overwrite the background!"
'    PointTo pMap
'    PointTo pMask
Case 12
    str = "We now BitBlt into the same coordinates (10,10) the actual image this time around with..." & vbCrLf & vbCrLf & "BitBlt pMap.hdc, 10, 10, pNormal.Width, pNormal.Height, pNormal.hdc, 0, 0, vbSrcPaint" & vbCrLf & vbCrLf & "1) This time, we used the paint type vbSrcPaint with our actual image." & vbCrLf & "2) The result is that the BLACK portions of our picture did NOT overwrite the background!" '"Click NEXT to understand what just happened..."
    PointTo pMap
    PointTo pNormal
    BitBlt pMap.hdc, 10, 10, pNormal.Width, pNormal.Height, pNormal.hdc, 0, 0, vbSrcPaint
    iStep = iStep + 1
'Case 13
'    str = "BitBlt pMap.hdc, 10, 10, pNormal.Width, pNormal.Height, pNormal.hdc, 0, 0, vbSrcPaint" & vbCrLf & vbCrLf & "1) This time, we used the paint type vbSrcPaint with our actual image." & vbCrLf & "2) The result is that the BLACK portions of our picture did NOT overwrite the background!"
'    PointTo pMap
'    PointTo pNormal
Case 14
    str = "The result is that our image has been placed over top of the background, without overwriting that part of the background we wished to leave intact!"
    PointTo pMap
Case 15
    str = "Now in order to create the illusion of a sprite moving across our map, we must remove the current image and place a new image at the newly desired location." & vbCrLf & vbCrLf & "So how do we ""remove"" our current sprite from the map?"
    PointTo pMap
Case 16
    str = "In order to ""remove"" our sprite from the map, we can do a couple things. One is refresh the whole map, however, this is not practical when looking for speed when dealing with larger maps. Plus, it will more than likely cause your map to ""flicker.""" & vbCrLf & vbCrLf & "So instead, we just replace that portion of our map with a copy from our Background as is highlighted in red below. (Note that the red highlight is simply for displaying in this tutorial -- you do not draw a red square for this.)"
    PointTo pMap
    PointTo pBackgrnd
    pBackgrnd.DrawWidth = 2
    pBackgrnd.Line (9, 9)-(11 + pNormal.Width, 11 + pNormal.Height), vbRed, B
Case 17
    str = "BitBlt pMap.hdc, 10, 10, pNormal.Width, pNormal.Height, pBackgrnd.hdc, 10, 10, vbSrcCopy" & vbCrLf & vbCrLf & "And PRESTO, he's gone! We then paint our next image into our newly desired location, again BitBlt'ing the MASK with vbSrcAnd and then the image itself with vbSrcPaint to create the illusion of movement!" & vbCrLf & vbCrLf & "Piece of cake, right? Well..."
    PointTo pMap
    PointTo pBackgrnd
    BitBlt pMap.hdc, 10, 10, pNormal.Width, pNormal.Height, pBackgrnd.hdc, 10, 10, vbSrcCopy
Case 18
    str = "Now for the bad news: Keeping track of all our images, and which ones to use at what time, can be a real chore!" & vbCrLf & vbCrLf & "Why do I say that? Let's re-visit our PacMan sprite..."
    ' This simple restores my background picture from the map since I drew lines on it last time -- just ignore
    BitBlt pBackgrnd.hdc, 0, 0, pMap.Width, pMap.Height, pMap.hdc, 0, 0, vbSrcCopy
Case 19
    str = "For every sprite you use, you need an image of each possible instance of your character in order to do its animation over a background, along with a corresponding mask for each image. So for our PacMan character alone, we need 16 images!" & vbCrLf & vbCrLf & "One in each direction of his mouth closing and opening, and a mask for each picture. That's a lot of work in paintbrush to flip all those images every which way and keep track of in files!"
    PointTo fraDummy
Case 20
    str = "Especially if you suddenly decide to change poor ol' Packy! Then you'll have to re-do all those images, all over again!!" & vbCrLf & vbCrLf & "But there's some good news..."
    PointTo fraDummy
    PointTo pFang
Case 21
    str = "But before I get into the good news, realize that I've told you what you need to know concerning images and how to paint them over a background. Exactly HOW you get/make your images, where/how you store them and the code you use to manipulate them is up to you." & vbCrLf & vbCrLf & "However, stick around and I'll show you some tips and tricks which can make it much easier -- as long as you're not afraid of a little math."
Case 22
    str = "Part 2: Creating/Generating/Manipulating your sprites." & vbCrLf & vbCrLf & "The primary advantage to generating your sprite images is saving you time down the long road. And secondly, if done correctly, can eliminate the possibiliy of creating a bad sprite(s) by mistake. With that said, let's revisit our PacMan friend..."
Case 23
    str = "You'll notice this time around he DOESN'T have a black background AT THIS POINT. Basically, it needs to be any color OTHER than a color included in your bitmap picture, and since we have black but not white, I chose to make it WHITE. (I'll get into why later)." & vbCrLf & vbCrLf & "And the PictureBox name is pPac(0) -- the first element in an array of PictureBoxes -- with each element of the array holding the next sprite in my animation sequence to face East." & vbCrLf & vbCrLf & "What exactly do I mean by that?"
    PointTo pPac(0)
Case 24
    str = "What I mean is that for each image (animation) of my PacMan which will face East, I have a PictureBox array. For our PacMan, I have only two Easterly images I want to use in my game -- one with his mouth closing, the other with it opening. The PictureBox names are pPac(0) and pPac(1), respectively." & vbCrLf & vbCrLf & "Note that if I desired more animation sprites, I'd simply create however many more I need."
    PointTo pPac(0)
    PointTo pPac(1)
Case 25
    str = "Next, I generate images for each of them, starting with pPac(0). If you wish to see how I've done this, simply look at my included code later for the function ""CreatePacManFrames.""" & vbCrLf & vbCrLf & "The reason I'm not going into detail how will be explained in a moment. But for now, notice that I'm creating the other directional sprites WITHIN THE SAME PICTUREBOX."
    CreatePacManFrames pPac(0)
    PointTo pPac(0)
Case 26
    str = "Next, I run my ""CreateMasks"" procedure on pPac(0) to create the corresponding masks for each image as a ""second"" row WITHIN THE SAME PICTUREBOX." & vbCrLf & vbCrLf & "At this point, it's important to note that the end result contains all my needed sprites, meaning my actual images now have a BLACK background, and my MASKS are perfect compliments to the images."
    CreateMasks pPac(0)
    PointTo pPac(0)
Case 27
    str = "Side Note: My ""CreateMasks"" function assumes that the color in the upper left hand was the color of the background for the whole sprite. That way, it could effectively create the MASKS, and then BLACK OUT the backgrounds of the real images." & vbCrLf & vbCrLf & "So why again did I use WHITE as my initial background? Simply because that was one color that was NOT in my image. I could also have used GREEN or BLUE or any other color for my initial background that was NOT IN MY IMAGE for my particular ""CreateMasks"" function to work."
    PointTo pPac(0)
Case 28
    str = "Now I do the same with pPac(1), giving me all 16 PacMan images required for my particular game -- 8 actual images and their 8 cooresponding MASKS."
    CreatePacManFrames pPac(1)
    CreateMasks pPac(1)
    PointTo pPac(0)
    PointTo pPac(1)
Case 29
    str = "The reason I'm not going into detail on the ""how"" of these functions is because this is simply ONE WAY of doing sprite animations. What if your sprite requires diagonal directions? Or what if your sprite is an overhead view of a character rather than a side view of a flat PacMan?" & vbCrLf & vbCrLf & "For each sprite, the required work to create the generated images is different for each game, and even different per sprite within the game, as you'll see now..."
Case 30
    str = "For my ""Ghost"" images, I can't implement the same rotation logic for them since it would turn them sideways which I certainly don't want since they'd be lying on their side or upside down! So instead, I took the time to manually create my required 4 directional images within the same bitmap in paintbrush."
    PointTo pGhosts(0)
Case 31
    str = "But doesn't PacMan fight against 4 different colored Ghosts?" & vbCrLf & vbCrLf & "Correct, so I now have a choice. I can choose to manually create (and place on my form) each different colored ghost for my game. OR I can generate them within code by Loading anew each new pGhost(Index), copy over the picture from pGhost(0), and replace its primary color with their new color."
    PointTo pGhosts(0)
Case 32
    str = "It's indeed a dilemna..." & vbCrLf & vbCrLf & "Do I take the little bit of time needed to make them in paintbrush and just slap them onto the form? Or do I take a longer amount of time to write the generation code in the event I may later want to change their appearance, in which case I'd only have to change pGhost(0) rather than all 4 of them?"
    PointTo pGhosts(0)
Case 33
    str = "There's not necessarily a ""right"" answer. You just have to decide for yourself which would be the most practical and time-saving on different things when dealing with sprite creation." & vbCrLf & vbCrLf & "However, if your curious, I took the time to create a Ghost generator which I've just run. See my procedure ""CreateNewGhost"" if you're curious."
    CreateNewGhost vbCyan, 1
    PointTo pGhosts(0)
    PointTo pGhosts(1)
Case 34
    str = "One thing you'll notice with the Ghosts is that I have a YELLOW background instead of WHITE." & vbCrLf & vbCrLf & "Again, that's simply because for my particular MASK generator function, I'm assuming the upper left pixel of any given PictureBox of sprites is the color I ultimately wish to BLACKEN out. And since my Ghosts contain the color WHITE within them, I had to make the background any color other than any of those contained within my actual image." ' for my particular MASK generator function to work for them."
    PointTo pGhosts(0)
    PointTo pGhosts(1)
Case 35
    str = "So now I run my particular MASK generator on each of my Ghost(Index) Pictureboxes."
    CreateMasks pGhosts(0)
    CreateMasks pGhosts(1)
    PointTo pGhosts(0)
    PointTo pGhosts(1)
Case 36
    str = "IMPORTANT: Note that I'm simply showing you ONE WAY of arranging your sprites. Maybe you want all your sprite MASKS in their own PictureBox? Or maybe you want each directional image to be in their own PictureBox?" & vbCrLf & vbCrLf & "The bottom line is, once you've determined a system for how you want your sprites to be arranged/stored/generated, you then have to know how to access each sprite as needed."
    PointTo pPac(0)
    PointTo pPac(1)
    PointTo pGhosts(0)
    PointTo pGhosts(1)
Case 37
    str = "But already my particular method is a little sticky since I'm not consistent in how my sprites are arranged." & vbCrLf & vbCrLf & "Why is that? Because pPac(0) and pPac(1) are different images of PacMan moving in the same direction, whereas Ghost(0) and Ghost(1) are different ghosts altogether. But as long as I can keep it all straight, it's no problem..."
    PointTo pPac(0)
    PointTo pPac(1)
    PointTo pGhosts(0)
    PointTo pGhosts(1)
Case 38
    str = "I'll now give you a basic overview of how I am able to keep track of my particular arrangement of sprites, which makes BitBlt'ing the proper images at certain times much easier overall." & vbCrLf & vbCrLf & "Here's where the math comes into play..."
Case 39
    str = "For my sprites, I set up an Enum called Direction with the following elements..." & vbCrLf & vbCrLf & "Private Enum Direction" & vbCrLf & "East = 0" & vbCrLf & "West = 1" & vbCrLf & "South = 2" & vbCrLf & "North = 3" & vbCrLf & "End Enum"
Case 40
    str = "I also declared a couple constants for simplicity, for the Width and Height of my individual sprites..." & vbCrLf & vbCrLf & "Const W = 32" & vbCrLf & "Const H = 32"
Case 41
    str = "Now when I wish to refer to one of my sprites, I simply refer to a particular PictureBox of sprites, as well as it's Direction the sprite is supposed to face (in this case North), along with the Map Coordinates to place it on (10,10) by using this math..." & vbCrLf & vbCrLf & "BitBlt pMap.hdc, 10, 10, W, H, pPac(0).hdc, W * North, H, vbSrcAnd" & vbTab & "(The MASK)" & vbCrLf & "BitBlt pMap.hdc, 10, 10, W, H, pPac(0).hdc, W * North, 0, vbSrcPaint" & vbTab & "(The Image)"
    PointTo pPac(0)
    PointTo pMap
    BitBlt pMap.hdc, 0, 0, pMap.Width, pMap.Height, pBackgrnd.hdc, 0, 0, vbSrcCopy
    BitBlt pMap.hdc, 10, 10, W, H, pPac(0).hdc, W * North, H, vbSrcAnd
    BitBlt pMap.hdc, 10, 10, W, H, pPac(0).hdc, W * North, 0, vbSrcPaint
Case 42
    str = "Same goes for Ghost #2 heading West at location (20,20)..." & vbCrLf & vbCrLf & "BitBlt pMap.hdc, 20, 20, W, H, pGhosts(1).hdc, W * West, H, vbSrcAnd" & vbTab & "(The MASK)" & vbCrLf & "BitBlt pMap.hdc, 20, 20, W, H, pGhosts(1).hdc, W * West, 0, vbSrcPaint" & vbTab & "(The Image)"
    PointTo pGhosts(1)
    PointTo pMap
    BitBlt pMap.hdc, 0, 0, pMap.Width, pMap.Height, pBackgrnd.hdc, 0, 0, vbSrcCopy
    BitBlt pMap.hdc, 20, 20, W, H, pGhosts(1).hdc, W * West, H, vbSrcAnd
    BitBlt pMap.hdc, 20, 20, W, H, pGhosts(1).hdc, W * West, 0, vbSrcPaint
Case 43
    str = "As you can see, there's lots of power and flexibility as long as you can follow the math!" & vbCrLf & vbCrLf & "So take a deep breath and I'll try to quickly explain it for the particular way my sprites are arranged, as works for this example of Ghost #2 heading West..."
    PointTo pGhosts(1)
    PointTo pMap
Case 44
    str = "BitBlt pMap.hdc, 20, 20, W, H, pGhosts(1).hdc, W * West, H, vbSrcAnd" & vbTab & "(The MASK)" & vbCrLf & "BitBlt pMap.hdc, 20, 20, W, H, pGhosts(1).hdc, W * West, 0, vbSrcPaint" & vbTab & "(The Image)" & vbCrLf & vbCrLf & "When given a direction (East to North = 0 to 3), you multiply it by the Width of the each sprite (W = 32). So West gives (32 * 3 = 96), meaning we use that as the x-Coor on our PictureBox sprite group in question. First we paint the MASK, whose y-Coor is equal to the height of one sprite since we're going to the second row in our sprite picturebox, H = 32. We then paint the Image (still found at x-Coor of 96) but this time with a Height of 0 since it's at the top. Get it?"
    PointTo pGhosts(1)
    PointTo pMap
Case 45
    str = "Whew! Now if you didn't follow that entirely, DON'T WORRY!!" & vbCrLf & vbCrLf & "As long as you understood Part 1 of this tutorial on how to paint an Image over top of a background, while keeping parts of the background around the sprite intact, you're golden! Because again, this is just ONE WAY if arranging/creating sprites." & vbCrLf & vbCrLf & "Experience is still the best teacher, so just play with it and have fun!"
Case 46
    str = "Part 3: Now I'll share just a few bonus misc ""Tips & Tricks"" to help you a little more..." & vbCrLf & vbCrLf & "Screen Resolution: Not everyone's is the same as yours, so what do you do? Well, you can take a simpler approach to it like I did in my PacMan game, or you can change the user's Resolution and ""Colors"" with multiple API calls. (While ideal, you risk messing up their monitor display if your code isn't robust)."
Case 47
    str = "Two other solutions..." & vbCrLf & vbCrLf & "One is to use StretchBlt to ""Stretch"" or ""Contract"" your images using some ratio math with the Screen dimensions, but the result is that ""Stretching"" make your image ""blocky""..."
    PointTo pMap
    PointTo pPac(0)
    BitBlt pMap.hdc, 0, 0, pMap.Width, pMap.Height, pBackgrnd.hdc, 0, 0, vbSrcCopy
    StretchBlt pMap.hdc, 0, 0, W * 2, H * 2, pMask.hdc, 0, 0, pMask.Width, pMask.Height, vbSrcAnd
    StretchBlt pMap.hdc, 0, 0, W * 2, H * 2, pNormal.hdc, 0, 0, pNormal.Width, pNormal.Height, vbSrcPaint
Case 48
    str = "While ""Contracting"" makes them lose some clarity..."
    PointTo pMap
    PointTo pPac(0)
    BitBlt pMap.hdc, 0, 0, pMap.Width, pMap.Height, pBackgrnd.hdc, 0, 0, vbSrcCopy
    StretchBlt pMap.hdc, 0, 0, W * 0.75, H * 0.75, pMask.hdc, 0, 0, pMask.Width, pMask.Height, vbSrcAnd
    StretchBlt pMap.hdc, 0, 0, W * 0.75, H * 0.75, pNormal.hdc, 0, 0, pNormal.Width, pNormal.Height, vbSrcPaint
    'StretchBlt pMap.hdc, 10, 10, pMask.Width, pMask.Height, pMask.hdc, 0, 0, W * 2, H * 2, vbSrcAnd
    'StretchBlt pMap.hdc, 10, 10, pNormal.Width, pNormal.Height, pNormal.hdc, 0, 0, W * 2, H * 2, vbSrcPaint
Case 49
    str = "The final solution is to make seperate images for seperate Screen Resolutions if you want the best picture quality for various resolutions, but this is not recommended -- too much work and not enough pay-off unless you're a big commericial game company and you allow the user to run on various Screen Resolutions." & vbCrLf & vbCrLf & "So for now, keep it simple and just do something like I did in my PacMan game, though again, altering the Screen Resolution to fit your program is the ideal solution, but the most risky if you do not have robust code."
Case 50
    str = "Also, take note that even though I did my sprite creation/generation ""on the fly,"" you certainly don't have to." & vbCrLf & vbCrLf & "Feel free to create a separate program which takes a bitmap and generates all the images (and masks) into various files which your program can then load either into PictureBoxes or memory when your program starts up."
Case 51
    str = "Timer Vs. Loops..." & vbCrLf & vbCrLf & "In order to make our animation run at the same speed on everyone's machine, we need to be able to control how quickly the animation is drawn. You could use Timers, but this is highly NOT recommended as it is far too clunky for smooth game animation." & vbCrLf & vbCrLf & "Instead, use a Do...Loop with various ways needed to exit this ""game loop."""
Case 52
    str = "When using the ""game loop"" method..." & vbCrLf & vbCrLf & "DO NOT use the ""Sleep"" API command to ""slow down"" your animation (inefficent CPU usage). Instead, control your FPS (frames per second). This also makes it a snap to instantly change your game's speed. See my PacMan game for how this can be done."
Case 53
    str = "One final thing to talk about is the ""AutoRedraw"" property. Earlier on, I'd mentioned to set this property on all your Pictureboxes and Forms to TRUE. However, while you definitely need that set to TRUE on the PictureBoxes with your images (since they're invisible) you should really make your Form (or whatever your painting canvas is) to FALSE." & vbCrLf & vbCrLf & "It's faster, and there's less ""flicker."" It's a little more work to maintain a consistent image on your ""map,"" but just to sure to re-paint your ""map"" onto your Form in your Form's Paint event as I did in my PacMan game and you should be fine."
Case 54
    str = "That's it!" & vbCrLf & vbCrLf & "I hope you found this Graphical Tutorial on sprite animation helpful and thorough. While I didn't bother going into the game logic for my particular PacMan game, you can check it out in my included Work Easter Egg called Pac-MANager" & vbCrLf & vbCrLf & "And let me know how you felt about this on PSC and feel free to give me a vote! Thanks, and happy coding!"
Case Else
    Unload Me
    Exit Sub
End Select
    
    Tutor str, x, y
    
End Sub
