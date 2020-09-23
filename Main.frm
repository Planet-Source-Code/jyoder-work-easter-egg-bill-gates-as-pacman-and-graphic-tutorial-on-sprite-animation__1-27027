VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Select Program"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pbPacMan 
      Caption         =   "Pac-MANager"
      Height          =   1215
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton pbTutorial 
      Caption         =   "Graphical Tutorial on Sprite Animation using BitBlt"
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    pbTutorial.Caption = pbTutorial.Caption & vbCrLf & vbCrLf & "Check this out FIRST!!"
End Sub

Private Sub pbTutorial_Click()
Dim frm As frmTutor
    Set frm = New frmTutor
    frm.Show vbModal
    Set frm = Nothing
End Sub

Private Sub pbPacMan_Click()
Dim frm As frmPacMan
    Set frm = New frmPacMan
    frm.LoadUpForm
    Set frm = Nothing
End Sub
