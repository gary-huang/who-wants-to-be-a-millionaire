VERSION 5.00
Begin VB.Form frmMaster 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who Wants To Be A Millionaire?"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   2010
   ClientWidth     =   12075
   Icon            =   "frmMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   12075
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10920
      Top             =   6000
   End
   Begin VB.Image imgScreen2 
      Height          =   1455
      Left            =   5880
      Picture         =   "frmMaster.frx":0442
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image imgScreen1 
      Height          =   1440
      Left            =   3960
      Picture         =   "frmMaster.frx":15DEE
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgDisplay 
      Height          =   6975
      Left            =   0
      Picture         =   "frmMaster.frx":299A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12135
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuGameHelp 
         Caption         =   "View Help"
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: Final Project - Who Wants To Be A Millionaire (WWTBM)
'Author: Gary Huang
'Date: June 06th, 2013
'Files: WWTBM.vbp, frmAbout.frm, frmCheque.frm, frmFinalAnswer.frm, frmGame.frm,
'       frmHelp.frm, frmMaster.frm, frmPhoneBook.frm, frmPoll.frm, frmWalkOrStay.frm,
'       Exit.bas, FinalProject.bas, Beep.wav, Boo.wav, Cheer.wav, Chicken.wav,
'       Click.wav, CoinDrop,wav, CoinToss.wav, HangUp.wav, PageFlip.wav, Ready.wav,
'       Ring.wav, Tada.wav, Theme.wav
'Purpose: The purpose of this program is to provide the user with the experience of
'         the TV show 'Who Wants To Be A Millionaire' through simulation in this game.
'         This game has 3 levels of difficulty of questions, and 5 questions each level,
'         the grand prize of this game is one million dollars. Lifelines can be used,
'         the available lifelines are 'Call A Buddy', 'Fifty Fifty', and 'Audience Poll',
'         just like the TV show, each lifelines can only be used once. Users can choose
'         to walk away with the money they earned so far after correctly answering a
'         question, or they can choose to remain in the game to go for the next question.
    
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpSound As String, ByVal Flag As Long) As Long

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim Result As Long
    
    Select Case KeyAscii
        Case 13
            
            'Get a valid name from user.
            
            PlayerName = GetName()
            
            'Allow user to enter the game if name is valid.
            
            If PlayerName <> "" Then
                Timer1.Enabled = False
                Result = sndPlaySound("", SND_PURGE + SND_ASYNC)
                Unload Me
                frmGame.Show vbModal
            End If
        Case 27
            End_Program
    End Select
    
End Sub

Private Sub Form_Load()
    
    Dim Result As Long
    
    'Play the theme song.
    
    Result = sndPlaySound(App.Path & "\Sounds\Theme.wav", SND_ASYNC + SND_PURGE + SND_LOOP)
    Timer1.Enabled = True
    
End Sub



Private Sub mnuAbout_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
    
    End_Program
    
End Sub

Private Sub mnuGameHelp_Click()
    
    frmHelp.Show vbModal
    
End Sub

Private Sub mnuNewGame_Click()
            
    'Get a valid name if user hasn't entered a name already.
            
    If PlayerName = "" Then
            PlayerName = GetName()
    End If
    
    'Start a new game if entered name is valid.
    
    If PlayerName <> "" Then
        Timer1.Enabled = False
        Unload Me
        frmGame.Show vbModal
    End If
    
End Sub

Private Sub Timer1_Timer()
        
    'Change pictures every interval for a flashing effect.
        
    If Timer1.Enabled = True Then
        If imgDisplay.Picture = imgScreen2.Picture Then
            imgDisplay.Picture = imgScreen1.Picture
        Else
            imgDisplay.Picture = imgScreen2.Picture
        End If
    End If
    
End Sub
