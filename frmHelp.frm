VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   4455
   End
   Begin VB.Label lblLL 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Audience Poll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblLL 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Fifty Fity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLL 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Call A Buddy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblHelpContent 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLifeLines 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Life Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblGame 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "How To Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmHelp"
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

Const HOWTOPLAY = "Who Wants To Be A Millionaire? (WWTBM) is a simple game to play. Once the game starts, the player starts with $0. With every correct answer they get in this trivia game, they move up to the next prize, which is roughly double the amount of the prize they already possess, or they can walk away with the current amount of money, ending the game.The player is provided with the option of 'Call A Buddy', 'Fifty Fifty', and 'Audience Poll'. Each of these tools can only be used once in the game. The 'check points' are the prizes that the player is garanteed after they get a wrong answer. As soon as the player get one wrong answer, the game is over, and their prize is their last check point. As the game progresses on, the questions will become more difficult."
Const LIFELINE1 = "'Call A Buddy' is a tool that permits the player to call one of his/her buddies, and they will provide an answer based on their intelligence level. Like 'Fifty Fifty' and 'Call A Buddy', this tool can only be used once."
Const LIFELINE2 = "'Fifty Fifty' is a tool that emits every option except two, one of these options will be the correct answer.Like 'Fifty Fifty' and 'Call A Buddy', this tool can only be used once."
Const LIFELINE3 = "'Audience Poll' is a tool that permits the player to poll the audience for the answer, the audience will provide what they think is the right answer, and using that information, the player must decide which one is the correct answer. Like 'Fifty Fifty' and 'Call A Buddy', this tool can only be used once."

Option Explicit

Private Sub cmdBack_Click()
    
    Unload Me
    
End Sub

Private Sub lblGame_Click()
    
    Dim K As Integer
    
    lblHelpContent.Caption = HOWTOPLAY
    
    'Hide the help options for life lines.
    
    For K = 1 To 3
        lblLL(K).Visible = False
    Next K

End Sub

Private Sub lblGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call lblGame_Click
    
End Sub

Private Sub lblLifeLines_Click()
    
    Dim K As Integer
    
    'Show the help options for the life lines.
    
    For K = 1 To 3
        lblLL(K).Visible = True
    Next K
    
End Sub

Private Sub lblLifeLines_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call lblLifeLines_Click

End Sub

Private Sub lblLL_Click(Index As Integer)
    
    Select Case Index
        Case 1
            lblHelpContent.Caption = LIFELINE1
        Case 2
            lblHelpContent.Caption = LIFELINE2
        Case 3
            lblHelpContent.Caption = LIFELINE3
    End Select
    
End Sub

Private Sub lblLL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
        Case 1
            lblHelpContent.Caption = LIFELINE1
        Case 2
            lblHelpContent.Caption = LIFELINE2
        Case 3
            lblHelpContent.Caption = LIFELINE3
    End Select
        
End Sub
