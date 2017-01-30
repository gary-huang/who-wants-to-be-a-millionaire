VERSION 5.00
Begin VB.Form frmWalkOrStay 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Walk Away Or Continue?"
   ClientHeight    =   6210
   ClientLeft      =   1125
   ClientTop       =   1905
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdWalk 
      Caption         =   "Walk Away?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1200
      TabIndex        =   0
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Your Answer Was Correct!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   7815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Risk your money and push your luck or your knowledge...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Walk away with what you have earned so far..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   5520
      Picture         =   "frmWalkOrStay.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   1560
      Picture         =   "frmWalkOrStay.frx":3792
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmWalkOrStay"
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

Option Explicit

Private Sub cmdContinue_Click()
    
    'Store the player's decision to stay in the game.
    
    Open App.Path & "\Temp.txt" For Output As #1
        Write #1, "Stay"
    Close #1
    
    Unload Me
    
End Sub

Private Sub cmdWalk_Click()
    
    'Store the player's decision to walk away.
    
    Open App.Path & "\Temp.txt" For Output As #1
        Write #1, "Walk"
    Close #1
    
    Unload Me
        
End Sub

