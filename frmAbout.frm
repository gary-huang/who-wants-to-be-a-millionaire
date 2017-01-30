VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3570
   ClientLeft      =   2760
   ClientTop       =   3810
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "#GIMME100PERCENT"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Questions from: Gary Huang and Who Wants To Be A Millionaire"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Sound from: various artists"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Program made by: Gary Huang"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Who Wants To Be A Millionaire?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdBack_Click()
    
    Unload Me
    
End Sub

