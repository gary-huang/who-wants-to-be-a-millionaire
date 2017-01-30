VERSION 5.00
Begin VB.Form frmFinalAnswer 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Final Answer?"
   ClientHeight    =   1725
   ClientLeft      =   3360
   ClientTop       =   4140
   ClientWidth     =   4680
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Is this your final answer?"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   360
      Picture         =   "frmFinalAnswer.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFinalAnswer"
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

Private Sub cmdNo_Click()
    
    'Open the file to write the choice of user.
    
    Open App.Path & "\FATemp.txt" For Output As #1
        Write #1, "No"
    Close #1
    
    Unload Me
    
End Sub

Private Sub cmdYes_Click()
    
    'Open the file to write the choice of user.
    
    Open App.Path & "\FATemp.txt" For Output As #1
        Write #1, "Yes"
    Close #1
    
    Unload Me
    
End Sub

