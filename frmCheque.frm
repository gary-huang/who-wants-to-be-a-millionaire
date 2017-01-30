VERSION 5.00
Begin VB.Form frmCheque 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque"
   ClientHeight    =   5085
   ClientLeft      =   690
   ClientTop       =   2190
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   10965
   Begin VB.CommandButton cmdExit 
      Caption         =   "End Game"
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   4440
      Width           =   5535
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back To Game"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   5415
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblTextDollars 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   6840
      Picture         =   "frmCheque.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label lblMemo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label lblPrize 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script"
         Size            =   24
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Image imgCheque 
      Height          =   4425
      Left            =   0
      Picture         =   "frmCheque.frx":136D5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "frmCheque"
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

Private Sub cmdExit_Click()
    
    End_Program
    
End Sub

