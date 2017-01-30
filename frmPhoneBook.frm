VERSION 5.00
Begin VB.Form frmPhoneBook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C000C0&
   BorderStyle     =   0  'None
   Caption         =   "Phone Book"
   ClientHeight    =   4830
   ClientLeft      =   2715
   ClientTop       =   3015
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picConversation 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   5595
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect Call"
         Height          =   735
         Left            =   1560
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   4800
   End
   Begin VB.Label lblJob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   2880
      TabIndex        =   11
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblJob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   2880
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblJob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblJob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblJob 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmPhoneBook"
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

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Dim CorrectAnswer As String
Dim Order(1 To 5) As Integer
Dim BuddyName(1 To 15) As String
Dim Job(1 To 15) As String
Dim Intelligence(1 To 15) As Integer

Option Explicit
    
Private Sub cmdDisconnect_Click()
    
    If frmGame!mnuSound.Checked = True Then
        PlaySound App.Path & "\Sounds\HangUp.wav", 0, SND_ASYNC Or SND_NODEFAULT
    End If
    
    Unload Me
    frmGame!QuestionTimer.Enabled = True
    
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
        
    K = 0
    
    'Load the friends' names and their properties.
    
    Open App.Path & "\Phonebook.txt" For Input As #1
        Do While Not EOF(1)
            K = K + 1
            Input #1, BuddyName(K)
            Input #1, Job(K)
            Input #1, Intelligence(K)
        Loop
    Close #1
    
    'Read the correct answer.
    
    Open App.Path & "\Answer.txt" For Input As #1
        Do While Not EOF(1)
            Input #1, CorrectAnswer
        Loop
    Close #1
    
    'Delete the temporary file.
    
    Kill App.Path & "\Answer.txt"
    
    'Scramble the order that the friends are displayed in.
    
    RandomUniqueArray 1, 15, Order()
    
    'Display the friends and their job.
    
    For K = 0 To 4
        lblName(K).Caption = BuddyName(Order(K + 1))
        lblJob(K).Caption = Job(Order(K + 1))
    Next K
    
End Sub


Private Sub lblJob_Click(Index As Integer)
    
    Call lblName_Click(Index)
    
End Sub

Private Sub lblJob_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim K As Integer
    
    If frmGame!mnuSound.Checked = True Then
        PlaySound App.Path & "\Sounds\Beep.wav", 0, SND_ASYNC Or SND_NODEFAULT
    End If
    
    'Highlight the fields that the cursor is on.
    
    lblName(Index).BackColor = RGB(0, 255, 0)
    lblJob(Index).BackColor = RGB(0, 255, 0)
    
    'Unhighlight the fields that aren't selected.
    
    For K = 0 To 4
        If K <> Index Then
            lblName(K).BackColor = &HFF&
            lblJob(K).BackColor = &HFF&
        End If
    Next K
    
End Sub

Private Sub lblName_Click(Index As Integer)
    
    Dim K As Integer
    Dim Response As Integer
    Dim DType As Integer
    Dim Prompt As String
    Dim ChosenOne As String
    Dim Difficulty As Integer
    
    'Highlight the selected fields.
    
    lblName(Index).BackColor = RGB(0, 255, 0)
    lblJob(Index).BackColor = RGB(0, 255, 0)
    
    'Determine if user really wants to call the selected friend.
    
    DType = vbYesNo + vbQuestion
    Prompt = "Are you sure you want to call " & lblName(Index).Caption & "?"
    Response = MsgBox(Prompt, DType, "Call Confirmation")

    If Response = vbYes Then
        ChosenOne = lblName(Index).Caption
        
        'Show the chat.
        
        picConversation.Visible = True
        Difficulty = DetermineDifficulty(frmGame!lblPrize())
        If frmGame!mnuSound.Checked = True Then
            PlaySound App.Path & "\Sounds\Ring.wav", 0, SND_ASYNC Or SND_NODEFAULT
        End If
        Call_A_Buddy CorrectAnswer, frmGame!lblAnswer(), Intelligence(Order(Index + 1)), Difficulty, ChosenOne
        
        'Show the hang up button and allow the user to return to game.
        
        cmdDisconnect.Visible = True
    Else
        
        'Restore highlighted fields color.
        
        lblName(Index).BackColor = &HFF&
        lblJob(Index).BackColor = &HFF&
    End If
    
End Sub

Private Sub lblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim K As Integer
        
    If frmGame!mnuSound.Checked = True Then
        PlaySound App.Path & "\Sounds\Beep.wav", 0, SND_ASYNC Or SND_NODEFAULT
    End If

    'Highlight the fields that the cursor is on.
    
    lblName(Index).BackColor = RGB(0, 255, 0)
    lblJob(Index).BackColor = RGB(0, 255, 0)
    
    'Unhighlight the fields that aren't selected.
    
    For K = 0 To 4
        If K <> Index Then
            lblName(K).BackColor = &HFF&
            lblJob(K).BackColor = &HFF&
        End If
    Next K
    
End Sub

