VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   9600
   ClientLeft      =   1005
   ClientTop       =   1365
   ClientWidth     =   9180
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   9180
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   9495
      Left            =   120
      ScaleHeight     =   9435
      ScaleWidth      =   8835
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Timer QuestionTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   3480
   End
   Begin VB.Frame Frame2 
      Caption         =   "Answers"
      Height          =   5655
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   6615
      Begin VB.Label Label4 
         Caption         =   "D."
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
         Left            =   240
         TabIndex        =   25
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "C."
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
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "B."
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
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A."
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
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblAnswer 
         Alignment       =   2  'Center
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
         Height          =   975
         Index           =   3
         Left            =   840
         TabIndex        =   21
         Top             =   4440
         Width           =   5415
      End
      Begin VB.Label lblAnswer 
         Alignment       =   2  'Center
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
         Height          =   975
         Index           =   2
         Left            =   840
         TabIndex        =   20
         Top             =   3120
         Width           =   5415
      End
      Begin VB.Label lblAnswer 
         Alignment       =   2  'Center
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
         Height          =   975
         Index           =   1
         Left            =   840
         TabIndex        =   19
         Top             =   1800
         Width           =   5415
      End
      Begin VB.Label lblAnswer 
         Alignment       =   2  'Center
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
         Height          =   975
         Index           =   0
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Question"
      Height          =   2295
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   6615
      Begin VB.Label lblQuestion 
         BackColor       =   &H8000000E&
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
         Height          =   1695
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Line Cross 
      BorderColor     =   &H000000C0&
      BorderWidth     =   8
      Index           =   5
      X1              =   7080
      X2              =   8520
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Cross 
      BorderColor     =   &H000000C0&
      BorderWidth     =   8
      Index           =   4
      X1              =   7080
      X2              =   8520
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Line Cross 
      BorderColor     =   &H000000C0&
      BorderWidth     =   8
      Index           =   3
      X1              =   4800
      X2              =   6240
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Cross 
      BorderColor     =   &H000000C0&
      BorderWidth     =   8
      Index           =   2
      X1              =   4800
      X2              =   6240
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Line Cross 
      BorderColor     =   &H000000C0&
      BorderWidth     =   8
      Index           =   0
      X1              =   2400
      X2              =   3840
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Line Cross 
      BorderColor     =   &H000000C0&
      BorderWidth     =   8
      Index           =   1
      X1              =   2400
      X2              =   3840
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Time Left:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgPoll 
      Height          =   975
      Left            =   7080
      Picture         =   "frmGame.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image imgFifty 
      Height          =   855
      Left            =   4920
      Picture         =   "frmGame.frx":1A3BF
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblTimeLeft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   1815
   End
   Begin VB.Image imgCall 
      Height          =   855
      Left            =   2400
      Picture         =   "frmGame.frx":558EA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$1000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$555000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$255000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$125000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$64000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$32000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$16000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$8000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$4000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "Sound On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Life Lines"
      Begin VB.Menu mnuFifty 
         Caption         =   "Fifty Fifty"
      End
      Begin VB.Menu mnuCallBuddy 
         Caption         =   "Call A Buddy"
      End
      Begin VB.Menu mnuAudiencePoll 
         Caption         =   "Audience Poll"
      End
   End
End
Attribute VB_Name = "frmGame"
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

Dim CheckPoint As Long
Dim TimeVal As Single
Dim QuestionsDone As Integer
Dim NextUp As Integer
Dim QuestionOrder(1 To 5) As Integer
Dim Prize As Integer
Dim GameOver As Boolean
Dim Difficulty As Integer
Dim Answers(1 To 3, 1 To 25, 1 To 4) As String
Dim CorrectAnswer As String
Dim Question(1 To 3, 1 To 25) As String

Option Explicit

Private Sub Form_Load()
    Dim k, j, m
    Dim Response As Integer
    Dim DType As Integer
    
    If mnuSound.Checked = True Then
        PlaySound App.Path & "\Sounds\Ready.wav", 0, SND_ASYNC Or SND_NODEFAULT
    End If
    
    'Determines if the user is ready to play.
    
    DType = vbYesNo + vbQuestion
    Response = MsgBox(PlayerName & ", are you ready to play?", DType, "Ready?")
    GameOver = True
            
            
'    LoadQuestions "Database.txt", Question()
'    LoadAnswers "Database.txt", Answers()
'    For k = 1 To 25
'        For j = 1 To 4
'            If Answers(3, k, j) <> "" Then
'                m = m + 1
'            End If
'        Next j
'    Next k
'    Picture1.Print m
    
    If Response = vbYes Then
        EnableLifeLines
        CheckPoint = 0
        GameOver = False
        QuestionsDone = 0
        ResetAnswerColors
        ResetPrizeColors
        Randomize
        NextUp = 1
        LoadQuestions "Database.txt", Question()
        LoadAnswers "Database.txt", Answers()
        RandomUniqueArray 1, 25, QuestionOrder()

        'Changes the first prize indicator color to green.

        lblPrize(0).BackColor = RGB(0, 255, 0)
        Difficulty = DetermineDifficulty(lblPrize())
        DisplayQuestion Question(), Difficulty, QuestionOrder(NextUp), lblQuestion
        CorrectAnswer = Answer(Answers(), Difficulty, QuestionOrder(NextUp))
        DisplayAnswers Answers(), Difficulty, QuestionOrder(NextUp), lblAnswer()

        'Begins countdown.

        QuestionTimer.Enabled = True
        TimeVal = 30
    End If

End Sub

Private Sub imgCall_Click()
             
    If Not GameOver Then
        If mnuSound.Checked = True Then
            PlaySound App.Path & "\Sounds\PageFlip.wav", 0, SND_ASYNC Or SND_NODEFAULT
        End If
        
        'Disables ways to activate this lifeline again.

        imgCall.Enabled = False
        mnuCallBuddy.Enabled = False
        
        'Pause the timer.
        
        QuestionTimer.Enabled = False
        
        'Store the correct answer.
        
        Open App.Path & "\Answer.txt" For Output As #1
            Write #1, CorrectAnswer
        Close #1
        frmPhoneBook.Show vbModal
        
        'Cross out the lifeline.
        
        Cross(0).Visible = True
        Cross(1).Visible = True
    End If
    
End Sub

Private Sub imgFifty_Click()
    
    If Not GameOver Then
        If mnuSound.Checked = True Then
            PlaySound App.Path & "\Sounds\CoinToss.wav", 0, SND_ASYNC Or SND_NODEFAULT
            Delay 0.5
            PlaySound App.Path & "\Sounds\CoinDrop.wav", 0, SND_ASYNC Or SND_NODEFAULT
        End If
        Fifty_Fifty CorrectAnswer, lblAnswer()
    End If
    
End Sub

Private Sub imgPoll_Click()
    
    If Not GameOver Then
        
        'Pause the timer.
        
        QuestionTimer.Enabled = False
        Audience_Poll CorrectAnswer, lblAnswer(), Difficulty
        
        'Disables ways to activate this lifeline again.
        
        imgPoll.Enabled = False
        mnuAudiencePoll.Enabled = False
        
        'Cross out the lifeline.
        
        Cross(4).Visible = True
        Cross(5).Visible = True
    End If
    
End Sub

Private Sub lblAnswer_Click(Index As Integer)
    
    Dim k As Integer
    Dim DType As Integer
    Dim Response As Integer
    Dim Temp As Boolean
    Dim AnswerIsFinal As Boolean
                      
    If Not GameOver Then
        If mnuSound.Checked = True Then
            PlaySound App.Path & "\Sounds\Click.wav", ByVal 0&, SND_ASYNC Or SND_NODEFAULT
        End If
        
        'Color the selected answer green.
        
        lblAnswer(Index).BackColor = RGB(0, 255, 0)
        AnswerIsFinal = IsFinalAnswer(GameOver)
        If AnswerIsFinal Then
            
            'Determine if selected answer is correct answer.
            
            If lblAnswer(Index).Caption = CorrectAnswer Then
                QuestionsDone = QuestionsDone + 1
                CheckPoint = DetermineCheckPoint(QuestionsDone)
                
                'Display cheque if user beat the game.
                
                If QuestionsDone = 15 Then
                    GameOver = True
                    If mnuSound.Checked = True Then
                        PlaySound App.Path & "\Sounds\Cheer.wav", 0&, SND_ASYNC Or SND_NODEFAULT
                    End If
                    MsgBox "You win!", vbInformation, "Winner"
                    ShowCheque 1000000
                Else
                    If mnuSound.Checked = True Then
                        PlaySound App.Path & "\Sounds\Tada.wav", 0&, SND_ASYNC Or SND_NODEFAULT
                    End If
                    
                    'Determine if player is walking away or staying in game.
                    
                    Temp = IsWalkAway()
                    Select Case Temp
                        Case True
                            GameOver = True
                            QuestionTimer.Enabled = False
                            If mnuSound.Checked = True Then
                                PlaySound App.Path & "\Sounds\Chicken.wav", ByVal 0&, SND_ASYNC Or SND_NODEFAULT
                            End If
                        Case False
                            GameOver = False
                    End Select
                    If GameOver = False Then
                        
                        'Generate a new order for questions.
                        
                        If QuestionsDone Mod 5 = 0 Then
                            RandomUniqueArray 1, 25, QuestionOrder()
                            NextUp = 0
                        End If
                             
                        'Reset alloted time and display next question if user chose to stay.

                        Difficulty = DetermineDifficulty(lblPrize())
                        UpdatePrize lblPrize()
                        NextUp = NextUp + 1
                        DisplayQuestion Question(), Difficulty, QuestionOrder(NextUp), lblQuestion
                        CorrectAnswer = Answer(Answers(), Difficulty, QuestionOrder(NextUp))
                        DisplayAnswers Answers(), Difficulty, QuestionOrder(NextUp), lblAnswer()
                        ResetAnswerColors
                        QuestionTimer.Enabled = True
                        TimeVal = 30
                    Else
                        
                        'Display cheque if user chose to walk away.
                        
                        ShowCheque CurrentPrize(lblPrize())
                        DisableLifeLines
                    End If
                End If
            Else
                If mnuSound.Checked = True Then
                    PlaySound App.Path & "\Sounds\Boo.wav", 0, SND_ASYNC Or SND_NODEFAULT
                End If
                QuestionTimer.Enabled = False
                MsgBox "You lose!", vbCritical, "Loser"
                
                'If already passed a checkpoint, final prize is checkpoint.
                
                If CheckPoint > 0 Then
                    ShowCheque CheckPoint
                End If
                HighlightWrongAnswer Index
                HighlightRightAnswers CorrectAnswer
                GameOver = True
                DisableLifeLines
            End If
        Else
            
            'Continue countdown if selected answer is not final answer.
            
            ResetAnswerColors
            If Not GameOver Then
                QuestionTimer.Enabled = True
            End If
        End If
    Else
        MsgBox "Game is already over!", vbCritical, "Error"
    End If
    
    'If game is over, determine if user wants to start a new game.
    
    If GameOver Then
        DType = vbYesNo + vbQuestion
        Response = MsgBox("Play again?", DType, "New Game")
        If Response = vbYes Then
            Call Form_Load
        End If
    End If
    
End Sub

Private Sub lblAnswer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim k As Integer
    
    'If game is not over, highlight the answer that the cursor is on.
    
    If Not GameOver Then
        lblAnswer(Index).BackColor = RGB(0, 255, 0)
        For k = 0 To 3
            If k <> Index Then
                lblAnswer(k).BackColor = &H8000000F
            End If
        Next k
    End If
    
End Sub

Private Sub mnuAudiencePoll_Click()
    
    'Pause the countdown.
    
    QuestionTimer.Enabled = False
    Audience_Poll CorrectAnswer, lblAnswer(), Difficulty

End Sub

Private Sub mnuCallBuddy_Click()
    
    'Pause the countdown.
    
    QuestionTimer.Enabled = False
    
    'Disable the ways the lifeline could be activated again.
    
    imgCall.Enabled = False
    mnuCallBuddy.Enabled = False
    
    'Store the correct answer.
    
    Open App.Path & "\Answer.txt" For Output As #1
        Write #1, CorrectAnswer
    Close #1
    frmPhoneBook.Show vbModal
    
    'Cross out the lifeline.
    
    Cross(0).Visible = True
    Cross(1).Visible = True
    
End Sub

Private Sub mnuExit_Click()
    
    End_Program
    
End Sub

Private Sub mnuFifty_Click()
    
    If mnuSound.Checked = True Then
        PlaySound App.Path & "\Sounds\CoinToss.wav", 0, SND_ASYNC Or SND_NODEFAULT
        Delay 0.5
        PlaySound App.Path & "\Sounds\CoinDrop.wav", 0, SND_ASYNC Or SND_NODEFAULT
    End If
    Fifty_Fifty CorrectAnswer, lblAnswer()
    
End Sub

Private Sub mnuNewGame_Click()
    
    Call Form_Load
    
End Sub

Private Sub mnuSound_Click()
    
    If mnuSound.Checked = True Then
        mnuSound.Checked = False
        PlaySound "", 0, SND_PURGE
    Else
        mnuSound.Checked = True
    End If
    
End Sub

Private Sub QuestionTimer_Timer()
        
    If QuestionTimer.Enabled = True Then
        
        'Update the time left.
        
        TimeVal = TimeVal - 0.01
        
        'Change the countdown timer color if user has less than 10 seconds.
        
        If TimeVal <= 10 Then
            lblTimeLeft.ForeColor = RGB(255, 0, 0)
        Else
            lblTimeLeft.ForeColor = &H80000008
        End If
        
        'Display the new time.
        
        lblTimeLeft.Caption = Format$(TimeVal, "0.00")
        
        'End game if out of time.
        
        If TimeVal <= 0 Then
            GameOver = True
            Unload frmFinalAnswer
            MsgBox "Too Slow"
            QuestionTimer.Enabled = False
            HighlightRightAnswers CorrectAnswer
            DisableLifeLines
        End If
    End If
    
End Sub
