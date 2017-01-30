VERSION 5.00
Begin VB.Form frmPoll 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C000C0&
   BorderStyle     =   0  'None
   Caption         =   "Poll Results"
   ClientHeight    =   6510
   ClientLeft      =   1275
   ClientTop       =   3180
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back To Game"
      Height          =   1455
      Left            =   5880
      TabIndex        =   5
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lblResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   6240
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6240
      TabIndex        =   9
      Top             =   2520
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6240
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6240
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6240
      TabIndex        =   6
      Top             =   360
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpBarGraph 
      BackColor       =   &H80000008&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   3
      Left            =   4680
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape shpBarGraph 
      BackColor       =   &H80000008&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   2
      Left            =   3480
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape shpBarGraph 
      BackColor       =   &H80000008&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   1
      Left            =   2280
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape shpBarGraph 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   0
      Left            =   960
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C000C0&
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
      Left            =   1080
      TabIndex        =   4
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
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
      Left            =   2400
      TabIndex        =   3
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C000C0&
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
      Left            =   3600
      TabIndex        =   2
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C000C0&
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
      Left            =   4800
      TabIndex        =   1
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "frmPoll"
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

Option Explicit

Private Sub cmdBack_Click()
    
    Unload Me
    frmGame!QuestionTimer.Enabled = True
    
End Sub

Private Sub Form_Load()
    
    Dim NewHigh As Integer
    Dim Percent(0 To 3) As Integer
    Dim CorrectLocation As Integer
    Dim Difficulty As Integer
    Dim NoAnswer As Integer
    Dim Height As Single
    Dim K As Integer
    Dim CurrentLeft As Single
    Dim CurrentTop As Single
    
    For K = 1 To 5
        If frmGame!mnuSound.Checked = True Then
            PlaySound App.Path & "\Sounds\Beep.wav", 0, SND_SYNC Or SND_NODEFAULT
        End If
    Next K
    
    'Get the correct answer location and the difficulty of the question.
    
    Open App.Path & "\Poll.txt" For Input As #1
        Do While Not EOF(1)
            Input #1, CorrectLocation
            Input #1, Difficulty
        Loop
    Close #1
    
    'Delete the temporary file.
    
    Kill App.Path & "\Poll.txt"
        
    'Determine the percent of audience choosing their answer, depending on the difficulty, audience may choose the wrong answer.
        
    Select Case Difficulty
        Case 1
            Percent(CorrectLocation) = MakeRandom(70, 90)
            NewHigh = Percent(CorrectLocation)
            For K = 0 To 3
                If K <> CorrectLocation Then
                    If NewHigh <> 100 Then
                        Percent(K) = MakeRandom(1, 100 - NewHigh)
                    Else
                        Percent(K) = 0
                    End If
                    NewHigh = NewHigh + Percent(K)
                End If
            Next K
            NoAnswer = 100 - NewHigh
        Case 2
            Percent(CorrectLocation) = MakeRandom(30, 50)
            NewHigh = Percent(CorrectLocation)
            For K = 0 To 3
                If K <> CorrectLocation Then
                    If NewHigh <> 100 Then
                        Percent(K) = MakeRandom(1, 100 - NewHigh)
                    Else
                        Percent(K) = 0
                    End If
                    NewHigh = NewHigh + Percent(K)
                End If
            Next K
            NoAnswer = 100 - NewHigh
        Case 3
            Percent(CorrectLocation) = MakeRandom(1, 30)
            NewHigh = Percent(CorrectLocation)
            For K = 0 To 3
                If K <> CorrectLocation Then
                    If NewHigh <> 100 Then
                        Percent(K) = MakeRandom(1, 100 - NewHigh)
                    Else
                        Percent(K) = 0
                    End If
                    NewHigh = NewHigh + Percent(K)
                End If
            Next K
            NoAnswer = 100 - NewHigh
    End Select
    
    'Display the bar graph for each option.
    
    For K = 0 To 3
        shpBarGraph(K).Height = Percent(K) * 50
        CurrentLeft = shpBarGraph(K).Left
        CurrentTop = shpBarGraph(K).Top
        Height = shpBarGraph(K).Height
        shpBarGraph(K).Move CurrentLeft, CurrentTop - Height
    Next K
    
    'Display the results in sentences.
    
    DisplayResults Percent(), NoAnswer
    
End Sub

