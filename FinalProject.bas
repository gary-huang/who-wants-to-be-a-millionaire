Attribute VB_Name = "FinalProject"
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

Global Const SND_APPLICATION = &H80         '  look for application specific association
Global Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Global Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Global Const SND_ASYNC = &H1         '  play asynchronously
Global Const SND_FILENAME = &H20000     '  name is a file name
Global Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Global Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Global Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Global Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Global Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Global Const SND_PURGE = &H40               '  purge non-static events for task
Global Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Global Const SND_SYNC = &H0         '  play synchronously (default)

Global PlayerName As String

Option Explicit

'This general procedure eliminates two wrong choices as the 'Fifty Fifty' lifeline.

Sub Fifty_Fifty(CorrectAnswer As String, LabelNames As Variant)
    
    Dim CorrectLocation As Integer
    Dim K As Integer, Count As Integer
    
    Count = 0
    
    frmGame!mnuFifty.Enabled = False
    frmGame!imgFifty.Enabled = False
    frmGame!Cross(2).Visible = True
    frmGame!Cross(3).Visible = True
    
    For K = 0 To 3
        If CorrectAnswer = LabelNames(K).Caption Then
            CorrectLocation = K
        End If
    Next K
    
    For K = 0 To 3
        If LabelNames(K).Caption <> CorrectAnswer And Count <> 2 Then
            LabelNames(K).Enabled = False
            Count = Count + 1
        End If
    Next K
    
End Sub

'This general procedure allows the user to pick a friend and call for an answer.

Sub Call_A_Buddy(CorrectAnswer As String, LabelNames As Variant, Intelligence As Integer, Difficulty As Integer, ChosenOne As String)
    
    Dim CorrectLocation As Integer
    Dim K As Integer
    Dim Percent As Integer
    Dim Decision As Integer
    
    For K = 0 To 3
        If CorrectAnswer = LabelNames(K).Caption Then
            CorrectLocation = K
        End If
    Next K
    
    Percent = DetermineMaxChance(Intelligence, Difficulty)
    Decision = DetermineDecision(Percent, CorrectLocation)
    Conversation ChosenOne, Decision
    
End Sub

'This function determines the answer that the called up friend chooses.

Function DetermineDecision(Percent As Integer, CorrectLocation As Integer) As Integer
    
    Dim Temp As Integer
    Dim Num As Integer
    
    Select Case Percent
        Case 100
            DetermineDecision = CorrectLocation
        Case 90
            Temp = MakeRandom(1, 10)
            If Temp <= 9 Then
                DetermineDecision = CorrectLocation
            Else
                Num = MakeRandom(0, 3)
                Do
                    If Num = CorrectLocation Then
                        Num = MakeRandom(0, 3)
                    End If
                Loop While Num = CorrectLocation
                DetermineDecision = Num
            End If
        Case 80
            Temp = MakeRandom(1, 10)
            If Temp <= 8 Then
                DetermineDecision = CorrectLocation
            Else
                Num = MakeRandom(0, 3)
                Do
                    If Num = CorrectLocation Then
                        Num = MakeRandom(0, 3)
                    End If
                Loop While Num = CorrectLocation
                DetermineDecision = Num
            End If
        Case 60
            Temp = MakeRandom(1, 10)
            If Temp <= 6 Then
                DetermineDecision = CorrectLocation
            Else
                Num = MakeRandom(0, 3)
                Do
                    If Num = CorrectLocation Then
                        Num = MakeRandom(0, 3)
                    End If
                Loop While Num = CorrectLocation
                DetermineDecision = Num
            End If
        Case 50
            Temp = MakeRandom(1, 10)
            If Temp <= 5 Then
                DetermineDecision = CorrectLocation
            Else
                Num = MakeRandom(0, 3)
                Do
                    If Num = CorrectLocation Then
                        Num = MakeRandom(0, 3)
                    End If
                Loop While Num = CorrectLocation
                DetermineDecision = Num
            End If
        Case 30
            Temp = MakeRandom(1, 10)
            If Temp <= 3 Then
                DetermineDecision = CorrectLocation
            Else
                Num = MakeRandom(0, 3)
                Do
                    If Num = CorrectLocation Then
                        Num = MakeRandom(0, 3)
                    End If
                Loop While Num = CorrectLocation
                DetermineDecision = Num
            End If
        Case 20
            Temp = MakeRandom(1, 10)
            If Temp <= 2 Then
                DetermineDecision = CorrectLocation
            Else
                Num = MakeRandom(0, 3)
                Do
                    If Num = CorrectLocation Then
                        Num = MakeRandom(0, 3)
                    End If
                Loop While Num = CorrectLocation
                DetermineDecision = Num
            End If
    End Select
    
End Function

'This function determines the maximum chance of answering correctly that the called up friend will have.

Function DetermineMaxChance(Intelligence As Integer, Difficulty As Integer) As Integer
    
    Dim Percent As Integer
    
    Select Case Intelligence
        Case Is <= 3
            Select Case Difficulty
                Case 1
                Percent = 90
                Case 2
                Percent = 50
                Case 3
                Percent = 20
            End Select
        Case Is <= 5
            Select Case Difficulty
                Case 1
                Percent = 90
                Case 2
                Percent = 60
                Case 3
                Percent = 30
            End Select
        Case Is <= 8
            Select Case Difficulty
                Case 1
                Percent = 100
                Case 2
                Percent = 80
                Case 3
                Percent = 50
            End Select
        Case Is < 10
            Select Case Difficulty
                Case 1
                Percent = 100
                Case 2
                Percent = 90
                Case 3
                Percent = 80
            End Select
        Case 10
            Percent = 100
    End Select
    
    DetermineMaxChance = Percent
    
End Function

'This general procedure initiates the audience poll for the user.

Sub Audience_Poll(CorrectAnswer As String, LabelNames As Variant, Difficulty As Integer)
    
    Dim CorrectLocation As Integer
    Dim K As Integer, Count As Integer
    
    Count = 0
        
    frmGame!imgPoll.Enabled = False
    frmGame!mnuAudiencePoll.Enabled = False
    frmGame!Cross(4).Visible = True
    frmGame!Cross(5).Visible = True
    
    For K = 0 To 3
        If CorrectAnswer = LabelNames(K).Caption Then
            CorrectLocation = K
        End If
    Next K
    
    Open App.Path & "\Poll.txt" For Output As #1
        Write #1, CorrectLocation, Difficulty
    Close #1
    
    frmPoll.Show vbModal
    
End Sub

'This general procedure displays a new question.

Sub DisplayQuestion(Questions() As String, Difficulty As Integer, Index As Integer, LabelName As Label)
    
    LabelName.Caption = Questions(Difficulty, Index)
    
End Sub

'This general procedure displays all the possible answers.

Sub DisplayAnswers(Answers() As String, Difficulty As Integer, Index As Integer, LabelNames As Variant)
    
    Dim K As Integer
    Dim AnswersOrder(1 To 4) As Integer
    
    RandomUniqueArray 1, 4, AnswersOrder()
    
    For K = 1 To 4
        LabelNames(K - 1).Caption = Answers(Difficulty, Index, AnswersOrder(K))
        LabelNames(K - 1).Enabled = True
    Next K
    
End Sub

'This function determines the prize that the user is going for.

Function CurrentPrize(Labels As Variant) As Long
    
    Dim K As Integer
    Dim Prize As String
    
    K = 0
    Do
        If Labels(K).BackColor = RGB(0, 255, 0) Then
            Prize = Labels(K).Caption
            CurrentPrize = Val(Right$(Prize, Len(Prize) - 1))
        End If
        K = K + 1
    Loop Until CurrentPrize <> 0
    
End Function

'This function determines the next available prize after the current one.

Function NextPrize(Labels As Variant) As Integer
    
    Dim K As Integer
    Dim Prize As String
    
    K = 0
    Do
        If Labels(K).BackColor = RGB(0, 255, 0) Then
            Prize = Labels(K + 1).Caption
            NextPrize = Val(Right$(Prize, Len(Prize) - 1))
        End If
        K = K + 1
    Loop Until NextPrize <> 0
    
End Function

'This general procedure loads the questions from the database.

Sub LoadQuestions(FileName As String, Questions() As String)
    
    Dim K As Integer
    Dim LinesCount As Integer
    Dim QuestionsCount As Integer
    Dim Temp As String
    
    K = 1
    LinesCount = 0
    Open App.Path & "\" & FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, Temp
            LinesCount = LinesCount + 1
            If LinesCount Mod 5 = 1 Then
                QuestionsCount = (LinesCount \ 5) + 1
                Select Case QuestionsCount
                    Case Is <= 25
                    Questions(1, K) = Temp
                    Case Is <= 50
                    Questions(2, K) = Temp
                    Case Is <= 75
                    Questions(3, K) = Temp
                End Select
                K = K + 1
                If K = 26 Then
                    K = 1
                End If
            End If
        Loop
    Close #1

End Sub

'This general procedure loads all the answers from the database.

Sub LoadAnswers(FileName As String, Answers() As String)
    
    Dim K As Integer
    Dim Difficulty As Integer
    Dim LinesCount As Integer
    Dim QuestionsCount As Integer
    Dim AnswersCount As Integer
    Dim Temp As String
    
    K = 1
    LinesCount = 0
    
    Open App.Path & "\" & FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, Temp
            LinesCount = LinesCount + 1
            If LinesCount Mod 5 = 1 Then
                QuestionsCount = QuestionsCount + 1
            End If
            If LinesCount Mod 5 <> 1 Then
                AnswersCount = AnswersCount + 1
                Select Case QuestionsCount
                    Case Is <= 25
                        Difficulty = 1
                    Case Is <= 50
                        Difficulty = 2
                    Case Is <= 75
                        Difficulty = 3
                End Select
                Answers(Difficulty, K, AnswersCount) = Temp
                If AnswersCount = 4 Then
                    K = K + 1
                    AnswersCount = 0
                End If
                If K = 26 Then
                    K = 1
                    AnswersCount = 0
                End If
            End If
        Loop
    Close #1
    
End Sub

'This general procedure generates aa array of random and unique values.

Sub RandomUniqueArray(Low As Integer, High As Integer, Values() As Integer)
    
    Dim Temp As Integer
    Dim Pass As Boolean
    Dim K As Integer, N As Integer
    
    For K = 1 To UBound(Values())
        Values(K) = 0
    Next K
    
    N = 1
    Do
        Temp = MakeRandom(Low, High)
        Pass = True
        For K = 1 To UBound(Values())
            If Temp = Values(K) Then
                Pass = False
            End If
        Next K
        If Pass Then
            Values(N) = Temp
            N = N + 1
        End If
    Loop Until N = UBound(Values()) + 1
    
End Sub

'This function makes a random number in between the given range.

Function MakeRandom(Low As Integer, High As Integer) As Integer

    Dim Temp As Integer

    Temp = Int(Rnd * (High - Low + 1)) + Low
    
    MakeRandom = Temp

End Function

'This function determines the correct answer in the question.

Function Answer(Answers() As String, Difficulty As Integer, QuestionIndex As Integer) As String
    
    Dim K As Integer
    Dim Temp As String
    
    For K = 1 To 4
        Temp = Answers(Difficulty, QuestionIndex, K)
        If Left$(Temp, 1) = "*" Then
            Answer = Right$(Temp, Len(Temp) - 1)
            Answers(Difficulty, QuestionIndex, K) = Right$(Temp, Len(Temp) - 1)
        End If
    Next K
    
End Function

'This function determines the current difficulty of questions.

Function DetermineDifficulty(LabelNames As Variant) As Integer
    
    Dim K As Integer
    Dim Temp As String
    Dim TempPrize As Long
    
    DetermineDifficulty = 0
    K = 15
    
    Do
        K = K - 1
        If K = -1 Then
            DetermineDifficulty = 9000
        Else
            Temp = LabelNames(K).Caption
            If LabelNames(K).BackColor = RGB(0, 255, 0) Then
                TempPrize = Val(Right$(Temp, Len(Temp) - 1))
                Select Case TempPrize
                Case Is < 1000
                DetermineDifficulty = 1
                Case Is < 32000
                DetermineDifficulty = 2
                Case Is < 1000000
                DetermineDifficulty = 3
                End Select
            End If
        End If
    Loop While DetermineDifficulty = 0
    
End Function

'This general procedure highlights the current prize green.

Sub UpdatePrize(LabelNames As Variant)
    
    Dim K As Integer
    Dim Pass As Boolean
    
    K = 15
    
    Do
        K = K - 1
        If LabelNames(K).BackColor = RGB(0, 255, 0) Then
            LabelNames(K).BackColor = RGB(0, 0, 255)
            LabelNames(K + 1).BackColor = RGB(0, 255, 0)
            Pass = True
        End If
    Loop Until Pass

End Sub

'This general procedure freezes the program for a given time.

Sub Delay(Interval As Single)
    
    Dim Start, Finish
    
    Start = Timer
    
    Do
        Finish = Timer
    Loop Until Finish - Start > Interval
    
End Sub

'This general procedure displays the conversation in the lifeline "Call A Buddy".

Sub Conversation(BuddyName As String, AnswerIndex As Integer)
    Dim K As Integer, N As Integer
    Dim Msg(1 To 6) As String
    Dim Prize As Long
    Dim Question As String
    
    Question = frmGame!lblQuestion.Caption
    Prize = CurrentPrize(frmGame!lblPrize())
    Msg(1) = BuddyName & ": Hello."
    Msg(2) = "Host: Hello " & BuddyName & ". Your friend is going for $"
    Msg(2) = Msg(2) & Trim(Str(Prize)) & " and needs your help."
    Msg(3) = "You: " & BuddyName & ", here's the question: " & Question
    Msg(3) = Msg(3) & ". What is the answer?"
    Msg(4) = BuddyName & ": The answer is '" & frmGame!lblAnswer(AnswerIndex) & "'."
    Msg(5) = "You: Are you sure?"
    Msg(6) = BuddyName & ": I am positively sure."
    
    For N = 1 To 6
        For K = 1 To Len(Msg(N))
            frmPhoneBook!picConversation.Print Mid$(Msg(N), K, 1);
            Delay 0.03
            If K Mod 52 = 0 Or K = Len(Msg(N)) Then
                frmPhoneBook!picConversation.Print
            End If
            DoEvents
        Next K
    Next N

End Sub

'This general procedure displays the results of the audience poll in a sentence.

Sub DisplayResults(Percentage() As Integer, NoAnswerPercent As Integer)
    
    Dim Msg(1 To 4)
    Dim K As Integer
    Dim Choice As String
    
    For K = 1 To 4
        Select Case K
            Case 1
            Choice = "A"
            Case 2
            Choice = "B"
            Case 3
            Choice = "C"
            Case 4
            Choice = "D"
        End Select
        Msg(K) = Percentage(K - 1) & "% of the audience chose " & Choice
    Next K
    
    For K = 0 To 3
        frmPoll!lblResults(K).Caption = Msg(K + 1)
    Next K
    
    If NoAnswerPercent > 0 Then
        frmPoll!lblResults(4).Caption = NoAnswerPercent & "% of the audience did not vote."
    End If
    
End Sub

'This function gets a name from the user that is valid.

Function GetName() As String

    Dim Temp As String
    
    Temp = InputBox$("Enter your name:", "Name")
    If Temp = "" Then
        MsgBox "Invalid name.", vbCritical, "Error"
    End If
    
    If Temp <> "" Then
        GetName = UCase$(Left$(Temp, 1)) & Right$(Temp, Len(Temp) - 1)
    End If
    
End Function

'This general procedure displays the cheque.

Sub ShowCheque(Prize As Long)
    
    frmCheque!lblName.Caption = PlayerName
    frmCheque!lblPrize.Caption = Format$(Prize, "0.00")
    Select Case Prize
        Case 1000000
            frmCheque!lblMemo.Caption = "Winning 'WWTBM'."
        Case Is < 1000000
            frmCheque!lblMemo.Caption = "Being a chicken and walking away."
    End Select
    frmCheque!lblTextDollars.Caption = DetermineTextPrize(Prize)
    ShowDateCheque
    
    frmCheque.Show vbModal
    
End Sub

'This general procedure disables all the lifelines.

Sub DisableLifeLines()
    
    Dim K As Integer
    
    frmGame!mnuFifty.Enabled = False
    frmGame!mnuCallBuddy.Enabled = False
    frmGame!mnuAudiencePoll.Enabled = False
    frmGame!imgFifty.Enabled = False
    frmGame!imgCall.Enabled = False
    frmGame!imgPoll.Enabled = False
    For K = 0 To 5
        frmGame!Cross(K).Visible = True
    Next K

End Sub

'This general procedure enables all the lifelines.

Sub EnableLifeLines()
    
    Dim K As Integer
    
    frmGame!mnuFifty.Enabled = True
    frmGame!mnuCallBuddy.Enabled = True
    frmGame!mnuAudiencePoll.Enabled = True
    frmGame!imgFifty.Enabled = True
    frmGame!imgCall.Enabled = True
    frmGame!imgPoll.Enabled = True
    For K = 0 To 5
        frmGame!Cross(K).Visible = False
    Next K

End Sub

'This general procedure resets the background color of all answers.

Sub ResetAnswerColors()
    
    Dim K As Integer
    
    For K = 0 To 3
        frmGame!lblAnswer(K).BackColor = &H8000000F
    Next K

End Sub

'This function determines the current checkpoint of the user.

Function DetermineCheckPoint(QuestionsDone As Integer) As Long
    
    Dim CheckPoint As Long
    
    Select Case QuestionsDone
        Case 5 To 9
            CheckPoint = 1000
        Case 10 To 14
            CheckPoint = 32000
        Case 15
            CheckPoint = 1000000
        Case Else
            CheckPoint = 0
    End Select
    
    DetermineCheckPoint = CheckPoint
    
End Function

'This general procedure highlights the wrong answer that the user selected in red.

Sub HighlightWrongAnswer(ClickedIndex As Integer)

    frmGame!lblAnswer(ClickedIndex).BackColor = RGB(255, 0, 0)

End Sub

'This general procedure highlights the correct answer in green.

Sub HighlightRightAnswers(CorrectAnswer As String)
    
    Dim K As Integer
    
    For K = 0 To 3
        If CorrectAnswer = frmGame!lblAnswer(K).Caption Then
            frmGame!lblAnswer(K).BackColor = RGB(0, 255, 0)
        Else
            frmGame!lblAnswer(K).BackColor = &H8000000F
        End If
    Next K

End Sub

'This general procedure initiates the prize indicators.

Sub ResetPrizeColors()
    
    Dim K As Integer
    
    For K = 0 To 14
        frmGame!lblPrize(K).BackColor = &HC00000
    Next K

End Sub

'This function determines the prize in a sentence format.

Function DetermineTextPrize(Prize As Long) As String
    
    Dim K As Integer
    Dim ZeroCount As Integer
    Dim StLength As Integer
    Dim St As String
    Dim EndSt As String
    Dim NonZeroCount As Integer
    Dim Num As Integer
    Dim BeginningSt As String
    
    St = Str(Prize)
    StLength = Len(St)
    ZeroCount = 0
    
    For K = 1 To Len(St)
        If Mid$(St, K, 1) = "0" Then
            ZeroCount = ZeroCount + 1
        End If
    Next K
    
    Select Case ZeroCount
        Case 2
            EndSt = " hundred"
        Case Is <= 5
            EndSt = " thousand"
        Case 6
            EndSt = " million"
    End Select
    
    NonZeroCount = StLength - ZeroCount
    
    Num = Val(Left$(St, NonZeroCount))
    
    Select Case Num
        Case 1
            BeginningSt = "One"
        Case 2
            BeginningSt = "Two"
        Case 3
            BeginningSt = "Three"
        Case 4
            BeginningSt = "Four"
        Case 5
            BeginningSt = "Five"
        Case 8
            BeginningSt = "Eight"
        Case 16
            BeginningSt = "Sixteen"
        Case 32
            BeginningSt = "Thirty two"
        Case 64
            BeginningSt = "Sixty four"
        Case 125
            BeginningSt = "One hunred twenty five"
        Case 255
            BeginningSt = "Two hundred fifty five"
        Case 555
            BeginningSt = "Five hundred fifty five"
    End Select
    
    DetermineTextPrize = BeginningSt & EndSt
    
End Function

'This function determines if the answer that the user selected is final.

Function IsFinalAnswer(IsGameOver As Boolean) As Boolean
    
    Dim Temp As String
    
    frmFinalAnswer.Show vbModal
    
    If Not IsGameOver Then
        Open App.Path & "\FATemp.txt" For Input As #1
            Input #1, Temp
        Close #1
        Kill App.Path & "\FATemp.txt"
    End If
    
    If Temp = "Yes" Then
        IsFinalAnswer = True
    ElseIf Temp = "No" Then
        IsFinalAnswer = False
    End If
    
End Function

'This function determines if the user walked away with the current money.

Function IsWalkAway() As Boolean
    
    Dim Temp As String
    
    frmWalkOrStay.Show vbModal
    Open App.Path & "\Temp.txt" For Input As #1
        Input #1, Temp
    Close #1
    Kill App.Path & "\Temp.txt"
    
    Select Case Temp
        Case "Walk"
            IsWalkAway = True
        Case "Stay"
            IsWalkAway = False
    End Select

End Function

'This general procedure shows the date on the cheque.

Sub ShowDateCheque()
    
    frmCheque!lblDay.Caption = Format$(Day(Date), "00")
    frmCheque!lblMonth.Caption = Format$(Month(Date), "00")
    frmCheque!lblYear.Caption = Format$(Year(Date), "00")
    
End Sub
