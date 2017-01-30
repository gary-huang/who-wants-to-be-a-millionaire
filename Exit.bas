Attribute VB_Name = "Exit"
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

'This general procedure confirms if the user wants to exit the program.

Sub End_Program()
    
    Dim Response As Integer
    Dim DType As Integer
    
    DType = vbYesNo + vbQuestion
    
    Response = MsgBox("Are you sure you want to exit?", DType, "Exit")
    
    If Response = vbYes Then
        End
    End If
    
End Sub
