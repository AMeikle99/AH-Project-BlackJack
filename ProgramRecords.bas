Attribute VB_Name = "Records"
Public Type typeUserDetails
    userNumber As Integer
    fullName As String * 30
    username As String * 20
    password As String * 20
    timesWon As Integer
    blackjackCount As Integer
    moneyEarned As Integer
    balance As Integer
    achievement1 As Boolean
    achievement2 As Boolean
    fontSize As String * 5
    colourScheme As String * 5
End Type

Public Type typePlayingCards
    suit As String
    rank As String
    cardImage As String
    playingValue As Integer
End Type

Public Type typeUsersHand
    usersCards(1 To 7) As typePlayingCards
    handValue As Integer
    cardCount As Integer
    betAmount As Integer
    softHand As Boolean
    blackjack As Boolean
    bust As Boolean
End Type

Public Type typeLeaderboardDetails
    userNumber As Integer
    fullName As String * 30
    timesWon As Integer
End Type





