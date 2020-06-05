Attribute VB_Name = "SubRoutines"
Public Function shortenName(ByRef tempUser As typeUserDetails)

    Dim currentCharacter As String
    Dim newString As String
    Dim characterCount As Integer
    
    characterCount = 0
    currentCharacter = " "
    
    While currentCharacter = " "
        currentCharacter = Mid(tempUser.fullName, 30 - characterCount, 1)
        characterCount = characterCount + 1
    Wend
    newString = Left(tempUser.fullName, 30 - (characterCount - 1))
    shortenName = newString

End Function

Public Function shortenLeaderboardName(ByRef tempUser As typeLeaderboardDetails)

    Dim currentCharacter As String
    Dim newString As String
    Dim characterCount As Integer
    
    characterCount = 0
    currentCharacter = " "
    
    While currentCharacter = " "
        currentCharacter = Mid(tempUser.fullName, 30 - characterCount, 1)
        characterCount = characterCount + 1
    Wend
    newString = Left(tempUser.fullName, 30 - (characterCount - 1))
    shortenLeaderboardName = newString

End Function

