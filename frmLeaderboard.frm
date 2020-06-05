VERSION 5.00
Begin VB.Form frmLeaderboard 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leaderboard"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   7110
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1060
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   600
      ScaleHeight     =   6105
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      Begin VB.CommandButton cmdDescending 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Sort Descending"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   1365
      End
      Begin VB.CommandButton cmdAscending 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Sort Ascending"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   1365
      End
      Begin VB.ListBox lstLeaderboard 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3735
         ItemData        =   "frmLeaderboard.frx":0000
         Left            =   240
         List            =   "frmLeaderboard.frx":0002
         TabIndex        =   3
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "This leaderboard shows and ranks all players based on the number of ties they have won at Blackjack."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "    Blackjack, Or Bust"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmLeaderboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim userCount As Integer    'This variable holds the number of users stored in the leaderboard file

Private Sub cmdAscending_Click()
'This procedure fires when the user clicks on the Sort Ascending Button
'It sorts the users in ascending order (Lo - Hi) based on the number of times they've won

    Dim userArray() As typeLeaderboardDetails 'This array will store all the users from the Leaderboard file.
                                              'It is declared without a size so that it can be re-sized based on the no. of users in the file
    
    Call loadData(userArray()) 'Calls the sub-procedure to load data in from the file
    Call sortAscending(userArray()) 'Calls the sub-procedure which sorts the users in ascending order

End Sub

Private Sub cmdDescending_Click()

    Dim userArray() As typeLeaderboardDetails 'This array will store all the users from the Leaderboard file.
                                              'It is declared without a size so that it can be re-sized based on the no. of users in the file
    
    Call loadData(userArray()) 'Calls the sub-procedure to load data in from the file
    Call sortDescending(userArray()) 'Calls the sub-procedure which sorts the users in descending order
 
End Sub

Private Sub cmdReturn_Click()
'This sub-procedure fires when the user clicks on the Menu Button
'It hides and unloads the Leaderboard Form and then re-displays the Main Menu Form

    frmLeaderboard.Hide
    Load frmMainMenu
    frmMainMenu.Show
    Unload frmLeaderboard
End Sub

Private Sub Form_Load()
'This procedure fires when the form is loaded

    Dim index As Integer 'Variable used as a loop counter
    Dim tempUser As typeLeaderboardDetails 'Variable used to hold one user, used for counting how many are in file
    Dim userArray() As typeLeaderboardDetails 'This array will store all the users from the Leaderboard file.
                                              'It is declared without a size so that it can be re-sized based on the no. of users in the file
    
    Dim tempUser2 As typeUserDetails
    
    userCount = 0
    
    Open App.Path & "/Data Files/LeaderboardDetails.dat" For Random As #1 Len = Len(tempUser)
    
    Do
        Get #1, , tempUser
        If tempUser.userNumber >= 1 Then
            userCount = userCount + 1
        End If
    Loop Until tempUser.userNumber < 1
    
    Close #1
    
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(tempUser2)
    
    Get #1, frmMainMenu.userNumber, tempUser2
    
    If tempUser2.fontSize = "Large" Then
        Me.cmdAscending.fontSize = Me.cmdAscending.fontSize + 1
        Me.cmdDescending.fontSize = Me.cmdDescending.fontSize + 1
        Me.cmdReturn.fontSize = Me.cmdReturn.fontSize + 1
        Me.lstLeaderboard.fontSize = Me.lstLeaderboard.fontSize + 1
    End If
    
    Close #1
    Call loadData(userArray())
    Call sortDescending(userArray())
    
End Sub

Private Sub loadData(ByRef userArray() As typeLeaderboardDetails)

    Dim index As Integer
    ReDim userArray(userCount)
    
    Open App.Path & "/Data Files/LeaderboardDetails.dat" For Random As #1 Len = Len(userArray(1))
    
    For index = 1 To userCount
        Get #1, index, userArray(index)
    Next index
    
    Close #1

End Sub

Private Sub sortAscending(ByRef userlist() As typeLeaderboardDetails)

    Dim passCounter As Integer, innerCounter As Integer
    Dim index As Integer
    
    For passCounter = UBound(userlist()) To 1 Step -1
        For innerCounter = 1 To passCounter - 1
            If userlist(innerCounter).timesWon > userlist(innerCounter + 1).timesWon Then
                Call swapUsers(userlist(), innerCounter, innerCounter + 1)
            End If
        Next innerCounter
    Next passCounter
    
    lstLeaderboard.Clear
    For index = 1 To 10
        lstLeaderboard.AddItem (index & ". " & shortenLeaderboardName(userlist(index)) & " - " & userlist(index).timesWon & " wins")
        If index = userCount Then
            Exit For
        End If
    Next index


End Sub

Private Sub sortDescending(ByRef userlist() As typeLeaderboardDetails)

    Dim passCounter As Integer, innerCounter As Integer
    Dim index As Integer
    
    For passCounter = UBound(userlist()) To 1 Step -1
        For innerCounter = 1 To passCounter - 1
            If userlist(innerCounter).timesWon < userlist(innerCounter + 1).timesWon Then
                Call swapUsers(userlist(), innerCounter, innerCounter + 1)
            End If
        Next innerCounter
    Next passCounter
    
    lstLeaderboard.Clear
    For index = 1 To 10
        lstLeaderboard.AddItem (index & ". " & shortenLeaderboardName(userlist(index)) & " - " & userlist(index).timesWon & " wins")
        If index = userCount Then
            Exit For
        End If
    Next index
    

End Sub

Private Sub swapUsers(ByRef userlist() As typeLeaderboardDetails, ByVal lowerIndex As Integer, ByVal upperIndex As Integer)

    Dim tempUser As typeLeaderboardDetails
    
    tempUser = userlist(upperIndex)
    userlist(upperIndex) = userlist(lowerIndex)
    userlist(lowerIndex) = tempUser

End Sub

