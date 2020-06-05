VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      Caption         =   "User Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1215
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   4575
      Begin VB.Label lblWins 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label lblBlackjacks 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label lblBalance1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label lblFullName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSettings 
      BackColor       =   &H000080FF&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H000080FF&
      Caption         =   "Game Rules"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdLeaderboard 
      BackColor       =   &H000080FF&
      Caption         =   "Leaderboard"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   840
      ScaleHeight     =   5505
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      Begin VB.CommandButton cmdBlackjack 
         BackColor       =   &H000080FF&
         Caption         =   "Play Blackjack"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
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
   Begin WMPLibCtl.WindowsMediaPlayer titleMusic 
      Height          =   120
      Left            =   2400
      TabIndex        =   12
      Top             =   50
      Width           =   30
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   99
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   212
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim user As typeUserDetails 'This variable holds the details of the logged in user
Public userNumber As Integer 'This public variable is to identify which record position the user is saved at


Private Sub cmdBlackjack_Click()
'This sub routine unloads the main menu and loads and displays the Blackjack Form
    frmMainMenu.Hide
    Load frmBlackjack
    frmBlackjack.Show
    Unload Me
End Sub

Private Sub cmdExit_Click()
'This sub-routine fires when the user clicks the exit button
'It confirms that they wish to log out and if they confirm this then they return to the login screen

    Dim response As String
    
    'On a response from the messagebox which is 6, a yes, then the user is logged out and returned to the main menu
    response = MsgBox("Are you sure you wish to Log Out?", vbYesNo, "Log Out Confirmation")
    If response = "6" Then
        frmMainMenu.Hide
        Load frmLogin
        frmLogin.Show
        Unload frmMainMenu
    End If
End Sub

Private Sub cmdHelp_Click()
    frmMainMenu.Hide
    Load frmHelp
    frmHelp.Show
    Unload frmMainMenu
End Sub

Private Sub cmdLeaderboard_Click()
    frmMainMenu.Hide
    Load frmLeaderboard
    frmLeaderboard.Show
    Unload frmMainMenu
End Sub

Private Sub cmdSettings_Click()
    frmMainMenu.Hide
    Load frmSettings
    frmSettings.Show
    Unload frmMainMenu
End Sub

Private Sub Form_Load()
'On the load of the form the program opens the userDetails file and loads the record index which corresponds to the logged in user, i.e userNumber passed to the form

    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(user)
        Get #1, frmLogin.userNumber, user 'This recieves the record from the index which matches the value of userNumber (i.e matching user that was chosen to be logged in)
        Me.userNumber = user.userNumber
    Close #1
    
    'This section then displays the name, balance, number of blackjack and number of times the user has won
    lblFullName.Caption = "Name: " & user.fullName
    lblBalance1.Caption = "Balance: £" & user.balance
    lblBlackjacks.Caption = "Blackjacks: " & user.blackjackCount
    lblWins.Caption = "Times Won: " & user.timesWon
    
    'This fires if the users chosen font size is large
    If user.fontSize = "Large" Then
        
        'The following lines increase the font size of each label and button
        Me.lblFullName.fontSize = Me.lblFullName.fontSize + 1
        Me.lblBalance1.fontSize = Me.lblBalance1.fontSize + 1
        Me.lblBlackjacks.fontSize = Me.lblBlackjacks.fontSize + 1
        Me.lblWins.fontSize = Me.lblWins.fontSize + 1
        
        With Me
            .cmdBlackjack.fontSize = .cmdBlackjack.fontSize + 1
            .cmdExit.fontSize = .cmdExit.fontSize + 1
            .cmdHelp.fontSize = .cmdHelp.fontSize + 1
            .cmdLeaderboard.fontSize = .cmdLeaderboard.fontSize + 1
            .cmdSettings.fontSize = .cmdLeaderboard.fontSize + 1
        End With
        
    End If
    
    'This sets the source for the title music, relative to where the program is saved
    titleMusic.URL = App.Path & "\Sound Effects\Title Screen.wav"
    
End Sub

