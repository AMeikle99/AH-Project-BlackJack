VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   6150
   ClientLeft      =   4785
   ClientTop       =   1770
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H000080FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3040
      Width           =   2055
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   840
      ScaleHeight     =   4905
      ScaleWidth      =   4785
      TabIndex        =   2
      Top             =   600
      Width           =   4815
      Begin VB.CommandButton cmdRegister 
         BackColor       =   &H000080FF&
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MaskColor       =   &H00008000&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "            Blackjack, Or Bust            Login"
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
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H00004080&
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H00004080&
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1500
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempUser As typeUserDetails 'Declares a variable which is based on the user defined type, this creates a variable with several properties based on that type
Public userNumber As Integer    'This creates a public variable which stores the record number of the user being logged in, it is public so it can be accessed by
                                'the main menu during its load procedure
                                

Private Sub cmdEnd_Click()
    End
End Sub

Private Sub cmdLogin_Click()
'This subroutine executes on the click of the Login Button
'The function of it is to validate the username and password entered by the user against the stored userDetails file
'If the details match an existing user then access is granted, otherwise they will need to try again


    Dim foundUser As Boolean    'This variable is used to determine if a user can be found which matches the entered username
    Dim username As String * 20 'These variables store username and password entered by the user, they are fixed to 20 characters to match the length of the fields in the record structure
    Dim password As String * 20
    
    username = txtUsername.Text 'Initialization of variables values
    password = txtPassword.Text
    foundUser = False
    
    'Opens the userdetails file, which is random access, the length of a record is based on the size of a variable based on the corresponding record
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(tempUser)
        Do Until EOF(1) Or foundUser = True 'Iterates through file until it reaches the end or a matching user is found
            Get #1, , tempUser
            'Selection - matching user is found in file then foundUser becomes true
            If tempUser.username = username Then
                foundUser = True
            End If
        Loop
        
        'From the matching user record found, the porgram checks if the password also matches. If it does then the program provides the user further access
        If tempUser.password = password Then
            userNumber = tempUser.userNumber 'Public variable is assigned the value of the user number, this identifies the record number for direct access
            Close #1
            'This sequence loads the main menu form into memory,hides and unloads the login form and also displays the main menu form
            Load frmMainMenu
            frmLogin.Hide
            frmMainMenu.Show
            Unload frmLogin
        Else
            'If user details aren't valid then an error message is displayed, it is an Exclamation type so a graphic is displayed within the message box
            'which makes it seem more like an alert messsage
            Close #1
            MsgBox "Sorry, either the Username or Password is Incorrect", vbExclamation, "Incorrect Login"
        End If
    
End Sub

Private Sub cmdRegister_Click()
'On the click of the register button the program loads and displays the register form and then unloads and hides the login form
    frmLogin.Hide
    Unload frmLogin
    Load frmRegister
    frmRegister.Show
End Sub

Private Sub Form_Load()
'All for testing purposes. To ensure that the User file can be reset with my dummy data.
    Dim leaderboardUser As typeLeaderboardDetails
    

    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(tempUser)
    Open App.Path & "/Data Files/LeaderboardDetails.dat" For Random As #3 Len = Len(leaderboardUser)
    
    
    'Open App.Path & "/Data Files/DummyData.txt" For Input As #2
        
    Do
        Input #2, tempUser.userNumber, tempUser.fullName, tempUser.username, tempUser.password, tempUser.timesWon, tempUser.blackjackCount, tempUser.moneyEarned, tempUser.balance, tempUser.achievement1, tempUser.achievement2, tempUser.fontSize, tempUser.colourScheme
        Put #1, , tempUser
        
        leaderboardUser.userNumber = tempUser.userNumber
        leaderboardUser.fullName = tempUser.fullName
        leaderboardUser.timesWon = tempUser.timesWon
        
        If tempUser.userNumber >= 1 Then
            Put #3, , leaderboardUser
        End If
        
    Loop Until tempUser.userNumber < 1

    Close #1
    'Close #2
    Close #3

End Sub
