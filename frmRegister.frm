VERSION 5.00
Begin VB.Form frmRegister 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   6060
   ClientLeft      =   4740
   ClientTop       =   1965
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtPassword2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3620
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtPassword1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3620
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3620
      TabIndex        =   10
      Top             =   2505
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   720
      ScaleHeight     =   4905
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      Begin VB.TextBox txtFullName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdRegister 
         BackColor       =   &H000080FF&
         Caption         =   "Register"
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
         Left            =   1920
         MaskColor       =   &H00008000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   2655
      End
      Begin VB.Label lblFullName 
         BackColor       =   &H00004080&
         Caption         =   "Enter Full Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblPassword2 
         BackColor       =   &H00004080&
         Caption         =   "Re-enter Password:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H00004080&
         Caption         =   "Enter a Username:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H00004080&
         Caption         =   "Enter a Password:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "            Blackjack, Or Bust             Register"
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
         TabIndex        =   2
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
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
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempUser As typeUserDetails
Dim recordCount As Integer

Private Sub cmdCancel_Click()
'This sub-routine runs if the cancel button is clicked
'Its purpose is to return the user to the login screen but only if the confirm their action

    'This variable holds the value returned by the confirmation message
    Dim response As String
    
    'The message box asks the user to confirm their decision ot cancel, it gives them a yes and a no button
    
    response = MsgBox("Are you sure you wish to cancel your registration?", vbYesNo, "Registration Cancel")
    
    'If the response is 6, the yes button, then the rgister form closes and the login screen is loaded
    If response = "6" Then
        frmRegister.Hide
        Unload frmRegister
        Load frmLogin
        frmLogin.Show
    'Other wise the user is asked to continue with registration
    Else
        MsgBox ("Please continue with your registration")
    End If
        
   
End Sub

Private Sub cmdRegister_Click()
'Sub Routine runs when the register button is clicked
'The purpose is to validate the details entered and if valid to add a new account to the userDetails file

    Dim fullName As String * 30
    Dim username As String * 20
    Dim password1 As String * 20
    Dim password2 As String * 20
    Dim filePath As String
    Dim validDetails As Boolean, foundUser As Boolean
    
    fullName = txtFullName.Text
    username = txtUsername.Text
    password1 = txtPassword1.Text
    password2 = txtPassword2.Text
    validDetails = True
    foundUser = False
    filePath = App.Path & "/Data Files/UserDetails.dat"
    
    'These lines reset the text colour of the labels to black
    lblFullName.ForeColor = &H0&
    lblUsername.ForeColor = &H0&
    lblPassword.ForeColor = &H0&
    lblPassword2.ForeColor = &H0&
    
    'Opens the random access userDetails file to be read/written to
    Open filePath For Random As #1 Len = Len(tempUser)
    'The subProcedure which validates the entered details is called with the relevant parameters passed to it.
    Call ValidateDetails(fullName, username, password1, password2, validDetails, foundUser)
    
    'If the details entered were valid then a new record is added to the userDetails file
    If validDetails = True Then
    
        'The with command directs all the fields below to the variable tempUser, saves time with repeatedly typing out the variable identifier
        'This section also assigns the initial values to all the fields in the records, to ensure none are null and all have a value
        With tempUser
            .userNumber = recordCount + 1
            .fullName = fullName
            .username = username
            .password = password1
            .timesWon = 0
            .blackjackCount = 0
            .moneyEarned = 0
            .balance = 500
            .achievement1 = False
            .achievement2 = False
            .fontSize = "Small"
            .colourScheme = "Green"
        End With
        
        'This places the new record at the next unoccupied position in the file
        Put #1, recordCount + 1, tempUser
        Close #1
        
        MsgBox "Register Successful", vbInformation, "Successful Registration"
        
        'This section closes the Register form and loads and displays the Login Form
        Unload frmRegister
        Load frmLogin
        frmLogin.Show
         
    'If the entered username matches that of an existing user then the user is alerted that this is the reason for the error
    ElseIf foundUser = True Then
        MsgBox "The chosen username appears to have been taken", vbExclamation, "Invalid Username"
        Close #1
    'The other reason the user will be alerted is because passwords may not match or a field was left blank
    Else
        MsgBox "Whoops! Make sure all fields are filled, passwords match and are at least 8 characters", vbExclamation, "Invalid Details"
        Close #1
    End If
    
End Sub

Private Sub ValidateDetails(ByRef fullName As String, ByRef username As String, ByRef password1 As String, ByRef password2 As String, ByRef validDetails As Boolean, ByRef foundUser As Boolean)
'This sub-procedure is used to validate the details entered by the user in the provided textboxes
'It accepts parameters so that the sub-procedure is modular and can be used and tested independently.

    'This first validation is to check that a value has been entered into the full name field
    If Len(fullName) = 0 Then
        validDetails = False
        'If it is invalid then the corrseponding label is set to red
        lblFullName.ForeColor = &HFF&
    End If
    
    'This iterates through the userdetails file and compares the entered username with all the usernames associated with current accounts
    'If either the usrname is left blank or a matching user is found then validDetails is set to false and the label is coloured red to alert the user
    Do
        Get #1, , tempUser
        If tempUser.username = username Then
            foundUser = True
            validDetails = False
            lblUsername.ForeColor = &HFF&
        ElseIf Len(username) = 0 Then
            validDetails = False
            lblUsername.ForeColor = &HFF&
        End If
        
    Loop Until EOF(1) Or foundUser = True
    
    'If either password is found to be too short, under 8 characters, then the validDetails is set to false and the two labels are colored red
    If Len(password1) < 8 Or Len(password2) < 8 Then
        lblPassword.ForeColor = &HFF&
        lblPassword2.ForeColor = &HFF&
        validDetails = False
    End If
    
    'The last validation is to check whether the two entered values for password match, if they don't then validDetails is set to false and
    'the two labels are colored red.
    If password1 <> password2 Then
        lblPassword.ForeColor = &HFF&
        lblPassword2.ForeColor = &HFF&
        validDetails = False
    End If
        
End Sub

Private Sub Form_Load()
'The code in this load procedure iterates through the userDetails file and counts the number of records stored in it
'The purpose is so that when the user registers a new account then its details can be added as a record in the position after
'that of the last record

    recordCount = 0
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(tempUser)
        
    Do Until EOF(1)
        Get #1, , tempUser
        'if the userNumber is greater than 0 then it is a valid record and is added to the total
        If tempUser.userNumber > 0 Then
            recordCount = recordCount + 1
        End If
    Loop
    
    Close #1
    
End Sub
