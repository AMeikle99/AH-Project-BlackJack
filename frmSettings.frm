VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
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
   Begin VB.PictureBox picColour 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      DrawMode        =   16  'Merge Pen
      FillColor       =   &H00004080&
      ForeColor       =   &H00004080&
      Height          =   635
      Left            =   3390
      ScaleHeight     =   630
      ScaleMode       =   0  'User
      ScaleWidth      =   2879.843
      TabIndex        =   9
      Top             =   3540
      Width           =   2895
      Begin VB.OptionButton choiceColour 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "Red Scheme"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   11
         Top             =   130
         Width           =   1560
      End
      Begin VB.OptionButton choiceColour 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "Green Scheme"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   130
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1485
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
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   4575
         Begin VB.Label lblFullName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00004080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   45
         End
         Begin VB.Label lblBalance1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2400
            TabIndex        =   14
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label lblAchievements 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   4395
         End
      End
      Begin VB.PictureBox picFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   655
         Left            =   2520
         ScaleHeight     =   660
         ScaleWidth      =   2895
         TabIndex        =   6
         Top             =   2160
         Width           =   2895
         Begin VB.OptionButton choiceFont 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            Caption         =   "Large Font"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   1540
            TabIndex        =   8
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton choiceFont 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            Caption         =   "Small Font"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdResetBalance 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Reset Balance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4200
         Width           =   1485
      End
      Begin VB.Label lblColour 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "Choose Colour Scheme:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lblFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "Choose Font Size:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2340
         Width           =   1695
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Blackjack, Or Bust - Settings"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim user As typeUserDetails 'This variable holds the details of the logged in user
Dim userNumber As Integer

Private Sub choiceColour_Click(index As Integer)
    
    If index = 0 Then
        user.colourScheme = "Green"
    ElseIf index = 1 Then
        user.colourScheme = "Red  "
    End If
    
    Call saveDetails
    
End Sub

Private Sub choiceFont_Click(index As Integer)

    If index = 0 Then
        Call smallFont
        user.fontSize = "Small"
    ElseIf index = 1 Then
        Call largeFont
        user.fontSize = "Large"
    End If

    Call saveDetails
    
End Sub

Private Sub cmdMenu_Click()
    frmSettings.Hide
    Load frmMainMenu
    frmMainMenu.Show
    Unload frmSettings
End Sub

Private Sub cmdResetBalance_Click()

    Dim confirmation As String
    
    confirmation = MsgBox("Are you sure you wish to reset your balance?", vbYesNo, "Balance Reset!")
    
    If confirmation = "6" Then
        user.balance = 500
        lblBalance1.Caption = "Balance: £" & user.balance
        Call saveDetails
    End If

End Sub

Private Sub Form_Load()

    userNumber = frmMainMenu.userNumber
    
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(user)
        Get #1, userNumber, user 'This recieves the record from the index which matches the value of userNumber (i.e matching user that was chosen to be logged in)
    Close #1
    
    If user.colourScheme = "Red  " Then
        choiceColour(1).Value = True
    Else
        choiceColour(0).Value = True
    End If
    
    If user.fontSize = "Large" Then
        choiceFont(1).Value = True
        Call largeFont
    Else
        choiceFont(0).Value = True
    End If
    
    lblFullName.Caption = "Name: " & user.fullName
    lblBalance1.Caption = "Balance: £" & user.balance
    
    If user.achievement1 = False And user.achievement2 = False Then
        lblAchievements.Caption = "No Achievments Earned Yet!"
    ElseIf user.achievement1 = True And user.achievement2 = False Then
        lblAchievements.Caption = "Achievement 1 Earned! - 5 Blackjacks"
    ElseIf user.achievement1 = False And user.achievement2 = True Then
        lblAchievements.Caption = "Achievement 2 Earned! - Earn £150"
    Else
        lblAchievements.Caption = "'Earn £150' and '5 Blackjacks' Achievements Earned"
    End If
    
    
    
    

End Sub

Private Sub largeFont()

    cmdResetBalance.fontSize = cmdResetBalance.fontSize + 1
    cmdMenu.fontSize = cmdMenu.fontSize + 1
    lblFont.fontSize = lblFont.fontSize + 1
    lblColour.fontSize = lblColour.fontSize + 1
    choiceFont(0).fontSize = choiceFont(0).fontSize + 1
    choiceFont(1).fontSize = choiceFont(1).fontSize + 1
    choiceColour(0).fontSize = choiceColour(0).fontSize + 1
    choiceColour(1).fontSize = choiceColour(1).fontSize + 1
    lblFullName.fontSize = lblFullName.fontSize + 1
    lblBalance1.fontSize = lblBalance1.fontSize + 1
    lblAchievements.fontSize = lblAchievements.fontSize + 1
    frameDetails.fontSize = frameDetails.fontSize + 1
    
    frmMainMenu.lblFullName.fontSize = frmMainMenu.lblFullName.fontSize + 1
    frmMainMenu.lblBalance1.fontSize = frmMainMenu.lblBalance1.fontSize + 1
    frmMainMenu.lblBlackjacks.fontSize = frmMainMenu.lblBlackjacks.fontSize + 1
    frmMainMenu.lblWins.fontSize = frmMainMenu.lblWins.fontSize + 1
    
    With frmMainMenu
        .cmdBlackjack.fontSize = .cmdBlackjack.fontSize + 1
        .cmdExit.fontSize = .cmdExit.fontSize + 1
        .cmdHelp.fontSize = .cmdHelp.fontSize + 1
        .cmdLeaderboard.fontSize = .cmdLeaderboard.fontSize + 1
        .cmdSettings.fontSize = .cmdLeaderboard.fontSize + 1
    End With

End Sub

Private Sub smallFont()

    cmdResetBalance.fontSize = 12
    cmdMenu.fontSize = 13
    lblFont.fontSize = 8
    lblColour.fontSize = 8
    choiceFont(0).fontSize = 8
    choiceFont(1).fontSize = 8
    choiceColour(0).fontSize = 8
    choiceColour(1).fontSize = 8
    lblFullName.fontSize = 8
    lblBalance1.fontSize = 8
    lblAchievements.fontSize = 8
    frameDetails.fontSize = 12
    
    frmMainMenu.lblFullName.fontSize = 10
    frmMainMenu.lblBalance1.fontSize = 10
    frmMainMenu.lblBlackjacks.fontSize = 10
    frmMainMenu.lblWins.fontSize = 10
    
    With frmMainMenu
        .cmdBlackjack.fontSize = 14
        .cmdExit.fontSize = 14
        .cmdHelp.fontSize = 14
        .cmdLeaderboard.fontSize = 14
        .cmdSettings.fontSize = 14
    End With

End Sub

Private Sub saveDetails()

    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(user)
        Put #1, userNumber, user 'This saves the record to the file index which matches the value of userNumber (i.e matching user that was chosen to be logged in)
    Close #1

End Sub
