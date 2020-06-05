VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   840
      ScaleHeight     =   6465
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   720
      Width           =   6615
      Begin VB.ListBox lstLeaderboard 
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
         ForeColor       =   &H000000FF&
         Height          =   4755
         ItemData        =   "frmHelp.frx":0000
         Left            =   240
         List            =   "frmHelp.frx":0002
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
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
         Height          =   1080
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3120
         Width           =   1365
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
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "   Rules of Blackjack/How to Play"
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
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMenu_Click()
    frmHelp.Hide
    Load frmMainMenu
    frmMainMenu.Show
    Unload frmHelp
End Sub

Private Sub Form_Load()
    
    Dim tempUser As typeUserDetails
    Dim helpInstruction As String
    
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(tempUser)
    
    Get #1, frmMainMenu.userNumber, tempUser
    
    If tempUser.fontSize = "Large" Then
        Me.cmdMenu.fontSize = Me.cmdMenu.fontSize + 1
        Me.lstLeaderboard.fontSize = Me.lstLeaderboard.fontSize + 1
    End If
    
    Close #1
    
    Open App.Path & "/Data Files/HelpInstructions.csv" For Input As #1
    
    Do
        Input #1, helpInstruction
        Me.lstLeaderboard.AddItem helpInstruction
    Loop Until EOF(1)
    
    Close #1
    
End Sub
