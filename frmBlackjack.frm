VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmBlackjack 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackjack Game"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Add Player"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   9020
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7350
      Width           =   855
   End
   Begin VB.Frame frameDealer 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Dealer's Hand"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   9240
      TabIndex        =   8
      Top             =   280
      Width           =   1455
      Begin VB.Label lblDealerHand 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   330
         Left            =   650
         TabIndex        =   9
         Top             =   360
         Width           =   150
      End
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Height          =   735
      Index           =   3
      Left            =   4600
      ScaleHeight     =   705
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   5290
      Width           =   1900
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Height          =   735
      Index           =   5
      Left            =   9000
      ScaleHeight     =   705
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   5290
      Width           =   1900
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Height          =   735
      Index           =   4
      Left            =   6800
      ScaleHeight     =   705
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   5290
      Width           =   1900
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Height          =   735
      Index           =   2
      Left            =   2450
      ScaleHeight     =   705
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   5290
      Width           =   1900
   End
   Begin VB.PictureBox picToolbar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   2500
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   6060
      Width           =   11415
      Begin VB.CommandButton cmdHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   9850
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1280
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Hit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7350
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   9850
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdBegin 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Double"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdStand 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Stand"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdHit 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Hit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdNextPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.ListBox lstPlayerDetails 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1605
         ItemData        =   "frmBlackjack.frx":0000
         Left            =   4150
         List            =   "frmBlackjack.frx":0002
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
      Begin VB.Image imgChip50 
         Height          =   735
         Left            =   480
         Picture         =   "frmBlackjack.frx":0004
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   735
      End
      Begin VB.Image imgChip25 
         Height          =   735
         Left            =   1200
         Picture         =   "frmBlackjack.frx":B453
         Stretch         =   -1  'True
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgChip10 
         Height          =   735
         Left            =   480
         Picture         =   "frmBlackjack.frx":169C2
         Stretch         =   -1  'True
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgChip5 
         Height          =   735
         Left            =   1200
         Picture         =   "frmBlackjack.frx":20EFA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Image imgChip1 
         Height          =   735
         Left            =   480
         Picture         =   "frmBlackjack.frx":2B59A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Height          =   735
      Index           =   1
      Left            =   280
      ScaleHeight     =   705
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   5280
      Width           =   1900
   End
   Begin WMPLibCtl.WindowsMediaPlayer loseSound 
      Height          =   15
      Left            =   1560
      TabIndex        =   21
      Top             =   480
      Width           =   15
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
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
      _cx             =   26
      _cy             =   26
   End
   Begin WMPLibCtl.WindowsMediaPlayer blackjackSound 
      Height          =   15
      Left            =   480
      TabIndex        =   20
      Top             =   600
      Width           =   15
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
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
      _cx             =   26
      _cy             =   26
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   7
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   6
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Blackjack Pays 3 to 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Left            =   4530
      TabIndex        =   7
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   4
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   5
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   6
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   7
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgDealerHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   7
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   6
      Left            =   10200
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   5
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   4
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   9240
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   7
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   6
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   5
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   4
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   6800
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   7
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   6
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   5
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   4
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   4600
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   7
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   5
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   4
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   2460
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   11400
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   5
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   4
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   3
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   2
      Left            =   480
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgHand1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   300
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   11400
      Y1              =   6030
      Y2              =   6030
   End
   Begin VB.Label lblDealer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Dealer Stands on 17"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Left            =   4635
      TabIndex        =   6
      Top             =   1320
      Width           =   2505
   End
End
Attribute VB_Name = "frmBlackjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim usersHands(1 To 5) As typeUsersHand 'This array holds the data about each players playing hand
Dim cardDeck(1 To 208) As typePlayingCards  'This array stores 4 decks of cards and all the data about them
Dim players(1 To 5) As typeUserDetails 'This array holds the details about all the playing players
Dim leaderboardPlayers(1 To 5) As typeLeaderboardDetails 'This array will be used to hold the data which will be stored to the Leaderboard File
Dim playerCount As Integer  'Keeps track of how many players are in the game
Dim currentPlayer As Integer    'Tracks which players turn it currently is
Dim dealersHand(1 To 7) As typePlayingCards 'Stores the cards that the dealer has
Dim dealerHandValue As Integer 'Stores the value of the dealers hand
Dim dealerSoftHand As Boolean 'Tracks if the dealer has a soft hand
Dim dealerCardCount As Integer 'Tracks the number of cards the dealer has
Dim cardCount As Integer        'Tracks the number of cards that have been dealt out. So the program knows when to shuffle/the array position
Dim gameInProgress As Boolean   'Tracks if there is currently a game being played
Dim bettingInProgress As Boolean 'Tracks if the game is in the betting stage
Dim cardStage As Boolean 'Tracks if the game has reached the stage when the player can hit or stand
Dim payoutStage As Boolean 'Tracks if the game has reached the final stage of the game


Private Sub Form_Load()
    
    Dim index As Integer
    Dim cardLoop As Integer
    
    Call loadDeck 'Calls the sub-routine which loads 4 decks of cards from the related file
    Call shuffleDeck 'Calls the sub-routine which randomizes the order of the deck array, simulating a shuffle
    
    'Initializes the number of players to 1
    playerCount = 1
    
    'Initializes the game stage booleans to false
    gameInProgress = False
    bettingInProgress = False
    cardStage = False
    payoutStage = False
    
    'Disables the unnecessary buttons
    cmdHelp.Enabled = False
    cmdNextPlayer.Enabled = False
    cmdHit.Enabled = False
    cmdStand.Enabled = False
    cmdDouble.Enabled = False
    
    'Assigns the file paths for the two required sounds
    loseSound.URL = App.Path & "/Sound Effects/Lost Money.wav"
    blackjackSound.URL = App.Path & "/Sound Effects/Got Black Jack.wav"
    
    'Loops through all additional players to reset their details
    For index = 2 To 5
        'All fields refer to the usersHands(index) variable
        With usersHands(index)
            'Only iterates for the number of cards which were stored in the users hand previously
            For cardLoop = 1 To .cardCount
                .usersCards(cardLoop).cardImage = ""
                .usersCards(cardLoop).playingValue = 0
                .usersCards(cardLoop).rank = ""
                .usersCards(cardLoop).suit = ""
            Next cardLoop
            
            'Resets all variables stored in the other players records
            .betAmount = 0
            .blackjack = False
            .bust = False
            .cardCount = 0
            .handValue = 0
            .softHand = False
        End With
        
    Next index
    
    'Opens the userDetails file and loads the user which was logged in.
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(players(1))
        Get #1, frmMainMenu.userNumber, players(1)
    Close #1
    
    If players(1).fontSize = "Large" Then
        With Me
            .cmdAddPlayer.fontSize = .cmdAddPlayer.fontSize + 1
            .cmdBegin.fontSize = .cmdBegin.fontSize + 1
            .cmdDouble.fontSize = .cmdDouble.fontSize + 1
            .cmdExit.fontSize = .cmdExit.fontSize + 1
            .cmdHelp.fontSize = .cmdHelp.fontSize + 1
            .cmdHit.fontSize = .cmdHit.fontSize + 1
            .cmdNextPlayer.fontSize = .cmdNextPlayer.fontSize + 1
            .cmdStand.fontSize = .cmdStand.fontSize + 1
            .lblDealerHand.fontSize = .lblDealerHand.fontSize + 2
            .lblStatus.fontSize = .lblStatus.fontSize + 2
            .lblDealer.fontSize = .lblDealer.fontSize + 2
        End With
    End If
    
    If players(1).colourScheme = "Red  " Then
        Me.BackColor = &HC0&
        Me.lblStatus.BackColor = &HC0&
        Me.lblDealer.BackColor = &HC0&
        Me.lblDealerHand.BackColor = &HC0&
        Me.frameDealer.BackColor = &HC0&
    End If
    
    
    
End Sub

Private Sub cmdBegin_Click()
'This procedure runs when the user wishes to begin playing (Presses the Start Button)

    'This code runs when a game isn't already in progress
    If gameInProgress = False Then
    
        'Assigns the boolean track variables to appropriate values
        gameInProgress = True
        bettingInProgress = True
        cardStage = False
        payoutStage = False
        
        'Ensures the start and Add Player buttons are disabled and that the
        'Next Player button is re-enabled
        cmdBegin.Enabled = False
        cmdAddPlayer.Enabled = False
        cmdNextPlayer.Enabled = True
        
        'Initializes the values for the dealers variables
        dealerHandValue = 0
        dealerCardCount = 0
        dealerSoftHand = False
        
        'Where most of the deck has been used the order of it is shuffled
        If cardCount >= 166 Then
            Call shuffleDeck
        End If
        
        'Sets the current player to the first player
        currentPlayer = 1
        
        'Sets the status label to indicate to the first player it is their turn, shows dealers hand value to be 0
        lblStatus.Caption = shortenName(players(1)) & " please place your bets"
        lblDealerHand.Caption = "0"
        
        Call updatePlayerTag 'Calls procedure to update all players tags, name, initial bet amount, hand value
        Call refreshPlayerListBox 'Updates the list box to ensure current players are listed, with balance
        Call resetCards 'Calls procedure which clears the image boxes that store the card images
        
    End If
    
End Sub

Private Sub cmdAddPlayer_Click()
'This procedure fires when the user clicks on the add user button.
'It is responsible for letting the user add other players to the game

    Dim username As String * 20 'Both username and password are fixed length to match the length defined in the record structure
    Dim password As String * 20
    Dim confirmation As String  'Used for message box responses
    Dim index As Integer
    Dim existingUser As Boolean 'Decides if player the user wants to add is already playing
    Dim userFound As Boolean, usernameRetry As Boolean, passwordRetry As Boolean
    Dim tempUser As typeUserDetails 'A variable which temporarily holds user data from the user details file
    
    'Initializes the boolean variables to false
    userFound = False
    usernameRetry = False
    passwordRetry = False
    existingUser = False
    
    'Checks if the max number of players has been added
    If playerCount = 5 Then
        'Alerts the user they have reached the max number of players
        MsgBox "Sorry, You've Reached the Max Player Limit", vbOKOnly, "Max Player Limit"
        'Disables the Add Player Button
        cmdAddPlayer.Enabled = False
        'Halts execution of this particular sub procedure
        Exit Sub
    End If
    
    Call resetCards 'Calls the procedure which resets the game board
    Call updatePlayerTag  'Calls the procedure that rests the individual players tags
    
    'Opens the userDetails file identifying it as random access
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(tempUser)
        
    'Keeps looping while it hasn't found a matching user in the file or the user eants to keep retying
    While userFound = False Or usernameRetry = True
        
        Get #1, 1, tempUser 'Recieves first item from file, purely to reset the reading of the file to the beginning
        'Recieves desired username to search for from user
        username = InputBox("Enter the Username of the Player you Wish to Add:", "Enter Username")
        
        'Keeps looping until the username entered doesn't match an already playing user
        Do
            'Only loops up to the number of players which are currently playing
            For index = 1 To playerCount
                'Decides if the entered username matches an existign player
                If username = players(index).username Then
                    existingUser = True
                    confirmation = MsgBox("This User has Already Been Added. Do You Wish to Try Again?", vbYesNo, "Existing User")
                    
                    'If the user doesn't want to try again (value = 6), the first condition fires
                    If confirmation <> "6" Then
                        Close #1    'File closes
                        Exit Sub    'The add player sub procedure is ended
                    Else
                        'User will re-enter another username to add
                        username = InputBox("Enter the Username of the Player you Wish to Add:", "Enter Username")
                        index = 0
                    End If
                    
                Else
                    'If an existing user isn't found then the variable is set to false
                    existingUser = False
                End If
            Next index
        Loop Until existingUser = False
        
        index = 1   'Resets the index variable to 1
        
        'Will keep looping until a user is found or the end of the file is reached
        Do Until userFound = True Or tempUser.userNumber < 1
            Get #1, index, tempUser 'Recieves user details from current index position
            index = index + 1   'Increments index variable
            
            'Checks if the current users username matches the entered one
            If tempUser.username = username Then
                userFound = True
                usernameRetry = False
            End If
        Loop
            
        'Decides if the username doesn't match an existing user
        If userFound = False Then
            'Allows the user to try again if they wish
            confirmation = MsgBox("User Not Found. Do you Wish to Retry?", vbYesNo, "User Not Found")
            
            'If the user wishes to try again then the retry boolean is set to true
            If confirmation = "6" Then
                usernameRetry = True
            Else
                'Otherwise the file is closed and sub procedure ended
                Close #1
                Exit Sub
            End If
        End If
    Wend
    
    'Keeps iterating while the user wishes to retry passwords or the passwords don't match
    While passwordRetry = True Or password <> tempUser.password
        'Takes in the password from the user
        password = InputBox("Enter the Password for the User:", "Enter Password")
        
        'Decides if the entered password matches the found user
        If password = tempUser.password Then
            'If passwords match then player count increases
            playerCount = playerCount + 1
            players(playerCount) = tempUser 'Matching player added to next position in player array
            passwordRetry = False
            
            'If there are now 5 players then the button is disabled
            If playerCount = 5 Then
                cmdAddPlayer.Enabled = False
            End If
            
            'Player tags now updated to include newly added player
            Call updatePlayerTag
            
        Else
            'If don't match then user is given choice to try again
            confirmation = MsgBox("Password Incorrect. Do you wish to retry?", vbYesNo, "Incorrect Password")
            
            'If player wants to retry then retry variable is set to true
            If confirmation = "6" Then
                passwordRetry = True
            Else
                'Otherwise the file is closed and sub procedure ended
                Close #1
                Exit Sub
            End If
            
        End If
    Wend
            
       
    Close #1  'Closes opened file

End Sub

Private Sub cmdHit_Click()
'Executes when player presses the Hit button
'During play, allows user to recieve another card

    'Players can only double on two cards so any hitting will disable this button
    cmdDouble.Enabled = False
    
    'Checks that the player has less than 7 cards (no of image boxes i've made available)
    If usersHands(currentPlayer).cardCount < 7 Then
    
        Call dealCard   'Calls procedure to deal out a card to the player
        Call decideOutcome  'Calls procedure to check if the player has busted
        
    End If
    
End Sub

Private Sub cmdDouble_Click()
'Sub procedure fires when the user presses the Double Button
'Purpose is the double the player's bet and they will recieve only 1 extra card

    'If the player has enough money in their bank balance to cover doubling their bet
    'then the first statement fires
    If players(currentPlayer).balance >= usersHands(currentPlayer).betAmount Then
        'Decreases the players balance by the value of the original bet
        players(currentPlayer).balance = players(currentPlayer).balance - usersHands(currentPlayer).betAmount
        'Doubles the players bet
        usersHands(currentPlayer).betAmount = 2 * usersHands(currentPlayer).betAmount

        Call updatePlayerTag    'Updates the players tag with this new info
        Call refreshPlayerListBox   'Refreshes the list box with new balance info
        Call dealCard   'Deals out one new card to the player
        
        'If statements to decide if, by doubling, the player has busted themselves
        If usersHands(currentPlayer).handValue > 21 And usersHands(currentPlayer).softHand = True Then
        'WIth a soft hand and a value over 21 this drops the hand value by 10 and resets it to a hard hand
            usersHands(currentPlayer).softHand = False
            usersHands(currentPlayer).handValue = usersHands(currentPlayer).handValue - 10
        ElseIf usersHands(currentPlayer).handValue > 21 And usersHands(currentPlayer).softHand = False Then
        'Without a soft hand and a value over 21 then the player busts
            usersHands(currentPlayer).bust = True
            loseSound.Controls.play
        End If
    
        Call advancePlayer  'Advances play to the next player
    Else
        'With a lack of money, a message box alerts the user of this
        MsgBox "Insufficient Funds!", vbInformation, "Low Funds!"
    End If

End Sub

Private Sub cmdStand_Click()
'Sub procedure fires when the user presses the Stand Button

    'Advances play to the next player
    Call advancePlayer
    
End Sub

Private Sub cmdHelp_Click()

    'Variable to hold the suggested play option
    Dim suggestion As String
    
    'The first condition fires whereby the user has a hard hand (no aces)
    If usersHands(currentPlayer).softHand = False Then
        
        'A select case statement will fire the condition with parameters which match the
        'value stored in the usersHands(currentPlayer).handValue field
        Select Case usersHands(currentPlayer).handValue
            Case 4, 5, 6, 7, 8  'Fires for a hand value of 4-8 - Hit
                suggestion = "Hit"
            Case 9  'Fires for hand value = 9
                If cmdDouble.Enabled = True And dealersHand(1).playingValue <= 6 Then
                    'If the player can double and dealer's showing card 6 or less then Double
                    suggestion = "Double Down"
                Else
                    'Otherwise suggestion = hit
                    suggestion = "Hit"
                End If
            Case 10, 11 'Fires for hand value 10/11
                If usersHands(currentPlayer).handValue > dealerHandValue And cmdDouble.Enabled = True Then
                'If the players hand value is greater than dealers and can double then they should
                    suggestion = "Double Down"
                Else
                    suggestion = "Hit"
                End If
            Case 12, 13, 14, 15, 16 'Fires for hand value 12 - 16
                If dealersHand(1).playingValue >= 2 And dealersHand(1).playingValue <= 6 Then
                'If dealer's up-card 2-6 then fire
                    suggestion = "Stand"
                ElseIf dealersHand(1).playingValue >= 7 And dealersHand(1).playingValue <= 11 Then
                'If dealers up-card 7-A then fire
                    suggestion = "Hit"
                End If
            Case 17, 18, 19, 20, 21 'Fires for hand value 17-21
                suggestion = "Stand"
            Case Default    'Fires if, for some reason, no value matches one of the conditions
                suggestion = "Play it Safe"
        End Select
    ElseIf usersHands(currentPlayer).softHand = True Then
    'Second condition fires where the user has a soft hand (has aces)
    
        'A select case statement will fire the condition with parameters which match the
        'value stored in the usersHands(currentPlayer).handValue field
        Select Case usersHands(currentPlayer).handValue
            Case 13, 14, 15 'Fires for hand value 13-15
                suggestion = "Hit"
            Case 16, 17, 18 'Fires for hand value 16-18
                If cmdDouble.Enabled = True And dealersHand(1).playingValue <= 6 Then
                'Fires where double button is enbaled and dealers up-card <=6
                    suggestion = "Double Down"
                ElseIf cmdDouble.Enabled = False And usersHands(currentPlayer).handValue = 18 Then
                'Fires where double button is disabled and hand value = 18
                    suggestion = "Stand"
                Else
                    suggestion = "Hit"
                End If
            Case 19, 20, 21 'Fires for hand value 19-21
                suggestion = "Stand"
            Case Default 'Fires if, for some reason, no value matches one of the conditions
                suggestion = "Play it Safe"
        End Select
    End If
    
    'Updates the status label to show what the program's suggestion is
    lblStatus.Caption = "The Best Option Would be to " & suggestion

End Sub

Private Sub cmdNextPlayer_Click()
'Fires when the player preses the Next Button
'The purpose is to deal with advancing through the players for the betting phase

    If bettingInProgress = True Then
    'Only executes if the game is in the betting phase
        
        'Checks to makes ure the player has made a bet
        If usersHands(currentPlayer).betAmount = 0 Then
            MsgBox "You must place a bet to continue", vbExclamation, "Bet Required!"
        Else
        'If they have made a bet then this code executes
        
            If currentPlayer = playerCount Then
            'Where the the current player matches the number of players, all players placed bet
                'Updates game stage variables, disables the next button
                bettingInProgress = False
                cardStage = True
                cmdNextPlayer.Enabled = False
                
                Call initialCardSetup 'Calls procedure to setup the gametable
                
                currentPlayer = 1   'Resets player counter back to first player
                
                'Checks if the first player has a blackjack
                If usersHands(currentPlayer).blackjack = True Then
                'If they do then avance to next player and update player tags
                    Call advancePlayer
                    Call updatePlayerTag
                Else
                'Otherwise update status label to inform current player its their turn
                    lblStatus.Caption = shortenName(players(currentPlayer)) & " please take your turn"
                End If
            ElseIf currentPlayer < playerCount Then
                'If last player hasn't been reached yet then current player count is advanced
                currentPlayer = currentPlayer + 1
                'Status label is updated to inform current player to place bet
                lblStatus.Caption = shortenName(players(currentPlayer)) & " please place your bets"
            End If
        End If
        
    End If
    
    
End Sub

Private Sub loadDeck()
'This sub-routine loads the data about the playing cards from the related csv file

    Dim fileLoop As Integer
    Open App.Path & "/Data Files/PlayingCards.csv" For Input As #1
        'The program loops 52 times, once for each different card in the deck
        For fileLoop = 1 To 52
            'Below each separate field is loaded into the array for the deck
            'The last three lines makes three copies of each card, thus achieving 4 decks
            Input #1, cardDeck(fileLoop).suit, cardDeck(fileLoop).rank, cardDeck(fileLoop).cardImage, cardDeck(fileLoop).playingValue
            cardDeck(fileLoop + 52) = cardDeck(fileLoop)
            cardDeck(fileLoop + 104) = cardDeck(fileLoop)
            cardDeck(fileLoop + 156) = cardDeck(fileLoop)
        Next fileLoop

    Close #1
    

End Sub

Private Sub shuffleDeck()
'This sub-routine randomizes the order of the deck

    'Two random numbers are used to pick two random locations in the array to swap
    Dim randomNumber1 As Integer, randomNumber2 As Integer
    Dim tempCard As typePlayingCards
    Dim index As Integer
    
    cardCount = 0
    
    'The program loops 1664 times as this provides enough passes to sufficiently shuffle the deck
    For index = 1 To 1664
        Randomize 'This command seeds the randomize function
        randomNumber1 = Int((Rnd * 207) + 1)    'Chooses a random number between 1 and 208
        randomNumber2 = Int((Rnd * 207) + 1)
        
        'This section swaps the two randomly chosen array positions
        tempCard = cardDeck(randomNumber1)
        cardDeck(randomNumber1) = cardDeck(randomNumber2)
        cardDeck(randomNumber2) = tempCard
        
    Next index
    
    'For testing purposes the order of the shuffled deck is outputted for developers purposes
    Open App.Path & "/shuffledCardList.csv" For Output As #1
    For index = 1 To 208
        Write #1, cardDeck(index).suit, cardDeck(index).rank, cardDeck(index).playingValue, cardDeck(index).cardImage
    Next index
    Close #1
    

End Sub

Private Sub dealCard()
'This sub procedure deals with dealing out a card to the current player

    Select Case currentPlayer
    'Based on which players turn it is the Select structure will fire the most appropriate case statement
        Case 1
            cardCount = cardCount + 1   'Total card count incremented
            usersHands(1).cardCount = usersHands(1).cardCount + 1   'Current players card count incremented
            usersHands(1).usersCards(usersHands(1).cardCount) = cardDeck(cardCount) 'Current item in deck array added to current position in players hand array
            usersHands(1).handValue = usersHands(1).handValue + cardDeck(cardCount).playingValue 'Current Players hand value increased
            imgHand1(usersHands(currentPlayer).cardCount) = LoadPicture(App.Path & cardDeck(cardCount).cardImage) 'Appropriate card image loaded into image box from file path
            imgHand1(usersHands(currentPlayer).cardCount).Visible = True    'Image box is made visible to display card
        Case 2
            cardCount = cardCount + 1
            usersHands(2).cardCount = usersHands(2).cardCount + 1
            usersHands(2).usersCards(usersHands(2).cardCount) = cardDeck(cardCount)
            usersHands(2).handValue = usersHands(2).handValue + cardDeck(cardCount).playingValue
            imgHand2(usersHands(currentPlayer).cardCount) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
            imgHand2(usersHands(currentPlayer).cardCount).Visible = True
        Case 3
            cardCount = cardCount + 1
            usersHands(3).cardCount = usersHands(3).cardCount + 1
            usersHands(3).usersCards(usersHands(3).cardCount) = cardDeck(cardCount)
            usersHands(3).handValue = usersHands(3).handValue + cardDeck(cardCount).playingValue
            imgHand3(usersHands(currentPlayer).cardCount) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
            imgHand3(usersHands(currentPlayer).cardCount).Visible = True
        Case 4
            cardCount = cardCount + 1
            usersHands(4).cardCount = usersHands(4).cardCount + 1
            usersHands(4).usersCards(usersHands(4).cardCount) = cardDeck(cardCount)
            usersHands(4).handValue = usersHands(4).handValue + cardDeck(cardCount).playingValue
            imgHand4(usersHands(currentPlayer).cardCount) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
            imgHand4(usersHands(currentPlayer).cardCount).Visible = True
        Case 5
            cardCount = cardCount + 1
            usersHands(5).cardCount = usersHands(5).cardCount + 1
            usersHands(5).usersCards(usersHands(5).cardCount) = cardDeck(cardCount)
            usersHands(5).handValue = usersHands(5).handValue + cardDeck(cardCount).playingValue
            imgHand5(usersHands(currentPlayer).cardCount) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
            imgHand5(usersHands(currentPlayer).cardCount).Visible = True
    End Select
            
    Call softHandCheck  'Calls subprocedure that handles aces being dealt
    Call updatePlayerTag    'Updatres the players individual tags
    
End Sub

Private Sub decideOutcome()
'Sub procedure which check the outcome from a dealt card
'It mainly checks for when the users hand value exceeds 21, equates to 21 or the user has been dealt 7 cards

    If usersHands(currentPlayer).handValue > 21 And usersHands(currentPlayer).softHand = True Then
    'If the users hand value exceeds 21 when they have a soft hand then their hand value is decremented by 10
        usersHands(currentPlayer).softHand = False
        usersHands(currentPlayer).handValue = usersHands(currentPlayer).handValue - 10
        Call updatePlayerTag
    ElseIf usersHands(currentPlayer).handValue > 21 And usersHands(currentPlayer).softHand = False Then
    'If the users hand value exceeds 21 with a hard hand then they bust, lose sound plays, play advances to next player
        usersHands(currentPlayer).bust = True
        loseSound.Controls.play
        Call updatePlayerTag
        Call advancePlayer
    ElseIf usersHands(currentPlayer).handValue = 21 Then
    'If users hand value = 21 then play advances to next player
        Call advancePlayer
    ElseIf usersHands(currentPlayer).cardCount = 7 Then
    'On a card count of 7 the play advances to next player
        Call updatePlayerTag
        Call advancePlayer
    End If

End Sub

Private Sub softHandCheck()
'Sub procedure that ensures when aces are dealt out they are handled correctly


    If usersHands(currentPlayer).softHand = False Then
    'Fires if the user doesn't currently have a soft hand
        If usersHands(currentPlayer).usersCards(usersHands(currentPlayer).cardCount).playingValue = 11 And (usersHands(currentPlayer).handValue <= 21) Then
        'Fires if value of the most recent card dealt to user is 11 (ace) and players hand value <=21
            usersHands(currentPlayer).softHand = True   'Hand becomes soft
        ElseIf usersHands(currentPlayer).usersCards(usersHands(currentPlayer).cardCount).playingValue = 11 And (usersHands(currentPlayer).handValue > 21) Then
        'Fires is value of the most recent card dealt to user is 11 (ace) and their hand value exceeds 21
            usersHands(currentPlayer).handValue = usersHands(currentPlayer).handValue - 10 'Hand value drops by 10
        End If
    ElseIf usersHands(currentPlayer).softHand = True Then
    'Fires if the user has a soft hand
        If usersHands(currentPlayer).usersCards(usersHands(currentPlayer).cardCount).playingValue = 11 And (usersHands(currentPlayer).handValue > 21) Then
         'Fires if the most recently dealt card is an ace and the users hand value exceeds 21
            usersHands(currentPlayer).handValue = usersHands(currentPlayer).handValue - 10 'Hand value drops by 10
        End If
    End If
End Sub

Private Sub updatePlayerTag()
'This sub procedure deals with refreshing the player's indiviual tags

    Dim index As Integer
    
    'Iterates through each player who is playing
    For index = 1 To playerCount
        
        'Fires if the player has 7 cards and hasn't busted, telling them they've won
        If usersHands(index).cardCount = 7 Then
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "7 Card Trick. You Win!"
            picPlayer(index).Print "Payout: £" & 2 * usersHands(index).betAmount
        ElseIf usersHands(index).blackjack = True Then
        'Fires if the player has a blackjack, telling they have a blackjack and have won
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "BlackJack! You Win!"
            picPlayer(index).Print "Payout: £" & 2.5 * usersHands(index).betAmount
        ElseIf payoutStage = True And usersHands(index).handValue = dealerHandValue And usersHands(index).bust = False Then
        'Fires if it is currently the payout stage, players hand value matches the dealers and hasn't busted
        'Tells them no one has won
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "It's a Push. Neither Win or Lose!"
            picPlayer(index).Print "Payout: £0"
        ElseIf payoutStage = True And usersHands(index).bust = False And (usersHands(index).handValue < dealerHandValue And dealerHandValue <= 21) Then
        'Fires if it is the payout stage, user hasn't busted and user's hand value is less than dealer's
        'Tells them theyve lost
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "Hand:" & usersHands(index).handValue & " - LOSE!"
            picPlayer(index).Print "Lost: £" & usersHands(index).betAmount
        ElseIf usersHands(index).bust = True Then
        'Fires if the player has busted, tells them they have busted
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "Hand:" & usersHands(index).handValue & " - BUST!"
            picPlayer(index).Print "Lost: £" & usersHands(index).betAmount
        ElseIf payoutStage = True Then
        'Fires if it is the payout stage and have managed to beat the dealer's hand value
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "Hand:" & usersHands(index).handValue & " - You've Won!"
            picPlayer(index).Print "Payout: £" & 2 * usersHands(index).betAmount
        ElseIf usersHands(index).bust = False Then
        'If all other conditions fail then it updates the player on their current bet and hand value
            picPlayer(index).Cls
            picPlayer(index).Print players(index).username
            picPlayer(index).Print "Hand:" & usersHands(index).handValue
            picPlayer(index).Print "Bet: £" & usersHands(index).betAmount
        End If
    Next index
    
End Sub

Private Sub updatePlayerBets(ByVal chipValue As Integer)
'This sub procedure handles the updataing of player bets
'The parameter holds the value that the player is adding on to his bet
'It comes from the click event of image boxes representing playing chips

    If bettingInProgress = True Then
    'Only fires if the game has begun and the game is currently in the betting phase
    
        If (usersHands(currentPlayer).betAmount + chipValue) > 250 Then
        'If, by adding on the value of chip to current bet, the limit of 250 is exceeded this fires
            MsgBox "Sorry, Max Bet is 250", vbInformation, "Max Bet Reached!"
        ElseIf players(currentPlayer).balance >= chipValue Then
        'Fires if the value of balance can cover adding on the value of the chip
            'Increments players bet amount
            usersHands(currentPlayer).betAmount = usersHands(currentPlayer).betAmount + chipValue
            'Decrements players balance by chip value
            players(currentPlayer).balance = players(currentPlayer).balance - chipValue
            
            Call updatePlayerTag
            Call refreshPlayerListBox
        Else
        'Fires if balance is less than chip value
            MsgBox "Insufficient Funds.", vbExclamation, "Low Funds"
        End If
    ElseIf gameInProgress = False Then
    'When a game isn't in progress it informs the user to press start.
        MsgBox "Game Not in Progress. Please press Start", vbExclamation, "No Current Game"
    End If

End Sub

Private Sub refreshPlayerListBox()
'This sub procedure refreshes the listbox with players names and balance

    Dim index As Integer
    
    lstPlayerDetails.Clear 'Clears the listbox first
    
    'Iterates through each player currently playing and re-adds their details
    For index = 1 To playerCount
        lstPlayerDetails.AddItem shortenName(players(index)) & " - £" & players(index).balance
    Next index

End Sub

Private Sub initialCardSetup()
'This sub procedure is incharge of setting up the playing board after the betting stage

    Dim playerIndex As Integer, cardIndex As Integer
    
    'All buttons needed for play are enabled
    cmdHelp.Enabled = True
    cmdHit.Enabled = True
    cmdStand.Enabled = True
    cmdDouble.Enabled = True
    
    'As each player needs 2 cards, it loops twice to give each player the right number
    For cardIndex = 1 To 2
        'This iterates through each player currently playing
        For playerIndex = 1 To playerCount
        
            'The select statement fires the case with a condition which matches the value of playerIndex
            'This is needed to make sure the cards are added to the correct array of image boxes
            Select Case playerIndex
                Case 1
                    cardCount = cardCount + 1 'The total card count is incremented
                    usersHands(1).cardCount = usersHands(1).cardCount + 1 'The players card count is incremented
                    usersHands(1).usersCards(usersHands(1).cardCount) = cardDeck(cardCount) 'Current item in deck array added to current position in players hand array
                    usersHands(1).handValue = usersHands(1).handValue + cardDeck(cardCount).playingValue 'Current Players hand value increased
                    imgHand1(cardIndex) = LoadPicture(App.Path & cardDeck(cardCount).cardImage) 'Appropriate card image loaded into image box from file path
                    imgHand1(cardIndex).Visible = True 'Image box is made visible to display card
                Case 2
                    cardCount = cardCount + 1
                    usersHands(2).cardCount = usersHands(2).cardCount + 1
                    usersHands(2).usersCards(usersHands(2).cardCount) = cardDeck(cardCount)
                    usersHands(2).handValue = usersHands(2).handValue + cardDeck(cardCount).playingValue
                    imgHand2(cardIndex) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
                    imgHand2(cardIndex).Visible = True
                Case 3
                    cardCount = cardCount + 1
                    usersHands(3).cardCount = usersHands(3).cardCount + 1
                    usersHands(3).usersCards(usersHands(3).cardCount) = cardDeck(cardCount)
                    usersHands(3).handValue = usersHands(3).handValue + cardDeck(cardCount).playingValue
                    imgHand3(cardIndex) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
                    imgHand3(cardIndex).Visible = True
                Case 4
                    cardCount = cardCount + 1
                    usersHands(4).cardCount = usersHands(4).cardCount + 1
                    usersHands(4).usersCards(usersHands(4).cardCount) = cardDeck(cardCount)
                    usersHands(4).handValue = usersHands(4).handValue + cardDeck(cardCount).playingValue
                    imgHand4(cardIndex) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
                    imgHand4(cardIndex).Visible = True
                Case 5
                    cardCount = cardCount + 1
                    usersHands(5).cardCount = usersHands(5).cardCount + 1
                    usersHands(5).usersCards(usersHands(5).cardCount) = cardDeck(cardCount)
                    usersHands(5).handValue = usersHands(5).handValue + cardDeck(cardCount).playingValue
                    imgHand5(cardIndex) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
                    imgHand5(cardIndex).Visible = True
            End Select
                    
        Next playerIndex
        
        cardCount = cardCount + 1
        dealerCardCount = dealerCardCount + 1
        dealersHand(dealerCardCount) = cardDeck(cardCount) 'Adds the next card into the dealers hand
        
    Next cardIndex
    
    imgDealerHand(1) = LoadPicture(App.Path & dealersHand(1).cardImage) 'First dealer card is loaded from file path
    dealerHandValue = dealerHandValue + dealersHand(1).playingValue 'Dealers hand value is increased
    lblDealerHand.Caption = dealerHandValue 'Label with dealer hand value is updated
    imgDealerHand(2) = LoadPicture(App.Path & "/Playing Cards/card_back.jpg") 'To make is anonymous the second card displays the back of a card
    
    'Iterates through each of the players currently playing
    For playerIndex = 1 To playerCount
        
        If (usersHands(playerIndex).usersCards(1).playingValue + usersHands(playerIndex).usersCards(2).playingValue) = 21 Then
        'Fires if the player has a blackjack (i.e value of 21)
            usersHands(playerIndex).blackjack = True
            blackjackSound.Controls.play 'Plays the sound for recieving blackjack
        ElseIf usersHands(playerIndex).usersCards(1).playingValue = 11 Or usersHands(playerIndex).usersCards(2).playingValue = 11 Then
        'Fires if one of the players cards are an ace
            usersHands(playerIndex).softHand = True
        Else
        'Fires otherwise
            usersHands(playerIndex).softHand = False
        End If
        
    Next playerIndex
    
    Call updatePlayerTag
    
End Sub

Private Sub advancePlayer()
'This sub procedure is used to advance play to the next player
'Or if the last player has had their turn then it handles the final stage of the game
'and saving all details to file

    If currentPlayer < playerCount Then
    'Fires if the current player is less than the number of people playing
    
        cmdDouble.Enabled = True
        
        'Loops until the current player doesn't have a blackjack
        'This is to advance past any blackjack holding players
        Do
            'Increments currentPlayer counter
            currentPlayer = currentPlayer + 1
        
            'If the currentPlayer counter exceeds 5 then the procedure is called again
            If currentPlayer > 5 Then
                Call advancePlayer
            End If
        Loop Until usersHands(currentPlayer).blackjack = False
        
        'The status label updates to inform the current player to take their turn
        lblStatus.Caption = shortenName(players(currentPlayer)) & " please take your turn"
        
    ElseIf currentPlayer = 5 And usersHands(5).blackjack = True Then
    'Fires if it is the last players turn and he has a blackjack
        currentPlayer = 6 'Gives the counter a value which will force the else condition to be fired
        Call advancePlayer
    Else
    'If other conditions fail then this will fire
    
        'currentPlayer counter is reset, game tracking booleans are updated
        currentPlayer = 1
        cardStage = False
        payoutStage = True
        
        'The hit, stand and double buttons are disabled and status label has it caption reset
        cmdHit.Enabled = False
        cmdStand.Enabled = False
        cmdDouble.Enabled = False
        lblStatus.Caption = "Blackjack Pays 3 to 2"
        
        Call dealersTurn     'Calls procedure which plays the dealers turn
        Call payoutBets      'Calls the procedure which deals with paying bets
        Call updatePlayerTag 'Updates players individual tags
        Call endGame         'Calls procedure which handles ending the game
        Call saveDetails     'Calls the procedure which saves details to user file
        
    End If

End Sub

Private Sub dealersTurn()
'This sub procedure deals with processing the dealers turn

    'Initializes the dealers hand to soft
    dealerSoftHand = False
    
    'The second dealers card is given the appropriate image
    imgDealerHand(2) = LoadPicture(App.Path & dealersHand(2).cardImage)
    'Dealers hand value is increased
    dealerHandValue = dealerHandValue + dealersHand(2).playingValue
    lblDealerHand.Caption = dealerHandValue
    
    If dealersHand(1).playingValue = 11 Or dealersHand(2).playingValue = 11 Then
    'Fires if either of the dealers cards are an ace (value = 11)
        dealerSoftHand = True
    End If
    
    'Keeps iterating while the dealer has less than 7 cards and his hand value is less than 17
    Do While dealerCardCount < 7 And dealerHandValue < 17
    
        'This block of code mimics the card deal function but is specific to the dealers details
        'So dealers card count is increased, next card in deck is loaded into next image box
        'and the hand value is updated
        dealerCardCount = dealerCardCount + 1
        cardCount = cardCount + 1
        imgDealerHand(dealerCardCount) = LoadPicture(App.Path & cardDeck(cardCount).cardImage)
        imgDealerHand(dealerCardCount).Visible = True
        dealerHandValue = dealerHandValue + cardDeck(cardCount).playingValue
        lblDealerHand.Caption = dealerHandValue
        
        'Fires if the card dealt is an ace and the dealer doesn't have a soft hand
        If cardDeck(cardCount).playingValue = 11 And dealerSoftHand = False Then
            dealerSoftHand = True
        End If
        
        If dealerHandValue > 21 And dealerSoftHand = True Then
        'Fires if the dealers hand value is > 21 and he has a soft hand
            dealerSoftHand = False
            dealerHandValue = dealerHandValue - 10 'Hand value dropped by 10
            lblDealerHand.Caption = dealerHandValue 'Dealer value label is updated
        ElseIf dealerHandValue > 21 And dealerSoftHand = False Then
        'Fires if the hand value exceeds 21 and the dealer doesn't have a soft hand
            lblStatus.Caption = "Dealer Busts!"
        End If
        
    Loop
End Sub

Private Sub payoutBets()
'This sub procedure handles paying out the correct bets to winning players

    Dim index As Integer
    
    'Iterates through each of the players who are playing
    For index = 1 To playerCount
    
        If usersHands(index).blackjack = True Then
        'Fires if current player has a blackjack
        'It pays out original bet plus 1.5 times the bet
            players(index).balance = players(index).balance + (2.5 * usersHands(index).betAmount)
            players(index).blackjackCount = players(index).blackjackCount + 1 'Blackjack count in incremented
            players(index).timesWon = players(index).timesWon + 1   'Number of times won is incremented
            players(index).moneyEarned = players(index).moneyEarned + (2.5 * usersHands(index).betAmount)
        ElseIf dealerHandValue > 21 And usersHands(index).bust = False Then
        'Fires if the dealers hand exceeds 21 and the curent player isn't bust
        'Pays out money equal to the money bet
            players(index).balance = players(index).balance + (2 * usersHands(index).betAmount)
            players(index).timesWon = players(index).timesWon + 1
            players(index).moneyEarned = players(index).moneyEarned + (2 * usersHands(index).betAmount)
        ElseIf usersHands(index).handValue > dealerHandValue And usersHands(index).bust = False Then
        'Fires if the users hand value is greater than dealers and they aren't bust
        'Pays out money equal to the money bet
            players(index).balance = players(index).balance + (2 * usersHands(index).betAmount)
            players(index).timesWon = players(index).timesWon + 1
            players(index).moneyEarned = players(index).moneyEarned + (2 * usersHands(index).betAmount)
        ElseIf usersHands(index).cardCount = 7 And usersHands(index).bust = False Then
        'Fires if the player has 7 cards and aren't bust
        'Pays out equal money to the mney bet
            players(index).balance = players(index).balance + (2 * usersHands(index).betAmount)
            players(index).timesWon = players(index).timesWon + 1
            players(index).moneyEarned = players(index).moneyEarned + (2 * usersHands(index).betAmount)
        ElseIf usersHands(index).handValue = dealerHandValue And usersHands(index).bust = False Then
        'Fires if the player's hand value matches the dealers hand value and also isn't bust
        'It smply returns the players bet to them
            players(index).balance = players(index).balance + usersHands(index).betAmount
        End If
        
        'Fires if the players blackjack count = 5, sets their profile so that it knows its achieved achievement 1
        If players(index).blackjackCount = 5 Then
            players(index).achievement1 = True
        End If
        
        'Fires if the player has earned more than 150 pounds, sets boolean that tracks achievement 2 to true
        If players(index).moneyEarned >= 150 Then
            players(index).achievement2 = True
        End If
        
    Next index
    
    Call refreshPlayerListBox

End Sub

Private Sub endGame()
'This sub procedure is used to prepare the game for ending, and to allow another game to start

    Dim index As Integer
    Dim cardLoop As Integer
    
    'Sets game tracking booleans to false
    gameInProgress = False
    payoutStage = False
    
    'Enables the Begin button but disables the help button
    cmdBegin.Enabled = True
    cmdHelp.Enabled = False
    
    'Fires if there are less than 5 players, if so it re-enables the add player button
    If playerCount < 5 Then
        cmdAddPlayer.Enabled = True
    End If
    
    'Iterates through each of the players that were playing
    For index = 1 To playerCount
        'The with statement means all fields refer to the given variable
        With usersHands(index)
        
            'Loops through each of the dealt out cards and resets their value
            For cardLoop = 1 To .cardCount
                .usersCards(cardLoop).cardImage = ""
                .usersCards(cardLoop).playingValue = 0
                .usersCards(cardLoop).rank = ""
                .usersCards(cardLoop).suit = ""
            Next cardLoop
            
            'Resets each of the values for the players hand
            .betAmount = 0
            .blackjack = False
            .bust = False
            .cardCount = 0
            .handValue = 0
            .softHand = False
            
        End With
        
        With players(index)
            leaderboardPlayers(index).userNumber = .userNumber
            leaderboardPlayers(index).fullName = .fullName
            leaderboardPlayers(index).timesWon = .timesWon
        End With
    Next index
    
    With dealersHand(cardLoop)
        'Iterates through each of the dealers cards and resets their value
        For cardLoop = 1 To 7
            .cardImage = ""
            .playingValue = 0
            .rank = ""
            .suit = ""
        Next cardLoop
        
    End With
    
    'Resets each of the dealers variables for a new game
    dealerCardCount = 0
    dealerHandValue = 0
    dealerSoftHand = False
    

End Sub

Private Sub saveDetails()
'This sub procedure is used to save back the users details to the userDetails.dat file

    Dim index As Integer
    
    'Both the userDetails and the leaderboard files are open
    Open App.Path & "/Data Files/UserDetails.dat" For Random As #1 Len = Len(players(1))
    Open App.Path & "/Data Files/LeaderboardDetails.dat" For Random As #2 Len = Len(leaderboardPlayers(1))
    
    'Iterates through each of the players
    For index = 1 To playerCount
        'Saves all the details about the player back to the file
        Put #1, players(index).userNumber, players(index)
        'Saves the name and times won back to the leaderboard file
        Put #2, players(index).userNumber, leaderboardPlayers(index)
    Next index
    
    Close #1
    Close #2

End Sub

Private Sub resetCards()
'This sub procedure is used to remove the cards from the game board

    Dim index As Integer
    
    'Iterates 7 times (i.e number of card slots)
    For index = 1 To 7
    'Each of the players cards are hidden
    'The dealers cards are loaded with an empty picture (clearing it)
        imgHand1(index).Visible = False
        imgHand2(index).Visible = False
        imgHand3(index).Visible = False
        imgHand4(index).Visible = False
        imgHand5(index).Visible = False
        imgDealerHand(index) = LoadPicture()
        
    Next index

End Sub

Private Sub cmdExit_Click()
'This sub procedure is involved with handling the exit of the blackjack game

    Dim confirmation As String
    Dim index As Integer
    
    If gameInProgress = True Then
    'Fires if the user clicks the button and a game is in progress
        
        'Program asks the user for confirmation of their action
        confirmation = MsgBox("Are You Sure You Wish to Exit During the Game?", vbYesNo, "Exit Game?")
        
        If confirmation = "6" Then
        'Fires if the user selected yes (confirmed their intention to leave)
            
            'Iterates through each of the players and returns their bet to their balance
            For index = 1 To playerCount
                players(index).balance = players(index).balance + usersHands(index).betAmount
            Next index
            
            Call endGame     'Calls the procedure to end the game
            Call saveDetails 'Calls the procedure to save user details
            
            'Hides and unloads the blackjack form and loads and displays the main menu form
            frmBlackjack.Hide
            Load frmMainMenu
            frmMainMenu.Show
            Unload frmBlackjack
        End If
    Else
    'Fires if game not in progress
        playerCount = 1
        
        Call endGame 'Calls procedure to end game (resets certain variables)
        Call refreshPlayerListBox
        
        'Hides and unloads the blackjack form and loads and displays the main menu form
        frmBlackjack.Hide
        Load frmMainMenu
        frmMainMenu.Show
        Unload frmBlackjack
    End If
    
End Sub

'The below sub procedures are for each of the different playing chips
'They each call the update player bets procedure and pass a chip value
'equivalent to the chip the represent

Private Sub imgChip1_Click()
    Call updatePlayerBets(1)
End Sub

Private Sub imgChip10_Click()
    Call updatePlayerBets(10)
End Sub

Private Sub imgChip25_Click()
    Call updatePlayerBets(25)
End Sub

Private Sub imgChip5_Click()
    Call updatePlayerBets(5)
End Sub

Private Sub imgChip50_Click()
    Call updatePlayerBets(50)
End Sub
