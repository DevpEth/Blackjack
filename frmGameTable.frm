VERSION 5.00
Begin VB.Form Black 
   BorderStyle     =   0  'None
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15705
   FillColor       =   &H00C0FFC0&
   ForeColor       =   &H00C0FFC0&
   Icon            =   "frmGameTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12960
   ScaleWidth      =   15705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrSplitStand2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   8280
   End
   Begin VB.Timer tmrSplitStand1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   8280
   End
   Begin VB.Timer tmr 
      Left            =   1320
      Top             =   7560
   End
   Begin VB.ListBox lstSpare 
      Height          =   900
      Left            =   5520
      TabIndex        =   23
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer tmrSplitMove3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   7560
   End
   Begin VB.Timer tmrSplitMove2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   7560
   End
   Begin VB.Timer tmrSplitFix 
      Interval        =   1
      Left            =   1080
      Top             =   6720
   End
   Begin VB.Timer tmrSplitMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   5880
   End
   Begin VB.CommandButton cmdSplit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Split"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   855
      Left            =   11280
      TabIndex        =   19
      Top             =   11280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstDeck 
      Height          =   11190
      ItemData        =   "frmGameTable.frx":2A348
      Left            =   3120
      List            =   "frmGameTable.frx":2A34A
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAskBetOK 
      Caption         =   "Bet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13200
      TabIndex        =   15
      Top             =   11520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdDouble 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Double"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrMoniD 
      Interval        =   1
      Left            =   14280
      Top             =   2040
   End
   Begin VB.CommandButton cmdDev 
      Caption         =   "Off"
      Height          =   855
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstDevDeck 
      Height          =   11190
      ItemData        =   "frmGameTable.frx":2A34C
      Left            =   2040
      List            =   "frmGameTable.frx":2A34E
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrTest 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   9960
   End
   Begin VB.Timer tmrSHDealerAni2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14880
      Top             =   1080
   End
   Begin VB.Timer tmrDeckReset 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   14880
      Top             =   2640
   End
   Begin VB.Timer tmrSHPlayerAni2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14880
      Top             =   480
   End
   Begin VB.Timer tmrCardFlip 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14880
      Top             =   2040
   End
   Begin VB.CommandButton cmdStand 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stand"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstDealer 
      Height          =   1110
      Left            =   3480
      TabIndex        =   5
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstPlayer 
      Height          =   1110
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdHit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Hit"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   480
   End
   Begin VB.Timer tmrSHDealerAni 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14280
      Top             =   1080
   End
   Begin VB.Timer tmrSHPlayerAni 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14280
      Top             =   480
   End
   Begin VB.Timer tmrSetBackground 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   5040
      MaskColor       =   &H00C0FFC0&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   5895
   End
   Begin VB.Label lblHand2Stat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hand2"
      Height          =   210
      Left            =   375
      TabIndex        =   26
      Top             =   11040
      Width           =   465
   End
   Begin VB.Label lblHand1Stat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hand1"
      Height          =   210
      Left            =   375
      TabIndex        =   25
      Top             =   10680
      Width           =   465
   End
   Begin VB.Label lbltest 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   210
      Left            =   840
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblHand 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hand 1"
      Height          =   210
      Left            =   465
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblBetAmountD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BetAmount"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   11160
      TabIndex        =   20
      Top             =   10200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblChipAmount 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   15000
      TabIndex        =   18
      Top             =   10680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblBetD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   15240
      TabIndex        =   16
      Top             =   10680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgChipBet 
      Height          =   855
      Index           =   0
      Left            =   8520
      Top             =   12000
      Width           =   1095
   End
   Begin VB.Image imgChip 
      Height          =   855
      Index           =   4
      Left            =   9720
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Image imgChip 
      Height          =   855
      Index           =   3
      Left            =   9720
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Image imgChip 
      Height          =   855
      Index           =   2
      Left            =   9720
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Image imgChip 
      Height          =   855
      Index           =   1
      Left            =   9720
      Top             =   11040
      Width           =   1095
   End
   Begin VB.Image imgChip 
      Height          =   855
      Index           =   0
      Left            =   9720
      Top             =   12000
      Width           =   1095
   End
   Begin VB.Shape shpBet 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   10920
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblMoniDisplay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "moni"
      Height          =   270
      Left            =   2640
      TabIndex        =   12
      Top             =   11640
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblDeckReset 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   14490
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8115
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblDealerCount 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   13050
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblPlayerCount 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   10770
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgDCard 
      Height          =   3495
      Index           =   1
      Left            =   12840
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image imgPCard 
      Height          =   3495
      Index           =   0
      Left            =   10200
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image imgMainDeck 
      Height          =   3495
      Left            =   12840
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BLACKJACK"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   150
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   2715
      TabIndex        =   1
      Top             =   1200
      Width           =   11805
   End
   Begin VB.Image imgTable 
      Height          =   2055
      Left            =   0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Black"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                    '(@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@(@
                    '@.                                                                           .@@
                    '@.     *%.                                                                   .@@
                    '@.    /@@&.                                                                  .@@
                    '@.  .#@(.&@*                                                                 .@@
                    '@. ,%@@@@@@@(.                                                               .@@
                    '@./&@%*   (@@#,                                                              .@@
                    '@..,,,.   ,,,,                                                               .@@
                    '@.     ./                                                                    .@@
                    '@.  ./%@@@#,                                                                 .@@
                    '@. (@@@@@@@@&,                                                               .@@
                    '@. /&@@@@@@@%,                                                               .@@
                    '@.    ./&,                                                                   .@@
                    '@.                                                                           .@@
                    '@.                                                                           .@@
                    '@.                                    .(.                                    .@@
                    '@.                                  *%@@@%*                                  .@@
                    '@.                               ,(&@@@@@@@&(.                               .@@
                    '@.                             /%@@@@@@@@@@@@@%*                             .@@
                    '@.                          .#@@@@@@@@@@@@@@@@@@@#.                          .@@
                    '@.                        ,#@@@@@@@@@@@@@@@@@@@@@@@#.                        .@@
                    '@.                       *%@@@@@@@@@@@@@@@@@@@@@@@@@%*                       .@@
                    '@.                      *%@@@@@@@@@@@@@@@@@@@@@@@@@@@%,                      .@@
                    '@.                      *#&@@@@@@@@@@@@@@@@@@@@@@@@@&#,                      .@@
                    '@.                       ,#@@@@@@@@@@&%@%&@@@@@@@@@&#,                       .@@
                    '@.                         .,/#%&%(*,.(@/.,*(%%%(/,.                         .@@
                    '@.                                  ,#@@@(.                                  .@@
                    '@.                                                                           .@@
                    '@.                                                                           .@@
                    '@.                                                                           .@@
                    '@.                                                                   /@(.    .@@
                    '@.                                                               ,#@@@@@@@%* .@@
                    '@.                                                               *@@@@@@@@@( .@@
                    '@.                                                                 *%@@@&/.  .@@
                    '@.                                                                   .(,     .@@
                    '@.                                                                           .@@
                    '@.                                                              *&@@#   (&@@(.@@
                    '@.                                                               ,#@@@@@@@%* .@@
                    '@.                                                                ./@# (&#.  .@@
                    '@.                                                                  ,@@@/    .@@
                    '@.                                                                   .&*     .@@
                    '@.                                                                           .@@
                    '#&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&#@

Option Explicit

Const Dealer As Integer = 1
Const Player As Integer = 2
Const faceUp As Integer = 1
Const faceDown As Integer = 2
Dim PlayerHand As Integer
Dim DealerHand As Integer
Dim CardSizeX As Currency
Dim CardSizeY As Currency
Dim MainDeckX As Integer
Dim MainDeckY As Integer
Dim NumofPlayerCards As Integer
Dim NumOfDealerCards As Integer
Dim b As Integer
Dim p As Integer
Dim HitBtnChecker As Integer
Dim EndOfStartHandD As Integer
Dim targetXP As Integer
Dim targetYP As Integer
Dim targetXD As Integer
Dim targetYD As Integer
Dim slopeP As Currency
Dim slopeD As Currency
Dim delayT As Integer
Dim retVal As Long
Dim SoundBuffer As String
Dim SoundOnHover As String
Dim CardFlip As Integer
Const DealerLimit As Integer = 17
Dim BlackChk As Integer
Dim DealerCountShow As Integer
Dim PTimerFix As Integer
Dim DTimerFix As Integer
Dim DeckResetTmr As Integer
Const PlayerBust As Integer = 2
Const DealerBust As Integer = 3
Const PlayerBlack As Integer = 4
Const DealerBlack As Integer = 5
Const BothBlack As Integer = 6
Const Push As Integer = 7
Const PlayerWin As Integer = 8
Const DealerWin As Integer = 9
Dim HoverCounter As Integer
Dim TopOfDeck As Integer
Dim DevChk As Boolean
Dim BetAmount As Integer
Dim HandStart As Integer
Dim KeyBinds(5) As Integer
Dim KeyPressEnabled As Boolean
Dim AskBetPause As Boolean
Dim Chip As Integer
Const brown As Integer = 1
Const red As Integer = 2
Const blue As Integer = 3
Const green As Integer = 4
Const Black As Integer = 5
Dim Bet As Long
Dim DoubleCheck As Boolean
Dim SplitCheck As Boolean
Dim SplitNumOfCards As Integer
Dim SplitMode As Boolean
Dim TurnOff As Boolean
Dim CurrentHand As Integer
Dim TravelDis As Currency
Dim FlipTravelDis As Currency
Dim Hand1 As Integer
Dim Hand2 As Integer

Private Sub cmdAskBetOK_Click()
    If Bet <> 0 Then
        lblBetD.Visible = False
        cmdAskBetOK.Visible = False
        cmdClear.Visible = False
        shpBet.Visible = False
        AskBetPause = True
        Dim x As Integer
        For x = 0 To 4
            imgChip(x).Visible = False
        Next x
        
        For x = 1 To Chip
            On Error Resume Next
            Unload imgChipBet(x)
        Next x
        
        For x = 1 To 5
            Unload lblChipAmount(x)
        Next x
        
        Chip = 0
    End If
End Sub

Private Sub cmdClear_Click()
    Dim x  As Integer
    For x = 1 To Chip
        On Error Resume Next
        Call imgChipBet_Click(x)
    Next x
End Sub

Private Sub cmdDev_Click()
    If cmdDev.Caption = "Off" Then
        cmdDev.Caption = "On"
        
        With lstPlayer(0)
            .Top = (8212 - lstPlayer(0).Height) - 100
            .Left = 9360
            .Visible = True
        End With
        
        On Error Resume Next
        
        With lstPlayer(1)
            .Top = (8212 - lstPlayer(0).Height) - 100
            .Left = 9360 - lstPlayer(0).Width - 100
            .Visible = True
        End With
        
        With lstDealer
            .Top = (128 + CardSizeY) + 100
            .Left = 13680 - lstDealer.Width
            .Visible = True
        End With
        
        lstDevDeck.Visible = True
        
    Else
        On Error Resume Next
        cmdDev.Caption = "Off"
        lstPlayer(0).Visible = False
        lstDealer.Visible = False
        lstDevDeck.Visible = False
        lstPlayer(1).Visible = False
    End If
End Sub

Private Sub cmdDouble_Click()
    Call GiveCard(Player, faceUp)
    Call HideAll
    
    Moni = Moni - Bet
    Bet = Bet * 2
    
    lblBetAmountD.Caption = "Bet: $" & Bet
    
    Call Pause(50)
    DealerCountShow = 1
    If SplitMode = True Then
        Call SplitStand
    Else
        Call PlayerStand
    End If
End Sub

Private Sub cmdDouble_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 
    If HoverCounter > 1 Then
        HoverCounter = 2
    Else
        retVal = sndPlaySound(SoundOnHover, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
    
    HoverCounter = HoverCounter + 1
    
End Sub

Private Sub cmdHit_Click()
              
    HitBtnChecker = 1
    Call HideAll
    DoubleCheck = True
    SplitCheck = True
    Call GiveCard(Player, faceUp)
    
End Sub

Private Sub cmdHit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If HoverCounter > 1 Then
        HoverCounter = 2
    Else
        retVal = sndPlaySound(SoundOnHover, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
    
    HoverCounter = HoverCounter + 1
End Sub

Private Sub cmdPlay_Click()

'start game
    cmdPlay.Visible = False
    lblTitle.Visible = False
    cmdQuit.Visible = False
    
    frmExtra.Show vbModal

    If gCheckStart Then
        Call ResetDeck 'initalizes deck
        Call StartHand 'gives player and dealer two cards each
    Else
        cmdPlay.Visible = True
        lblTitle.Visible = True
        cmdQuit.Visible = True
    End If

End Sub

Private Sub cmdPlay_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If HoverCounter > 1 Then
        HoverCounter = 2
    Else
        retVal = sndPlaySound(SoundOnHover, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
    
    HoverCounter = HoverCounter + 1
    
End Sub

Private Sub cmdQuit_Click()
    Unload Me
    End
End Sub

Private Sub cmdQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If HoverCounter > 1 Then
        HoverCounter = 2
    Else
        retVal = sndPlaySound(SoundOnHover, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
    
    HoverCounter = HoverCounter + 1
End Sub

Private Sub cmdSplit_Click()
'create a new listbox to hold one card each
'move leftmost card over to the left and starthand with single card
'when done with initail hand, move current cards somewhere and move leftmost cards to center
    
    SplitMode = True
    imgPCard(1).Left = Me.Width / 2 - imgPCard(1).Width / 2
    Call HideAll
    tmrSplitMove.Enabled = True
    Do
        If tmrSplitMove.Enabled = False Then Exit Do
            DoEvents
    Loop
    
    Load lstPlayer(1)
    
    lstPlayer(1).AddItem lstPlayer(0).List(1)
    lstPlayer(0).RemoveItem 1
    NumofPlayerCards = 1
    b = 1
    Call FixPlayerPos
    
    Moni = Moni - Bet
    Bet = Bet * 2
    lblBetAmountD.Caption = "Bet: $" & Bet
    
    With lblHand
        .ForeColor = vbWhite
        .Font = "Microsoft Himalaya"
        .FontSize = 20
        .Left = imgPCard(1).Left + imgPCard(1).Width / 2 - lblHand.Width / 2
        .Top = imgPCard(1).Top + imgPCard(1).Height + 50
        .Visible = True
        .ZOrder 0
    End With
    
    CurrentHand = 1
    
    Call GiveCard(Player, faceUp)
    Pause (50)
    Call ShowAll
End Sub

Private Sub cmdStand_Click()
        
    If SplitMode = True Then
        Call SplitStand
    Else
        DealerCountShow = 1
        Call PlayerStand
    End If
End Sub
Private Sub SplitStand()
    
    Select Case CurrentHand
        
        Case 1
            
            Hand1 = PlayerHand
            
            Call HideAll
            
            tmrSplitMove2.Enabled = True
            tmrSplitMove3.Enabled = True
            Do Until tmrSplitMove3.Enabled = False And tmrSplitMove2.Enabled = False
                DoEvents
            Loop
            
            Load imgPCard(50)
            With imgPCard(50)
                .Width = CardSizeX
                .Height = CardSizeY
                .Stretch = True
                .Picture = imgPCard(100).Picture
                .Top = imgPCard(100).Top
                .Left = imgPCard(100).Top
                .Visible = True
                .ZOrder 0
                .Tag = imgPCard(100).Tag
            End With
            Unload imgPCard(100)
            
            Dim x As Integer
            Dim NewIndex As Integer
            For x = 1 To NumofPlayerCards
                NewIndex = x + 100
                Load imgPCard(NewIndex)
                With imgPCard(NewIndex)
                    .Width = CardSizeX
                    .Height = CardSizeY
                    .Stretch = True
                    .Picture = imgPCard(x).Picture
                    .Top = imgPCard(x).Top
                    .Left = imgPCard(x).Top
                    .Visible = True
                    .ZOrder 0
                    .Tag = imgPCard(x).Tag
                End With
                Unload imgPCard(x)
            Next x
            
            Load imgPCard(1)
            With imgPCard(1)
                .Width = CardSizeX
                .Height = CardSizeY
                .Stretch = True
                .Picture = imgPCard(50).Picture
                .Top = imgPCard(50).Top
                .Left = imgPCard(50).Top
                .Visible = True
                .ZOrder 0
                .Tag = imgPCard(50).Tag
            End With
            Unload imgPCard(50)
        
            Dim gap As Currency
            For x = 1 To NumofPlayerCards
                imgPCard(x + 100).Left = Me.Width / 11 + gap
                gap = gap + Me.Width * 0.02170138
            Next x
            
            NumofPlayerCards = 1
            b = 1
            
            lstPlayer(0).AddItem (lstPlayer(1).List(0))
            lstPlayer(1).RemoveItem 0
            
            For x = 0 To lstPlayer(0).ListCount - 2
                lstPlayer(1).AddItem (lstPlayer(0).List(0))
                lstPlayer(0).RemoveItem 0
            Next x
            
            DoubleCheck = False
            
            TurnOff = False
            
            
            cmdStand.Caption = "Stand"
            
            lblHand.Caption = "Hand 2"
            tmrSplitFix.Enabled = True
            
            CurrentHand = 2
            
            Call FixPlayerPos
            
            Pause (20)
            Call GiveCard(Player, faceUp)
            Pause (40)
            
            Call PlayerCounter
            
            Call ShowAll
            
        Case 2
            
            Hand2 = PlayerHand
            
            lblPlayerCount.Visible = False
            Call HideAll
            tmrSplitStand1.Enabled = True
            Do Until tmrSplitStand1.Enabled = False
                DoEvents
            Loop
            Pause (20)
    
            CardFlip = 2
            tmrCardFlip.Enabled = True
            SoundBuffer = StrConv(LoadResData("FLIP1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            Do Until tmrCardFlip.Enabled = False
                DoEvents
            Loop
            
            Pause (20)
            
            DealerCountShow = 1
            With lblDealerCount
                .FontSize = 24
                .Caption = "Dealer Count: " & CStr(DealerHand)
                .Left = (Me.Width / 2) - (lblDealerCount.Width / 2)
                .Top = (imgDCard(1).Top + imgDCard(1).Height) + 200
                .Visible = True
                .ZOrder 0
            End With
            
            Dim wins As Integer
            If Hand1 > 21 And Hand2 > 21 Then
                With lblDisplay
                    .Caption = "Both Bust"
                    .FontBold = True
                    .FontSize = 70
                    .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                    .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                    .Visible = True
                    .ZOrder 0
                End With
                wins = -2
                Select Case wins
                    Case 1
                        Moni = Moni + Bet / 2
                    Case 2
                        Moni = Moni + Bet * 2
                End Select
                Pause (50)
                Call CleanTable
            ElseIf DealerHand > 21 And Hand1 <= 21 And Hand2 <= 21 Then
                With lblDisplay
                    .Caption = "Both Win"
                    .FontBold = True
                    .FontSize = 70
                    .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                    .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                    .Visible = True
                    .ZOrder 0
                End With
                wins = 2
                 Select Case wins
                Case 1
                    Moni = Moni + Bet / 2
                Case 2
                    Moni = Moni + Bet * 2
                End Select
                Pause (50)
                Call CleanTable
            End If
            Select Case DealerHand
        
                Case 2 To 16
tryA:
                
                    Call GiveCard(Dealer, faceUp)
                    Call Pause(40)
                    If DealerHand < 17 Then GoTo tryA
            
            End Select
            
            
            If Hand1 > 21 Then
                wins = wins - 1
                With lblHand2Stat
                    .Caption = "Bust"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(Int(lstPlayer(1).ListCount) + 100).Left + CardSizeX
                    .Top = imgPCard(1).Top
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf DealerHand > 21 Then
                wins = wins + 1
                With lblHand2Stat
                    .Caption = "Win"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(Int(lstPlayer(1).ListCount) + 100).Left + CardSizeX
                    .Top = imgPCard(1).Top
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf Hand1 = DealerHand Then
                With lblHand2Stat
                    .Caption = "Push"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(Int(lstPlayer(1).ListCount) + 100).Left + CardSizeX
                    .Top = imgPCard(1).Top
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf Hand1 > DealerHand Then
                wins = wins + 1
                With lblHand2Stat
                    .Caption = "Win"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(Int(lstPlayer(1).ListCount) + 100).Left + CardSizeX
                    .Top = imgPCard(1).Top
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf DealerHand > Hand1 Then
                With lblHand2Stat
                    .Caption = "Lose"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(Int(lstPlayer(1).ListCount) + 100).Left + CardSizeX
                    .Top = imgPCard(1).Top
                    .Visible = True
                    .ZOrder 0
                End With
            End If
            
            If Hand2 > 21 Then
                With lblHand1Stat
                    .Caption = "Bust"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(1).Left - lblHand1Stat.Width
                    .Top = imgPCard(1).Top + imgPCard(1).Height - lblHand1Stat.Height
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf DealerHand > 21 Then
                wins = wins + 2
                With lblHand1Stat
                    .Caption = "Win"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(1).Left - lblHand1Stat.Width
                    .Top = imgPCard(1).Top + imgPCard(1).Height - lblHand1Stat.Height
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf Hand2 = DealerHand Then
                With lblHand1Stat
                    .Caption = "Push"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(1).Left - lblHand1Stat.Width
                    .Top = imgPCard(1).Top + imgPCard(1).Height - lblHand1Stat.Height
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf Hand2 > DealerHand Then
                wins = wins + 1
                With lblHand1Stat
                    .Caption = "Win"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(1).Left - lblHand1Stat.Width
                    .Top = imgPCard(1).Top + imgPCard(1).Height - lblHand1Stat.Height
                    .Visible = True
                    .ZOrder 0
                End With
            ElseIf DealerHand > Hand2 Then
                With lblHand1Stat
                    .Caption = "Lose"
                    .FontBold = True
                    .FontSize = 54
                    .Left = imgPCard(1).Left - lblHand1Stat.Width
                    .Top = imgPCard(1).Top + imgPCard(1).Height - lblHand1Stat.Height
                    .Visible = True
                    .ZOrder 0
                End With
            End If
            
            Select Case wins
                Case 1
                    Moni = Moni + Bet / 2
                Case 2
                    Moni = Moni + Bet * 2
            End Select
            Pause (80)
            Call CleanTable
    End Select
 
End Sub
Private Sub cmdStand_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If HoverCounter > 1 Then
        HoverCounter = 2
    Else
        retVal = sndPlaySound(SoundOnHover, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
    
    HoverCounter = HoverCounter + 1
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        
    If KeyPressEnabled Then
        
        Select Case Chr$(KeyAscii)
            Case "h"
                cmdHit.Value = True
            Case "s"
                cmdStand.Value = True
            Case "d"
                cmdDouble.Value = True
        End Select
    End If
End Sub

Private Sub Form_Load()
        
    Unload frmExtra
    KeyPressEnabled = False
    
'set table image to form
    Me.WindowState = 2
    'Me.Width = 28800
    'Me.Height = 16200
    imgTable.Top = 0
    imgTable.Left = 0
    imgTable.Picture = LoadResPicture("TABLE", vbResBitmap) 'load table background image
    
    
    
'set background to table image
    tmrSetBackground.Enabled = True
    
'set chip pictures
    imgChip(0).Picture = LoadResPicture("CHIPBROWN", vbResBitmap)
    imgChip(1).Picture = LoadResPicture("CHIPRED", vbResBitmap)
    imgChip(2).Picture = LoadResPicture("CHIPBLUE", vbResBitmap)
    imgChip(3).Picture = LoadResPicture("CHIPGREEN", vbResBitmap)
    imgChip(4).Picture = LoadResPicture("CHIPBLACK", vbResBitmap)
    
    Dim x As Integer
    For x = 0 To 4
        imgChip(x).Visible = False
        imgChip(x).Stretch = True
        imgChip(x).Width = Me.Width / 10
        imgChip(x).Height = imgChip(x).Width
    Next x
    
'initialize variables
    b = 1
    p = 1
    Chip = 0
    NumOfDecks = 0
'load sound that plays on hover
    SoundOnHover = StrConv(LoadResData("BUTTONHOVER", "SOUND"), vbUnicode)
    
End Sub
 
Private Sub ResetDeck() 'gives starting values to all cards in deck()
    
    Randomize
    
    lstDeck.Clear
    
    Dim i As Integer
    Dim x As Integer
    
'give each index in deck() a value
    
    For i = 0 To NumOfDecks
        For x = 0 To 51
            lstDeck.AddItem x
        Next x
    Next i

'shuffle sequence
    Dim rand As Integer
    Dim strSwp As String
    
    For i = lstDeck.ListCount - 1 To 0 Step -1
        rand = Int(i * Rnd)
        strSwp = lstDeck.List(i)
        lstDeck.List(i) = lstDeck.List(rand)
        lstDeck.List(rand) = strSwp
    Next i
    
    For x = 0 To lstDeck.ListCount - 1
        lstDevDeck.AddItem ReadCard(x)
    Next x
    
'play sound of deck being shuffled and animation of being placed down
    imgMainDeck.Visible = False
    
    SoundBuffer = StrConv(LoadResData("SHUFFLE", "SOUND"), vbUnicode)
    retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    Call Pause(50)
    
    imgMainDeck.Left = (Me.Width - Me.Width / 17) - (imgMainDeck.Width / 2)
    imgMainDeck.Top = (Me.Height - Me.Height / 1.01)
    MainDeckX = imgMainDeck.Left
    MainDeckY = imgMainDeck.Top
    imgMainDeck.Picture = LoadResPicture("FACEDOWN", vbResBitmap)
    imgMainDeck.Visible = True
    
    SoundBuffer = StrConv(LoadResData("FLIP2", "SOUND"), vbUnicode)
    retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    
    Call Pause(30)
End Sub

Private Sub StartHand()
    
    If Moni <= 0 And Bet = 0 Then
        Dim a As Integer
        a = MsgBox("Kick Rocks", vbOKOnly + vbSystemModal)
        Unload Me
        End
    End If
    
    Call AskBet
    
    Do
        If AskBetPause = True Then Exit Do
            DoEvents
    Loop
    AskBetPause = False
    
    With lblBetAmountD
        .ForeColor = vbWhite
        .Caption = "Bet: $" & Bet
        .Font = "Microsoft Himalaya"
        .FontSize = 58
        .Left = lblMoniDisplay.Left + lblMoniDisplay.Width
        .Top = lblMoniDisplay.Top
        .Visible = True
        .ZOrder 0
    End With
    
    Call Pause(20)
    
    If lstDeck.ListCount <= 7 Then
        Call ResetDeck
        tmrDeckReset.Enabled = True
    End If
    
    
    
    
    
    
    ''''''''''''''''''''''''''''''''
    Call GiveCard(Player, faceUp)
    Call Pause(20)
    ''''''''''''''''''''''''''''''''
    
    
    
    ''''''''''''''''''''''''''''''''
    Call GiveCard(Dealer, faceUp)
    ''''''''''''''''''''''''''''''''
    
    
    
    ''''''''''''''''''''''''''''''''
    Call GiveCard(Player, faceUp)
    Call Pause(20)
    ''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''''
    Call GiveCard(Dealer, faceDown)
    Call Pause(20)
    SoundBuffer = StrConv(LoadResData("CARDSX", "SOUND"), vbUnicode)
    retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    ''''''''''''''''''''''''''''''''
    
End Sub

Private Sub GiveCard(op As Integer, State As Integer)
    
    Dim x As Integer
    
    
        SoundBuffer = StrConv(LoadResData("CARDSX", "SOUND"), vbUnicode)
        retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
        
        Select Case op 'check who is getting the card
        
            Case Player
                
                Call DisplayCard(ReadCard(lstDeck.List(0)), op, State)
                
                Select Case Int(lstDeck.List(0))
                    Case 0 To 3
                            lstPlayer(0).AddItem "2"
                    Case 4 To 7
                            lstPlayer(0).AddItem "3"
                    Case 8 To 11
                            lstPlayer(0).AddItem "4"
                    Case 12 To 15
                            lstPlayer(0).AddItem "5"
                    Case 16 To 19
                            lstPlayer(0).AddItem "6"
                    Case 20 To 23
                            lstPlayer(0).AddItem "7"
                    Case 24 To 27
                            lstPlayer(0).AddItem "8"
                    Case 28 To 31
                            lstPlayer(0).AddItem "9"
                    Case 32 To 47
                            lstPlayer(0).AddItem "10"
                    Case 48 To 51
                            lstPlayer(0).AddItem "A"
                End Select
            
            Case Dealer
            
                Call DisplayCard(ReadCard(lstDeck.List(0)), op, State)
                
                Select Case Int(lstDeck.List(0))
                    Case 0 To 3
                            lstDealer.AddItem "2"
                    Case 4 To 7
                            lstDealer.AddItem "3"
                    Case 8 To 11
                            lstDealer.AddItem "4"
                    Case 12 To 15
                            lstDealer.AddItem "5"
                    Case 16 To 19
                            lstDealer.AddItem "6"
                    Case 20 To 23
                            lstDealer.AddItem "7"
                    Case 24 To 27
                            lstDealer.AddItem "8"
                    Case 28 To 31
                            lstDealer.AddItem "9"
                    Case 32 To 47
                            lstDealer.AddItem "10"
                    Case 48 To 51
                            lstDealer.AddItem "A"
                End Select
                
        End Select
            
    lstDeck.RemoveItem 0
    
    lstDevDeck.Clear
    For x = 0 To lstDeck.ListCount - 1
        lstDevDeck.AddItem ReadCard(lstDeck.List(x))
    Next x
    
End Sub

Private Sub imgMainDeck_Click()
    cmdDev.Visible = True
End Sub

Private Sub imgTable_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    HoverCounter = 0
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrCardFlip_Timer()
    
    Dim Target As Integer
    
    Select Case CardFlip
        
        Case 2
            
            Target = (-(imgDCard(2).Height)) - 10
            
            If imgDCard(2).Top > Target Then
            
                imgDCard(2).Top = imgDCard(2).Top - FlipTravelDis
            Else
                CardFlip = 3
                imgDCard(2).Picture = LoadResPicture(imgDCard(2).Tag, vbResBitmap)
                SoundBuffer = StrConv(LoadResData("FLIP2", "SOUND"), vbUnicode)
                retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            End If
            
        Case 3
                  
            If imgDCard(2).Top = imgDCard(1).Top Then
                
                CardFlip = 2
                tmrCardFlip.Enabled = False
                
            Else
                imgDCard(2).Top = imgDCard(2).Top + FlipTravelDis
                
            End If
            
            
    End Select
End Sub

Private Sub tmrDeckReset_Timer()
    
    With lblDeckReset
        .Caption = "Deck Reset!"
        .Left = (imgMainDeck.Left + imgMainDeck.Width / 2) - (lblDeckReset.Width / 2)
        .Top = (imgMainDeck.Top + imgMainDeck.Height) + 200
        .Visible = True
    End With
    
    DeckResetTmr = DeckResetTmr + 1
    
    If DeckResetTmr > 100 Then
        DeckResetTmr = 0
        lblDeckReset.Visible = False
        tmrDeckReset.Enabled = False
    End If
    
End Sub

Private Sub tmrDelay_Timer()
    delayT = delayT + 1
End Sub

Private Sub tmrMoniD_Timer()
    lblMoniDisplay.Caption = "$" & CStr(Moni)
End Sub

Private Sub tmrSHDealerAni_Timer()

    If p = NumOfDealerCards + 1 Then
        Call FixDealerPos
        tmrSHDealerAni.Enabled = False
    Else
        
        If imgDCard(p).Left < targetXD Then
            
            If Left$(imgDCard(p).Tag, 1) = "H" Then
                imgDCard(p).Tag = Mid(imgDCard(p).Tag, 2)
                imgDCard(p).Picture = LoadResPicture("FACEDOWN", vbResBitmap)
            Else
                imgDCard(p).Picture = LoadResPicture(imgDCard(p).Tag, vbResBitmap)
            End If
            
            p = p + 1
            
        Else
            
            imgDCard(p).Left = imgDCard(p).Left - TravelDis
            
        End If
    End If
        
End Sub

Private Sub tmrSHDealerAni2_Timer()
    
    If p = NumOfDealerCards + 1 Then
        Call FixDealerPos
        tmrSHDealerAni2.Enabled = False
    Else
        
        If imgDCard(p).Left < targetXD Then
            
            If Left$(imgDCard(p).Tag, 1) = "H" Then
                imgDCard(p).Tag = Mid(imgDCard(p).Tag, 2)
                imgDCard(p).Picture = LoadResPicture("FACEDOWN", vbResBitmap)
            Else
                imgDCard(p).Picture = LoadResPicture(imgDCard(p).Tag, vbResBitmap)
            End If
            
            p = p + 1
        Else
            
            imgDCard(p).Left = imgDCard(p).Left - TravelDis
            
        End If
    End If
    
End Sub

Private Sub tmrSHPlayerAni_Timer()

    If b = NumofPlayerCards + 1 Then
        Call FixPlayerPos
        tmrSHPlayerAni.Enabled = False
        
    Else
    
        If imgPCard(b).Top > targetYP And imgPCard(b).Left < targetXP Then
            
            
            imgPCard(b).Picture = LoadResPicture(imgPCard(b).Tag, vbResBitmap)
            b = b + 1
            If SplitMode = True Then TurnOff = True
            

        Else
                
                imgPCard(b).Top = imgPCard(b).Top + ((slopeP * imgPCard(b).Left) + targetYP) / 5
                imgPCard(b).Left = imgPCard(b).Left - TravelDis

        End If
        
    End If
    
End Sub

Private Sub tmrSHPlayerAni2_Timer()
    

    If b = NumofPlayerCards + 1 Then
        Call FixPlayerPos
        tmrSHPlayerAni2.Enabled = False
        
    Else
        
        If imgPCard(b).Top > targetYP And imgPCard(b).Left < targetXP Then
        
            imgPCard(b).Picture = LoadResPicture(imgPCard(b).Tag, vbResBitmap)
            b = b + 1
            
            If SplitMode = True Then TurnOff = True
            
        Else
    
            imgPCard(b).Top = imgPCard(b).Top + ((slopeP * imgPCard(b).Left) + targetYP) / 5
            imgPCard(b).Left = imgPCard(b).Left - TravelDis
        
        End If
        
    End If
    

End Sub

Private Sub tmrSetBackground_Timer()
            
    CardSizeX = Me.Width * 0.11
    CardSizeY = Me.Height * 0.269675
    
'set positions
    MainDeckX = (Me.Width - CardSizeX) - (imgMainDeck.Width / 25)
    MainDeckY = (CardSizeX / 25)
    
    imgTable.Width = Me.Width
    imgTable.Height = Me.Height
    
    lblTitle.Left = Me.Width / 2 - lblTitle.Width / 2
    
    cmdPlay.Left = Me.Width / 2 - cmdPlay.Width / 2
    cmdPlay.Top = (Me.Height / 2 - cmdPlay.Height / 2) + 1000
    
    imgMainDeck.Left = (Me.Width - Me.Width / 17) - (imgMainDeck.Width / 2)
    imgMainDeck.Top = (Me.Height - Me.Height / 1.01)
    
    lstDevDeck.Left = (imgMainDeck.Left - lstDevDeck.Width) - 100
    lstDevDeck.Top = imgMainDeck.Top
    
    TravelDis = Me.Width * 0.01302
    
    FlipTravelDis = Me.Height * 0.01574
    
    
    tmrSetBackground.Enabled = False
  
End Sub

Private Function ReadCard(cardId) As String

    Select Case cardId
        Case 0
            ReadCard = "2C"
        Case 1
            ReadCard = "2D"
        Case 2
            ReadCard = "2H"
        Case 3
            ReadCard = "2S"
        Case 4
            ReadCard = "3C"
        Case 5
            ReadCard = "3D"
        Case 6
            ReadCard = "3H"
        Case 7
            ReadCard = "3S"
        Case 8
            ReadCard = "4C"
        Case 9
            ReadCard = "4D"
        Case 10
            ReadCard = "4H"
        Case 11
            ReadCard = "4S"
        Case 12
            ReadCard = "5C"
        Case 13
            ReadCard = "5D"
        Case 14
            ReadCard = "5H"
        Case 15
            ReadCard = "5S"
        Case 16
            ReadCard = "6C"
        Case 17
            ReadCard = "6D"
        Case 18
            ReadCard = "6H"
        Case 19
            ReadCard = "6S"
        Case 20
            ReadCard = "7C"
        Case 21
            ReadCard = "7D"
        Case 22
            ReadCard = "7H"
        Case 23
            ReadCard = "7S"
        Case 24
            ReadCard = "8C"
        Case 25
            ReadCard = "8D"
        Case 26
            ReadCard = "8H"
        Case 27
            ReadCard = "8S"
        Case 28
            ReadCard = "9C"
        Case 29
            ReadCard = "9D"
        Case 30
            ReadCard = "9H"
        Case 31
            ReadCard = "9S"
        Case 32
            ReadCard = "10C"
        Case 33
            ReadCard = "10D"
        Case 34
            ReadCard = "10H"
        Case 35
            ReadCard = "10S"
        Case 36
            ReadCard = "JC"
        Case 37
            ReadCard = "JD"
        Case 38
            ReadCard = "JH"
        Case 39
            ReadCard = "JS"
        Case 40
            ReadCard = "QC"
        Case 41
            ReadCard = "QD"
        Case 42
            ReadCard = "QH"
        Case 43
            ReadCard = "QS"
        Case 44
            ReadCard = "KC"
        Case 45
            ReadCard = "KD"
        Case 46
            ReadCard = "KH"
        Case 47
            ReadCard = "KS"
        Case 48
            ReadCard = "AC"
        Case 49
            ReadCard = "AD"
        Case 50
            ReadCard = "AH"
        Case 51
            ReadCard = "AS"
    End Select
End Function

Private Sub DisplayCard(cardFace As String, op As Integer, State As Integer)
    
    Select Case op
    
        Case Player
        
            Select Case State
            
                Case faceUp
                    
                    NumofPlayerCards = NumofPlayerCards + 1
                    
                    Load imgPCard(NumofPlayerCards) 'create a new player card
                    
                    With imgPCard(NumofPlayerCards)
                        .Width = CardSizeX
                        .Height = CardSizeY
                        .Stretch = True
                        .Picture = LoadResPicture("FACEDOWN", vbResBitmap)
                        .Top = MainDeckY
                        .Left = MainDeckX
                        .Visible = True
                        .ZOrder 0
                        .Tag = cardFace
                    End With
                    
                    targetXP = (Me.Width / 2) - (CardSizeX / 2)
                    targetYP = (Me.Height / 2) - (CardSizeY / 2)
                    slopeP = (targetYP - (imgPCard(NumofPlayerCards).Top - (imgPCard(NumofPlayerCards).Top / 2))) / (targetXP - (imgPCard(NumofPlayerCards).Left + (imgPCard(NumofPlayerCards).Left / 2)))
                    
                    Select Case PTimerFix
                        Case 0
                            tmrSHPlayerAni2.Enabled = False
                            tmrSHPlayerAni.Enabled = True
                        Case 1
                            tmrSHPlayerAni.Enabled = False
                            tmrSHPlayerAni2.Enabled = True
                    End Select
                End Select
                
        Case Dealer
        
            Select Case State
            
                Case faceUp
                    
                    NumOfDealerCards = NumOfDealerCards + 1
                    
                    Load imgDCard(NumOfDealerCards + 1) 'create a new dealer card
                    
                    With imgDCard(NumOfDealerCards)
                        .Width = CardSizeX
                        .Height = CardSizeY
                        .Stretch = True
                        .Picture = LoadResPicture("FACEDOWN", vbResBitmap)
                        .Top = MainDeckY
                        .Left = MainDeckX
                        .Visible = True
                        .ZOrder 0
                        .Tag = cardFace
                    End With
                    
                    targetXD = (Me.Width / 2) - (CardSizeX / 2)
                    targetYD = imgMainDeck.Top
                    
                    Select Case DTimerFix
                        Case 0
                            tmrSHDealerAni2.Enabled = False
                            tmrSHDealerAni.Enabled = True
                        Case 1
                            tmrSHDealerAni.Enabled = False
                            tmrSHDealerAni2.Enabled = True
                    End Select
                    
                Case faceDown
                
                    NumOfDealerCards = NumOfDealerCards + 1
                    
                    Load imgDCard(NumofPlayerCards + 1) 'create a new dealer card
                    
                    With imgDCard(NumofPlayerCards)
                        .Width = CardSizeX
                        .Height = CardSizeY
                        .Stretch = True
                        .Picture = LoadResPicture("FACEDOWN", vbResBitmap)
                        .Top = MainDeckY
                        .Left = MainDeckX
                        .Visible = True
                        .ZOrder 0
                        .Tag = "H" & cardFace
                    End With
                    
                    targetXD = (Me.Width / 2) - (CardSizeX / 2)
                    targetYD = imgMainDeck.Top
                    
                    Select Case DTimerFix
                        Case 0
                            tmrSHDealerAni2.Enabled = False
                            tmrSHDealerAni.Enabled = True
                        Case 1
                            tmrSHDealerAni.Enabled = False
                            tmrSHDealerAni2.Enabled = True
                    End Select
                    
            End Select
            
            
    End Select
End Sub

Private Sub FixPlayerPos()
    Dim i As Integer
    Dim totalWidth As Currency
    Dim overlap As Currency
    Dim counter As Integer
    
    counter = NumofPlayerCards
    
    If counter > 5 Then
        overlap = Me.Width * 0.0915798611
    Else
        overlap = Me.Width * 0.0325520833
    End If
    
    ' calculate the total width of all player cards with overlap
    For i = 1 To counter
        totalWidth = totalWidth + imgPCard(i).Width - overlap
    Next i
    ' add the overlap value back to the last card to avoid extra gap
    totalWidth = totalWidth + overlap

    ' calculate the starting position to center the cards
    Dim xPos As Integer
    xPos = (Me.Width - totalWidth) / 2

    ' Set the Left property of each card with overlap and center them
    For i = 1 To counter
        imgPCard(i).Left = xPos
        xPos = xPos + imgPCard(i).Width - overlap
    Next i
    
    If HitBtnChecker = 1 Then
        Call ShowAll
        HitBtnChecker = 0
    End If
    
    Call PlayerCounter
    
End Sub

Private Sub FixDealerPos()
    Dim x As Integer
    Dim totalWidth As Integer
    Dim overlap As Integer
    
    If NumOfDealerCards > 4 Then
        overlap = Me.Width * 0.0915798611
    Else
        overlap = Me.Width * 0.0325520833
    End If
    
    For x = 1 To NumOfDealerCards
        totalWidth = totalWidth + imgDCard(x).Width - overlap
    Next x
    
    totalWidth = totalWidth + overlap

    ' calculate the starting position to center the cards
    Dim xPos As Integer
    xPos = (Me.Width - totalWidth) / 2

    ' Set the Left property of each card with overlap and center them
    For x = 1 To NumOfDealerCards
        imgDCard(x).Left = xPos
        xPos = xPos + imgDCard(x).Width - overlap
    Next x
    
    If EndOfStartHandD <> 1 Then
    
                                            '''*****************************'''
                                            '''After All Four Cards Are Delt'''
                                            '''*****************************'''
                                            
        With cmdHit
            .Left = (imgMainDeck.Left + imgMainDeck.Width / 2) - (cmdHit.Width / 2)
            .Top = (imgMainDeck.Top + imgMainDeck.Height) + 250
        End With
        
        With cmdDouble
            .Height = cmdHit.Height
            .Width = cmdHit.Width
            .Top = (cmdHit.Top + cmdHit.Height) + 500
            .Left = cmdHit.Left
        End With
        
        With cmdSplit
            .Height = cmdHit.Height
            .Width = cmdHit.Width
            .Top = (cmdDouble.Top + cmdDouble.Height) + 500
            .Left = cmdHit.Left
        End With
        
        With cmdStand
            .Height = cmdHit.Height
            .Width = cmdHit.Width
            .Top = (cmdSplit.Top + cmdSplit.Height) + 500
            .Left = cmdSplit.Left
        End With
        
        KeyPressEnabled = True
        
        Call ShowAll
        
        EndOfStartHandD = 1
    Else
        'do nothing
    End If
    
    Call DealerCounter
    Call NaturalBlackChk
    
End Sub

Private Sub PlayerCounter()
    
    Dim x As Integer
    Dim i As Integer
    Dim NumOfAce As Integer
    Dim SoftCheck As Boolean
    
    PlayerHand = 0
    
    For x = 0 To lstPlayer(0).ListCount - 1
        
        If lstPlayer(0).List(x) = "A" Then
        
            NumOfAce = NumOfAce + 1
            
        Else
            PlayerHand = PlayerHand + Int(lstPlayer(0).List(x))
        End If
        
    Next x
    
    If NumOfAce > 0 Then
        
        For i = 1 To NumOfAce
            
            If PlayerHand + 11 > 21 Then
                PlayerHand = PlayerHand + 1
            Else
                If NumOfAce >= 2 And PlayerHand + 11 >= 21 Then
                    PlayerHand = PlayerHand + 1
                Else
                    PlayerHand = PlayerHand + 11
                    SoftCheck = True
                    'xxx
                End If
            End If
        Next i
        
    End If
    
    '''''''''''''''''''''''
    '''Player Bust Logic'''
    '''''''''''''''''''''''
    
    
    
    
    If PlayerHand > 21 Then
        If SplitMode = True Then
            Call HideAll
            With lblDisplay
                .Caption = "Bust"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("AWW", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            Call Pause(110)
            lblDisplay.Visible = False
            Pause (40)
            If PTimerFix = 1 Then
                PTimerFix = 0
            Else
            PTimerFix = 1
            End If
            
            Call SplitStand
        Else
            Call WinCondition(PlayerBust)
        End If
    Else
    
        With lblPlayerCount
            .ForeColor = vbWhite
            .Caption = "Your Count: " & CStr(PlayerHand)
            If SoftCheck = True And PlayerHand <> 21 Then .Caption = .Caption & "/" & CStr(PlayerHand - 10)
            .FontSize = 24
            .Left = (Me.Width / 2) - (lblPlayerCount.Width / 2)
            .Top = imgPCard(1).Top - lblPlayerCount.Height - 200
            .Visible = True
            .ZOrder 0
        End With
    End If
End Sub

Private Sub DealerCounter()

    Dim x As Integer
    Dim NumOfAce As Integer
    
    DealerHand = 0
    
    For x = 0 To lstDealer.ListCount - 1
        
        If lstDealer.List(x) = "A" Then
        
            NumOfAce = NumOfAce + 1
            
        Else
            DealerHand = DealerHand + Int(lstDealer.List(x))
        End If
        
    Next x
    
    If NumOfAce > 0 Then
        
        For x = 1 To NumOfAce
            
            If DealerHand + 11 > 21 Then
                DealerHand = DealerHand + 1
            Else
                If NumOfAce >= 2 And DealerHand + 11 >= 21 Then
                    DealerHand = DealerHand + 1
                Else
                    DealerHand = DealerHand + 11
                End If
            End If
        Next x
        
    End If
    
    Select Case DealerCountShow
    
        Case 0
        
        Case 1
        
            With lblDealerCount
                .ForeColor = vbWhite
                .Caption = "Dealer Count: " & CStr(DealerHand)
                .FontSize = 24
                .Left = (Me.Width / 2) - (lblDealerCount.Width / 2)
                .Top = (imgDCard(1).Top + imgDCard(1).Height) + 200
                .Visible = True
                .ZOrder 0
            End With
    End Select
    
End Sub

Private Sub PlayerStand()

'set table up for dealer drawing
    Call HideAll
    
    CardFlip = 2
    
'reval dealer's hidden card
    tmrCardFlip.Enabled = True
    SoundBuffer = StrConv(LoadResData("FLIP1", "SOUND"), vbUnicode)
    retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    Do Until tmrCardFlip.Enabled = False
        DoEvents
    Loop
    
    With lblDealerCount
        .FontSize = 24
        .Caption = "Dealer Count: " & CStr(DealerHand)
        .Left = (Me.Width / 2) - (lblDealerCount.Width / 2)
        .Top = (imgDCard(1).Top + imgDCard(1).Height) + 200
        .Visible = True
        .ZOrder 0
    End With

                                                    '****************'
                                                    '**Dealer Logic**'
                                                    '****************'

    
        Select Case DealerHand
        
            Case 2 To 16
tryA:
                
                Call GiveCard(Dealer, faceUp)
                Call Pause(40)
                If DealerHand < 17 Then GoTo tryA
            
        End Select
    
'Determine Winner
    If DealerHand > 21 Then
        Call WinCondition(DealerBust)
    ElseIf PlayerHand > DealerHand Then
        Call WinCondition(PlayerWin)
    ElseIf DealerHand > PlayerHand Then
        Call WinCondition(DealerWin)
    ElseIf DealerHand = PlayerHand Then
        Call WinCondition(Push)
    End If
End Sub

Private Sub WinCondition(Result As Integer)
    
    Select Case Result
        
'****************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Case PlayerBust
        
            Call HideAll
            
            With lblDisplay
                .ForeColor = vbRed
                .Caption = "Bust"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
            
            With lblPlayerCount
                .Caption = "You Bust With: " & CStr(PlayerHand)
                .Left = (Me.Width / 2) - (lblPlayerCount.Width / 2)
                .Top = imgPCard(1).Top - lblPlayerCount.Height - 200
                .Visible = True
                .ZOrder 0
            End With
            
            lblBetAmountD.ForeColor = vbRed
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("AWW", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(110)
            CardFlip = 2
            tmrCardFlip.Enabled = True
            SoundBuffer = StrConv(LoadResData("FLIP1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            Do Until tmrCardFlip.Enabled = False
                DoEvents
            Loop
            
            Call Pause(40)
            lblDisplay.Visible = False
            
            Call CleanTable
'****************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Case DealerBust
            
            Moni = Moni + Bet * 2
            
            With lblDisplay
            .ForeColor = vbGreen
            .Caption = "Dealer Bust!"
            .FontBold = True
            .FontSize = 70
            .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
            .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
            .Visible = True
            End With
            
            lblBetAmountD.ForeColor = vbGreen
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("WIN1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(80)
            lblDisplay.Visible = False
            Call CleanTable
        
        
            
'****************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        Case PlayerBlack
            
            Call HideAll
            
            Moni = Moni + Bet * 2
            
            With lblDisplay
                .ForeColor = vbGreen
                .Caption = "BlackJack!"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
            
            lblBetAmountD.ForeColor = vbGreen
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("BLACKW1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(500)
            CardFlip = 2
            tmrCardFlip.Enabled = True
            SoundBuffer = StrConv(LoadResData("FLIP1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            Do Until tmrCardFlip.Enabled = False
                DoEvents
            Loop
            
            Call Pause(40)
            lblDisplay.Visible = False
            Call CleanTable
            
'****************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              
        Case DealerBlack
        
           Call HideAll
            
            With lblDisplay
                .ForeColor = vbRed
                .Caption = "Dealer BlackJack!"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
            
            lblBetAmountD.ForeColor = vbRed
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("LOSE1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(100)
            CardFlip = 2
            tmrCardFlip.Enabled = True
            SoundBuffer = StrConv(LoadResData("FLIP1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            Do Until tmrCardFlip.Enabled = False
                DoEvents
            Loop
            
            Call Pause(80)
            lblDisplay.Visible = False
            
            Call CleanTable
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'***************************************************************************************************************************
        
        Case BothBlack
        
           Call HideAll
           
           Moni = Moni + Bet
            
            With lblDisplay
                .Caption = "Royal Push!"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
        
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("DRAW", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(50)
            lblDisplay.Visible = False
            Call CleanTable
        
'***************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Case Push
            
            Moni = Moni + Bet
            
            With lblDisplay
                .Caption = "Push"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("DRAW", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(70)
            lblDisplay.Visible = False
            Call CleanTable
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'***************************************************************************************************************************
        
        Case PlayerWin
            
            Moni = Moni + Bet * 2
            
            With lblDisplay
                .ForeColor = vbGreen
                .Caption = "You Win!"
                .FontBold = True
                .FontSize = 70
                .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
                .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
                .Visible = True
            End With
            
            lblBetAmountD.ForeColor = vbGreen
            
            Call Pause(10)
            SoundBuffer = StrConv(LoadResData("WIN1", "SOUND"), vbUnicode)
            retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
            
            Call Pause(80)
            lblDisplay.Visible = False
            Call CleanTable
            
'****************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Case DealerWin
        
        With lblDisplay
            .ForeColor = vbRed
            .Caption = "Dealer Wins!"
            .FontBold = True
            .FontSize = 70
            .Left = (Me.Width / 2) - (lblDisplay.Width / 2)
            .Top = (Me.Height / 2) - 2 * (lblDisplay.Height / 2)
            .Visible = True
        End With
        
        lblBetAmountD.ForeColor = vbRed
        
        Call Pause(10)
        SoundBuffer = StrConv(LoadResData("LOSE1", "SOUND"), vbUnicode)
        retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
        
        Call Pause(100)
        lblDisplay.Visible = False
        Call CleanTable
        
'****************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    End Select
    
End Sub

Private Sub NaturalBlackChk()
    
    Select Case BlackChk
    
        Case 0
    
            If PlayerHand = 21 And DealerHand = 21 Then
                
                Call WinCondition(BothBlack)
                
            ElseIf PlayerHand = 21 Then
                
                Call WinCondition(PlayerBlack)
                
            ElseIf DealerHand = 21 And CStr(lstDealer.List(0)) = "10" Then
                
                Call WinCondition(DealerBlack)
                
            End If
    
            BlackChk = 1
            
    End Select
    
End Sub

Private Sub CleanTable()
    
    Select Case PTimerFix
        Case 0
            PTimerFix = 1
        Case 1
            PTimerFix = 0
    End Select
    
    Select Case DTimerFix
        Case 0
            DTimerFix = 1
        Case 1
            DTimerFix = 0
    End Select
    
    Dim x As Integer
    
    On Error Resume Next
    For x = 1 To NumofPlayerCards + 1
        Unload imgPCard(x)
    Next x
    
    For x = 2 To NumOfDealerCards + 1
        Unload imgDCard(x)
    Next x
    
    For x = 0 To lstPlayer(1).ListCount
        Unload imgPCard(x + 100)
    Next x
    
    On Error Resume Next
    Unload lstPlayer(1)
    
    lblHand.Visible = False
    
    imgPCard(0).Visible = False
    imgDCard(1).Visible = False
    
    lblDealerCount.Visible = False
    lblPlayerCount.Visible = False
    
    lblBetAmountD.Visible = False
    
    lblHand1Stat.Visible = False
    lblHand2Stat.Visible = False
    
    Call CleanVars
    
    Call Pause(50)
    
    Call StartHand
    
End Sub

Private Sub CleanVars()

PlayerHand = 0
DealerHand = 0
NumofPlayerCards = 0
NumOfDealerCards = 0
p = 1
b = 1
HitBtnChecker = 0
EndOfStartHandD = 0
CardFlip = 2
BlackChk = 0
DealerCountShow = 0
lstPlayer(0).Clear
lstDealer.Clear
Bet = 0
Hand1 = 0
Hand2 = 0
DoubleCheck = False
SplitCheck = False
SplitMode = False
End Sub

Private Sub Pause(delay As Integer)

    delayT = 0
    tmrDelay.Enabled = True
    Do
        If delayT = delay Then Exit Do
            DoEvents
    Loop
    tmrDelay.Enabled = False

End Sub

Private Sub HideAll()
    KeyPressEnabled = False
    cmdHit.Visible = False
    cmdDouble.Visible = False
    cmdSplit.Visible = False
    cmdStand.Visible = False
    cmdQuit.Visible = False
    
End Sub

Private Sub ShowAll()

    cmdDouble.Enabled = True
    If Bet > Moni Or DoubleCheck = True Then
        cmdDouble.Enabled = False
    End If
    
    cmdSplit.Enabled = True
    If lstPlayer(0).List(0) <> lstPlayer(0).List(1) Or SplitCheck = True Or Bet > Moni Then cmdSplit.Enabled = False
    
    With cmdHit
        .Left = (imgMainDeck.Left + imgMainDeck.Width / 2) - (cmdHit.Width / 2)
        .Top = (imgMainDeck.Top + imgMainDeck.Height) + 250
    End With
    
    Dim h As Integer
    h = cmdHit.Height
    
    If SplitMode = True Then
        cmdDouble.Top = cmdHit.Top + h + 500
        cmdStand.Top = cmdHit.Top + 2 * h + 2 * 500
        KeyPressEnabled = True
        cmdHit.Visible = True
        cmdDouble.Visible = True
        cmdSplit.Visible = False
        cmdStand.Visible = True
        cmdQuit.Visible = True
    Else
        cmdDouble.Top = cmdHit.Top + h + 500
        cmdSplit.Top = cmdHit.Top + 2 * h + 2 * 500
        cmdStand.Top = cmdHit.Top + 3 * h + 3 * 500
        KeyPressEnabled = True
        cmdHit.Visible = True
        cmdDouble.Visible = True
        cmdSplit.Visible = True
        cmdStand.Visible = True
        cmdQuit.Visible = True
    End If
    
    If PlayerHand = 21 Then
        cmdHit.Enabled = False
        cmdDouble.Enabled = False
        cmdSplit.Enabled = False
    End If
    
End Sub

Private Sub AskBet()
    
    With lblMoniDisplay
        .AutoSize = False
        .Caption = "$" & Moni
        .Font = "Microsoft Himalaya"
        .FontSize = 58
        .Width = 4000
        .Height = 800
        .Left = 0
        .Top = (Me.Height - lblMoniDisplay.Height)
        .Visible = True
        .ZOrder 0
        .Left = Me.Width / 2 - lblMoniDisplay.Width / 2
    End With
    
    With shpBet
        .Width = Me.Width / 1.5
        .Height = Me.Height / 1.1
        .Top = (Me.Height / 2) - (shpBet.Height / 2)
        .Left = (Me.Width / 2) - (shpBet.Width / 2)
        .Visible = True
        .ZOrder 0
    End With
    
    With cmdAskBetOK
        .Visible = True
        .Left = Me.Width / 2 - cmdAskBetOK.Width / 2
        .Top = Me.Height / 2 - cmdAskBetOK.Height / 2
    End With
    
    With lblBetD
        .Caption = "Bet Amount: $" & Bet
        .Font = "Microsoft Himalaya"
        .FontSize = 58
        .Left = shpBet.Left + shpBet.Width / 2 - lblBetD.Width / 2
        .Top = shpBet.Top + shpBet / 1.5
        .Visible = True
        .ZOrder 0
    End With
    
    Dim x As Integer
    
    With imgChip(2)
        .Left = (Me.Width / 2) - (imgChip(2).Width / 2)
        .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) + (shpBet.Height / 4)
        .ZOrder 0
        .Visible = True
    End With
    
    imgChip(1).Left = imgChip(2).Left - imgChip(2).Width - 200
    imgChip(0).Left = imgChip(2).Left - imgChip(2).Width * 2 - 400
    imgChip(3).Left = imgChip(2).Left + imgChip(2).Width + 200
    imgChip(4).Left = imgChip(2).Left + imgChip(2).Width * 2 + 400
    
    For x = 0 To 4
        With imgChip(x)
            .ZOrder 0
            .Visible = True
            .Top = imgChip(2).Top
        End With
    Next x
    
    For x = 1 To 5
        Load lblChipAmount(x)
        With lblChipAmount(x)
            .Font = "Microsoft Himalaya"
            .FontSize = 40
            .Visible = False
        End With
    Next x
    
    With cmdClear
        .Left = imgChip(4).Left + imgChip(4).Width + 500
        .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) - (shpBet.Height / 4)
    End With
    
    Call BetOk(3, 0)
    
    cmdQuit.Visible = True
    
    'Call imgChip_Click(0)
End Sub

Private Sub imgChipBet_Click(Index As Integer)
    
    SoundBuffer = StrConv(LoadResData("POKERCHIP", "SOUND"), vbUnicode)
    retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    
    Select Case Int(imgChipBet(Index).Tag)
        Case 1
            If Int(lblChipAmount(1).Caption) = 1 Then
                lblChipAmount(1).Caption = "0"
                lblChipAmount(1).Visible = False
            Else
                lblChipAmount(1).Caption = Int(lblChipAmount(1).Caption) - 1
            End If
        
        Case 5
            If Int(lblChipAmount(2).Caption) = 1 Then
                lblChipAmount(2).Caption = "0"
                lblChipAmount(2).Visible = False
            Else
                lblChipAmount(2).Caption = Int(lblChipAmount(2).Caption) - 1
            End If
            
        Case 10
            If Int(lblChipAmount(3).Caption) = 1 Then
                lblChipAmount(3).Caption = "0"
                lblChipAmount(3).Visible = False
            Else
                lblChipAmount(3).Caption = Int(lblChipAmount(3).Caption) - 1
            End If
            
        Case 25
            If Int(lblChipAmount(4).Caption) = 1 Then
                lblChipAmount(4).Caption = "0"
                lblChipAmount(4).Visible = False
            Else
                lblChipAmount(4).Caption = Int(lblChipAmount(4).Caption) - 1
            End If
            
        Case 100
            If Int(lblChipAmount(5).Caption) = 1 Then
                lblChipAmount(5).Caption = "0"
                lblChipAmount(5).Visible = False
            Else
                lblChipAmount(5).Caption = Int(lblChipAmount(5).Caption) - 1
            End If
            
    End Select
    Select Case Index
        Case Is >= 1
            Bet = Bet - CLng(imgChipBet(Index).Tag)
            Moni = Moni + CLng(imgChipBet(Index).Tag)
            Unload imgChipBet(Index)
    End Select
    
    lblBetD.Caption = "Bet Amount: $" & Bet
    Call BetOk(2, Index)
End Sub

Private Sub imgChip_Click(Index As Integer)
        
    SoundBuffer = StrConv(LoadResData("POKERCHIP", "SOUND"), vbUnicode)
    retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    Select Case Index
            
        Case 0
            Call AssignChip(brown)
            Bet = Bet + 1
            Moni = Moni - 1
        Case 1
            Call AssignChip(red)
            Bet = Bet + 5
            Moni = Moni - 5
        Case 2
            Call AssignChip(blue)
            Bet = Bet + 10
            Moni = Moni - 10
        Case 3
            Call AssignChip(green)
            Bet = Bet + 25
            Moni = Moni - 25
        Case 4
            Call AssignChip(Black)
            Bet = Bet + 100
            Moni = Moni - 100
    End Select
    
    cmdClear.Visible = True
    
    lblBetD.Caption = "Bet Amount: $" & Bet
    
    Call BetOk(1, Index)
End Sub

Private Sub AssignChip(color As Integer)
    
    Chip = Chip + 1
    Load imgChipBet(Chip)
    
    Select Case color
    
        Case brown
            With imgChipBet(Chip)
                .Visible = False
                .Width = imgChip(2).Width
                .Height = imgChip(2).Height
                .Stretch = True
                .Picture = LoadResPicture("CHIPBROWN", vbResBitmap)
                .Left = imgChip(0).Left
                .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) - (shpBet.Height / 4)
                .Visible = True
                .ZOrder 0
                .Tag = "1"
            End With
            
            With lblChipAmount(1)
                .Caption = Int(lblChipAmount(1).Caption) + 1
                .Left = imgChipBet(Chip).Left + imgChipBet(Chip).Width / 2 - lblChipAmount(1).Width / 2
                .Top = imgChipBet(Chip).Top + imgChipBet(Chip).Height + 100
                .Visible = True
                .ZOrder 0
            End With
            
        Case red
            With imgChipBet(Chip)
                .Visible = False
                .Width = imgChip(2).Width
                .Height = imgChip(2).Height
                .Stretch = True
                .Picture = LoadResPicture("CHIPRED", vbResBitmap)
                .Left = imgChip(1).Left
                .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) - (shpBet.Height / 4)
                .Visible = True
                .ZOrder 0
                .Tag = "5"
            End With
            
            With lblChipAmount(2)
                .Caption = Int(lblChipAmount(2).Caption) + 1
                .Left = imgChipBet(Chip).Left + imgChipBet(Chip).Width / 2 - lblChipAmount(2).Width / 2
                .Top = imgChipBet(Chip).Top + imgChipBet(Chip).Height + 100
                .Visible = True
                .ZOrder 0
            End With
        
        Case blue
            With imgChipBet(Chip)
                .Visible = False
                .Width = imgChip(2).Width
                .Height = imgChip(2).Height
                .Stretch = True
                .Picture = LoadResPicture("CHIPBLUE", vbResBitmap)
                .Left = imgChip(2).Left
                .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) - (shpBet.Height / 4)
                .Visible = True
                .ZOrder 0
                .Tag = "10"
            End With
            
            With lblChipAmount(3)
                .Caption = Int(lblChipAmount(3).Caption) + 1
                .Left = imgChipBet(Chip).Left + imgChipBet(Chip).Width / 2 - lblChipAmount(3).Width / 2
                .Top = imgChipBet(Chip).Top + imgChipBet(Chip).Height + 100
                .Visible = True
                .ZOrder 0
            End With
            
        Case green
            With imgChipBet(Chip)
                .Visible = False
                .Width = imgChip(2).Width
                .Height = imgChip(2).Height
                .Stretch = True
                .Picture = LoadResPicture("CHIPGREEN", vbResBitmap)
                .Left = imgChip(3).Left
                .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) - (shpBet.Height / 4)
                .Visible = True
                .ZOrder 0
                .Tag = "25"
            End With
            
            With lblChipAmount(4)
                .Caption = Int(lblChipAmount(4).Caption) + 1
                .Left = imgChipBet(Chip).Left + imgChipBet(Chip).Width / 2 - lblChipAmount(4).Width / 2
                .Top = imgChipBet(Chip).Top + imgChipBet(Chip).Height + 100
                .Visible = True
                .ZOrder 0
            End With
            
        Case Black
            With imgChipBet(Chip)
                .Visible = False
                .Width = imgChip(2).Width
                .Height = imgChip(2).Height
                .Stretch = True
                .Picture = LoadResPicture("CHIPBLACK", vbResBitmap)
                .Left = imgChip(4).Left
                .Top = (shpBet.Top + shpBet.Height / 2 - imgChip(2).Height / 2) - (shpBet.Height / 4)
                .Visible = True
                .ZOrder 0
                .Tag = "100"
            End With
            
            With lblChipAmount(5)
                .Caption = Int(lblChipAmount(5).Caption) + 1
                .Left = imgChipBet(Chip).Left + imgChipBet(Chip).Width / 2 - lblChipAmount(5).Width / 2
                .Top = imgChipBet(Chip).Top + imgChipBet(Chip).Height + 100
                .Visible = True
                .ZOrder 0
            End With
            
    End Select
End Sub

Private Sub BetOk(op As Integer, chipIndex As Integer)
    
    Dim x As Integer
    
    Select Case op
        
        Case 1
        
            Select Case Moni
            
                Case Is < 1
                    For x = 0 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 5
                    For x = 1 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 10
                    For x = 2 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 25
                    For x = 3 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 100
                    imgChip(4).Visible = False
            End Select
            
        Case 2
            Select Case Moni
                Case Is >= 100
                    For x = 0 To 4
                        imgChip(x).Visible = True
                    Next x
                Case Is >= 25
                    For x = 0 To 3
                        imgChip(x).Visible = True
                    Next x
                Case Is >= 10
                    For x = 0 To 2
                        imgChip(x).Visible = True
                    Next x
                Case Is >= 5
                    For x = 0 To 1
                        imgChip(x).Visible = True
                    Next x
                Case Is >= 1
                    imgChip(0).Visible = True
            End Select
            
            If Bet = 0 Then cmdClear.Visible = False
            
        Case 3
            Select Case Moni
                Case Is < 5
                    For x = 1 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 10
                    For x = 2 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 25
                    For x = 3 To 4
                        imgChip(x).Visible = False
                    Next x
                Case Is < 100
                    imgChip(4).Visible = False
            End Select
    End Select
End Sub


Private Sub tmrSplitFix_Timer()
    If SplitMode = True And CurrentHand = 1 Then imgPCard(1).Top = targetYP + CardSizeY - Me.Height * 0.00177469
    If SplitMode = True And CurrentHand = 2 Then imgPCard(1).Top = targetYP + CardSizeY - Me.Height * 0.00177469
    If TurnOff = True Then tmrSplitFix.Enabled = False
End Sub

Private Sub tmrSplitMove_Timer()
    
    
    imgPCard(2).Left = imgPCard(2).Left - TravelDis
    
    If imgPCard(2).Left < Me.Width / 11 Then
        Load imgPCard(100)
        With imgPCard(100)
            .Width = CardSizeX
            .Height = CardSizeY
            .Stretch = True
            .Picture = imgPCard(2).Picture
            .Top = imgPCard(2).Top
            .Left = imgPCard(2).Left
            .Visible = True
            .ZOrder 0
            .Tag = imgPCard(2).Tag
        End With
        
        Unload imgPCard(2)
        
        tmrSplitMove.Enabled = False
    End If
End Sub


Private Sub tmrSplitMove2_Timer()
    Dim x As Integer
    
    If imgPCard(1).Left < Me.Width / 11 Then
        tmrSplitMove2.Enabled = False
    End If
    
    For x = 1 To NumofPlayerCards
        imgPCard(x).Left = imgPCard(x).Left - TravelDis
    Next x
    
End Sub

Private Sub tmrSplitMove3_Timer()
    If imgPCard(100).Left + imgPCard(100).Width / 2 > Me.Width / 2 - imgPCard(100).Width / 2 Then
        tmrSplitMove3.Enabled = False
    End If
    imgPCard(100).Left = imgPCard(100).Left + TravelDis
End Sub

Private Sub tmrSplitStand1_Timer()
    Dim x As Integer
    If imgPCard(1).Left > Me.Width / 2 + Me.Width / 4 Then
        Dim gap As Currency
        For x = 1 To NumofPlayerCards
            imgPCard(x).Left = Me.Width / 2 + Me.Width / 4 + gap
            gap = gap + Me.Width * 0.0217013889
        Next x
        tmrSplitStand1.Enabled = False
    Else
        
    
        For x = 1 To NumofPlayerCards
            imgPCard(x).Left = imgPCard(x).Left + TravelDis
        Next x
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrTest_Timer()
On Error Resume Next
'Label1.Caption = "numofcards: " & NumofPlayerCards & " b = " & b
lbltest.Caption = targetYP & " + " & CardSizeY & " = " & targetYP + CardSizeY

End Sub
