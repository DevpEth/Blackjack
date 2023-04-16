VERSION 5.00
Begin VB.Form frmExtra 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Extra"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMoni 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox HideMe 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   150
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Start"
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblEnter 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Starting Balance:"
      BeginProperty Font 
         Name            =   "Microsoft Himalaya"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5640
      TabIndex        =   4
      Top             =   600
      Width           =   3045
   End
End
Attribute VB_Name = "frmExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim textval As String
Dim numval As String
    
Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()

    If txtMoni.Text = "" Then
        MsgBox "Please enter a valid amount."
    Else
        Moni = CLng(txtMoni.Text)
        gCheckStart = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
           
    '**********************************Moni Enter Screen****************************
    txtMoni.MaxLength = 5
    HideMe.Left = -10000
    
    cmdBack.Left = Me.Width / 100
    cmdBack.Top = Me.Height / 90
    
    cmdStart.Left = (Me.Width / 2) - (cmdStart.Width / 2)
    cmdStart.Top = (Me.Height / 1.1) - (cmdStart.Height / 2)
    
    txtMoni.Left = (Me.Width / 2) - (txtMoni.Width / 2)
    txtMoni.Top = (Me.Height / 2) - (txtMoni.Height / 2)
    
    lblEnter.Left = (Me.Width / 2) - (lblEnter.Width / 2)
    lblEnter.Top = (txtMoni.Top - lblEnter.Height) - 100
    txtMoni.Text = "100"
    txtMoni.SelStart = Len(txtMoni.Text)
    
    lblEnter.FontSize = 30
    lblEnter.Left = Me.Width / 2 - lblEnter.Width / 2
    
    '******************************************************************************
End Sub


Private Sub txtMoni_KeyPress(KeyAscii As Integer)

'check for numbers only
    If Not IsNumeric(txtMoni.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 43 Or KeyAscii = 46 Then
        KeyAscii = 0
    End If
'validate final input
    If KeyAscii = 13 Then
        Moni = CLng(txtMoni.Text)
        cmdStart.Value = True
    End If
End Sub
