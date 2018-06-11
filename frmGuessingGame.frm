VERSION 5.00
Begin VB.Form frmGuessingGame 
   Caption         =   "Guessing Game"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtNumberGuess 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.HScrollBar hsbNumberGuess 
      Height          =   375
      Left            =   240
      Max             =   100
      TabIndex        =   7
      Top             =   3000
      Width           =   4695
   End
   Begin VB.CommandButton cmdStartGuessingGame 
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblWarning 
      Height          =   735
      Left            =   3840
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblHighLow 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblMax 
      Caption         =   "Max:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblMin 
      Caption         =   "Min:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblSetRange 
      Caption         =   "Set Range:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblGuessingGameTitle 
      Alignment       =   2  'Center
      Caption         =   "Guessing Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmGuessingGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim secretNumber As Integer
Dim intGuess As Integer
Dim intMin As Integer
Dim intMax As Integer

Private Sub cmdReturn_Click()
    'The user can return to the menu
    Unload frmGuessingGame
End Sub

Private Sub cmdStartGuessingGame_Click()
    
    intMin = Val(txtMin.Text)
    intMax = Val(txtMax.Text)
    
    'The game is enabled when the user presses start
    txtNumberGuess.Enabled = True
    hsbNumberGuess.Enabled = True
    
    'The secret number is determined
     secretNumber = Int((intMax - intMin + 1) * Rnd + intMin)

    'If the user chooses to play again, the text box and scroll bar will be cleared
    txtNumberGuess.Text = ""
    hsbNumberGuess.Value = intMin
    lblHighLow.Caption = ""
    
    'The min can't be higher than the max
    If intMin > intMax Then
       txtNumberGuess.Enabled = False
       hsbNumberGuess.Enabled = False
       lblWarning.Caption = "Your maximum value must be higher than your minimum value."
    End If

    'The max and min will be set to the user's chosen values
    hsbNumberGuess.Min = intMin
    hsbNumberGuess.Max = intMax

End Sub

Private Sub Form_Load()
    
    'The game is disabled until the user presses start
    txtNumberGuess.Enabled = False
    hsbNumberGuess.Enabled = False
    
    'The secret number is defined until further notice
    secretNumber = 0
    
    Randomize Timer
    
End Sub

Private Sub hsbNumberGuess_Change()
    
    intGuess = Val(txtNumberGuess.Text)
    
    'The text box will change with the scroll bar input
    If hsbNumberGuess.Value = 0 Then
        txtNumberGuess.Text = ""
    ElseIf hsbNumberGuess.Value > 0 Then
        txtNumberGuess.Text = hsbNumberGuess.Value
    End If
    
    'The user's results will be displayed for them
    If intGuess > secretNumber Then
        lblHighLow.Caption = "Too High"
    ElseIf intGuess < secretNumber Then
        lblHighLow.Caption = "Too Low"
    ElseIf intGuess = secretNumber Then
        lblHighLow.Caption = "Correct!"
    ElseIf txtNumberGuess.Text = "" Then
        lblHighLow.Caption = ""
    End If
    
End Sub

Private Sub txtNumberGuess_Change()
    
    intGuess = Val(txtNumberGuess.Text)
    
    'The scroll bar will change with the text box input
    If txtNumberGuess.Text = "" Then
        hsbNumberGuess.Value = 0
    ElseIf txtNumberGuess.Text <> "" Then
        hsbNumberGuess.Value = txtNumberGuess.Text
    End If
        
    'The user's results will be displayed for them
    If intGuess > secretNumber Then
        lblHighLow.Caption = "Too High"
    ElseIf intGuess < secretNumber Then
        lblHighLow.Caption = "Too Low"
    ElseIf intGuess = secretNumber Then
        lblHighLow.Caption = "Correct!"
    ElseIf txtNumberGuess.Text = "" Then
        lblHighLow.Caption = ""
    End If
    
    'The game will end when the user guesses the right number
    If intGuess = secretNumber Then
        txtNumberGuess.Enabled = False
        hsbNumberGuess.Enabled = False
        cmdStartGuessingGame.Caption = "Play Again"
    End If
    
End Sub
