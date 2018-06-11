VERSION 5.00
Begin VB.Form frmComputerQuiz 
   Caption         =   "Computer Quiz"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraTFFive 
      Height          =   975
      Left            =   3840
      TabIndex        =   21
      Top             =   5640
      Width           =   1215
      Begin VB.OptionButton optFalseFive 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optTrueFive 
         Caption         =   "True"
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraTFFour 
      Height          =   975
      Left            =   3840
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
      Begin VB.OptionButton optFalseFour 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optTrueFour 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTFThree 
      Height          =   975
      Left            =   3840
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
      Begin VB.OptionButton optFalseThree 
         Caption         =   "False"
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
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optTrueThree 
         Caption         =   "True"
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
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame fraTFTwo 
      Height          =   975
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
      Begin VB.OptionButton optTrueTwo 
         Caption         =   "True"
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFalseTwo 
         Caption         =   "False"
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame fraTFOne 
      Height          =   975
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
      Begin VB.OptionButton optFalseOne 
         Caption         =   "False"
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optTrueOne 
         Caption         =   "True"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
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
      Left            =   480
      TabIndex        =   1
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label lblScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      TabIndex        =   24
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Image imgExFive 
      Height          =   825
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image ImgExFour 
      Height          =   825
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":099C
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgExThree 
      Height          =   825
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":1338
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgExTwo 
      Height          =   825
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":1CD4
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgExOne 
      Height          =   825
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":2670
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgCheckFive 
      Height          =   855
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":300C
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgCheckFour 
      Height          =   855
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":37A1
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgCheckThree 
      Height          =   855
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":3F36
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgCheckOne 
      Height          =   855
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":46CB
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgCheckTwo 
      Height          =   855
      Left            =   5280
      Picture         =   "frmComputerQuiz.frx":4E60
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblQuestionFive 
      Caption         =   "5. Bill Gates founded the company Apple."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label lblQuestionFour 
      Caption         =   "4. When a web page is sent over the internet, it is broken into packets so it can be sent more easily. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblQuestionThree 
      Caption         =   "3. Ada Lovlace is recognized as the first  computer programmer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblQuestionTwo 
      Caption         =   "2. The hexidecimal number system has six didgits."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblQuestionOne 
      Caption         =   "1. Transistors are often made of scillicon."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Computer Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmComputerQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim score As Integer

Private Sub cmdClear_Click()
    
    'All option uttons will be deselected
    optTrueOne.Value = False
    optTrueOne.Value = False
    optFalseOne.Value = False
    optTrueTwo.Value = False
    optFalseTwo.Value = False
    optTrueThree.Value = False
    optFalseThree.Value = False
    optTrueFour.Value = False
    optFalseFour.Value = False
    optTrueFive.Value = False
    optFalseFive.Value = False
    
    'The right and wrong pictures will dissapear
    imgCheckOne.Visible = False
    imgCheckTwo.Visible = False
    imgCheckThree.Visible = False
    imgCheckFour.Visible = False
    imgCheckFive.Visible = False
    imgExOne.Visible = False
    imgExTwo.Visible = False
    imgExThree.Visible = False
    ImgExFour.Visible = False
    imgExFive.Visible = False
    
    'The score message will be reset
    lblScore.Caption = ""
    
    'The score will be reset
    score = 0
    
End Sub

Private Sub cmdReturn_Click()
    Unload frmComputerQuiz
End Sub

Private Sub cmdSubmit_Click()
    
    'Which answers are right and wrong will be determined and the score will be tallied
    If optTrueOne.Value = True Then
        imgCheckOne.Visible = True
        score = score + 1
    ElseIf optFalseOne.Value = True Then
        imgExOne.Visible = True
        score = score
    End If
    
    If optFalseTwo.Value = True Then
        imgCheckTwo.Visible = True
        score = score + 1
    ElseIf optTrueTwo.Value = True Then
        imgExTwo.Visible = True
        score = score
    End If
    
    If optTrueThree.Value = True Then
        imgCheckThree.Visible = True
        score = score + 1
    ElseIf optFalseThree.Value = True Then
        imgExThree.Visible = True
        score = score
    End If
    
    If optTrueFour.Value = True Then
        imgCheckFour.Visible = True
        score = score + 1
    ElseIf optFalseFour.Value = True Then
        ImgExFour.Visible = True
        score = score
    End If
    
    If optFalseFive.Value = True Then
        imgCheckFive.Visible = True
        score = score + 1
    ElseIf optTrueFive.Value = True Then
        imgExFive.Visible = True
        score = score
    End If
    
    'The user's score or a completion message will appear
    If optTrueOne.Value = False And optFalseOne.Value = False Then
        lblScore.Caption = "Please complete the quiz."
        imgCheckOne.Visible = False
        imgCheckTwo.Visible = False
        imgCheckThree.Visible = False
        imgCheckFour.Visible = False
        imgCheckFive.Visible = False
        imgExOne.Visible = False
        imgExTwo.Visible = False
        imgExThree.Visible = False
        ImgExFour.Visible = False
        imgExFive.Visible = False
    ElseIf optTrueTwo.Value = False And optFalseTwo.Value = False Then
        lblScore.Caption = "Please complete the quiz."
        imgCheckOne.Visible = False
        imgCheckTwo.Visible = False
        imgCheckThree.Visible = False
        imgCheckFour.Visible = False
        imgCheckFive.Visible = False
        imgExOne.Visible = False
        imgExTwo.Visible = False
        imgExThree.Visible = False
        ImgExFour.Visible = False
        imgExFive.Visible = False
    ElseIf optTrueThree.Value = False And optFalseThree.Value = False Then
        lblScore.Caption = "Please complete the quiz."
        imgCheckOne.Visible = False
        imgCheckTwo.Visible = False
        imgCheckThree.Visible = False
        imgCheckFour.Visible = False
        imgCheckFive.Visible = False
        imgExOne.Visible = False
        imgExTwo.Visible = False
        imgExThree.Visible = False
        ImgExFour.Visible = False
        imgExFive.Visible = False
     ElseIf optTrueFour.Value = False And optFalseFour.Value = False Then
        lblScore.Caption = "Please complete the quiz."
        imgCheckOne.Visible = False
        imgCheckTwo.Visible = False
        imgCheckThree.Visible = False
        imgCheckFour.Visible = False
        imgCheckFive.Visible = False
        imgExOne.Visible = False
        imgExTwo.Visible = False
        imgExThree.Visible = False
        ImgExFour.Visible = False
        imgExFive.Visible = False
     ElseIf optTrueFive.Value = False And optFalseFive.Value = False Then
        lblScore.Caption = "Please complete the quiz."
        imgCheckOne.Visible = False
        imgCheckTwo.Visible = False
        imgCheckThree.Visible = False
        imgCheckFour.Visible = False
        imgCheckFive.Visible = False
        imgExOne.Visible = False
        imgExTwo.Visible = False
        imgExThree.Visible = False
        ImgExFour.Visible = False
        imgExFive.Visible = False
    Else
        lblScore.Caption = "Congratulations! Your score is " & score & " out of 5!"
    End If
    
End Sub

Private Sub Form_Load()
    score = 0
End Sub
