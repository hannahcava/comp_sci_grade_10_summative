VERSION 5.00
Begin VB.Form frmComputerQuiz 
   Caption         =   "Computer Quiz"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
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
      Left            =   3720
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
      Height          =   1215
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
Private Sub Label1_Click()

End Sub

Private Sub cmdClear_Click()
    
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
        
End Sub

Private Sub cmdReturn_Click()
    Unload frmComputerQuiz
End Sub

