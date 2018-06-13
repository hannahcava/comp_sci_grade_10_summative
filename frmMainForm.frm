VERSION 5.00
Begin VB.Form frmMainForm 
   Caption         =   "Main Form"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBinaryHexadecimal 
      Caption         =   "Decimal Number Conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdComputerDefinitions 
      Caption         =   "Computer Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuessingGame 
      Caption         =   "Guessing Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblMainSubtitle 
      Alignment       =   2  'Center
      Caption         =   "ICS2O Summative June 2018"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblMainTitle 
      Alignment       =   2  'Center
      Caption         =   "Hannah Cavanagh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBinaryHexadecimal_Click()
    frmDecimalNumberConversion.Show vbModal
End Sub

Private Sub cmdComputerDefinitions_Click()
    'The computer definition game will load
    frmComputerQuiz.Show vbModal
End Sub

Private Sub cmdExit_Click()
    'The user can exit the program
    End
End Sub

Private Sub cmdGuessingGame_Click()
    'The guessing game is loaded
    frmGuessingGame.Show vbModal
End Sub

