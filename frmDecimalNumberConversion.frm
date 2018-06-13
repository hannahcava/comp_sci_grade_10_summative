VERSION 5.00
Begin VB.Form frmDecimalNumberConversion 
   Caption         =   "Decimal Number Convirsion"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHexadecimal 
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
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtBinary 
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
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtDecimal 
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
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lbl16 
      Caption         =   "16"
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
      Left            =   5160
      TabIndex        =   10
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblHexadecimal 
      Caption         =   "Hexadecimal Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lbl2 
      Caption         =   "2"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblBinary 
      Caption         =   "Binary Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lbl10 
      Caption         =   "10"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblDecimal 
      Caption         =   "Decimal Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Decimal Number Conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDecimalNumberConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intDecimal As Integer
Dim intBinary As Integer
Dim intHexadecimal As Integer
Dim i As Integer
Dim strBinaryOutput As String
Dim intRemainder As Integer
Dim intHalf As Integer


Private Sub cmdConvert_Click()
    
    intDecimal = Val(txtDecimal.Text)
    
    'Make sure the input is valid
    If intDecimal < 0 Or intDecimal > 225 Or txtDecimal.Text = "" Then
        lblError.Caption = "Invalid Input"
        Exit Sub
    End If
    
    'Converting decimal to binary
    strBinaryOutput = ""
    intBinary = intDecimal
    For i = 1 To 8
       intRemainder = intBinary Mod 2
       strBinaryOutput = strBinaryOutput + Str(intRemainder)
       intHalf = intBinary / 2
       intBinary = intHalf - intRemainder
    Next
    
    txtBinary.Text = StrReverse(strBinaryOutput)
    
End Sub


Private Sub Form_Load()
    
    strBinaryOutput = ""
    
End Sub

