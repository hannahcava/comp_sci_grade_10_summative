VERSION 5.00
Begin VB.Form frmDecimalNumberConversion 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Decimal Number Convirsion"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6015
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
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblDecimal 
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
Dim strHexaOutput As String


Private Sub cmdConvert_Click()
    
    intDecimal = Val(txtDecimal.Text)
    
    'Make sure the input is valid
    If intDecimal < 0 Or intDecimal > 225 Or txtDecimal.Text = "" Then
        lblError.Caption = "Invalid Input"
        Exit Sub
    End If
    
    'Converting decimal to binary
    'Initializing variables
    strBinaryOutput = ""
    intBinary = intDecimal
    
    'Calculating each new digit in the binary number
    'It is important to have a backslash, because that is for dividing integers, instead of a forwardslash
    'I will only need a maximum of 8 digits, so 1 to 8 is sufficiant
    For i = 1 To 8
       intRemainder = intBinary Mod 2
       strBinaryOutput = Str(intRemainder) + strBinaryOutput
       intHalf = intBinary \ 2
       intBinary = intHalf
    Next
    
    'The binary number will be displayed in the text box
    txtBinary.Text = strBinaryOutput
    
    'Converting decimal to hexadecimal
    'Initialising variables
    strHexaOutput = ""
    intHexadecimal = intDecimal
    intRemainder = 0
    intHalf = 0
    
    'Calculating each digit
    'The same rules of the backslash apply
    For i = 1 To 2
        intRemainder = intHexadecimal Mod 16
    
        'This will convert number greater than 9 to their hexadecimal counterparts
        If intRemainder < 10 Then
            strHexaOutput = Str(intRemainder)
        ElseIf intRemainder = 10 Then
            strHexaOutput = "A"
        ElseIf intRemainder = 11 Then
            strHexaOutput = "B"
        ElseIf intRemainder = 12 Then
            strHexaOutput = "C"
        ElseIf intRemainder = 13 Then
            strHexaOutput = "D"
        ElseIf intRemainder = 14 Then
            strHexaOutput = "E"
        ElseIf intRemainder = 15 Then
            strHexaOutput = "F"
        End If
        
        intHalf = intHexadecimal \ 16
        intHexadecimal = intHalf
        
        'The hexadecimal number will be displayed in the text box
        txtHexadecimal.Text = strHexaOutput + txtHexadecimal.Text
    Next
    
End Sub


Private Sub cmdReturn_Click()
    
    'The user can return to the main form
    Unload frmDecimalNuberConvirsion
    
End Sub

Private Sub Form_Load()
    
    'Initializing variables
    strBinaryOutput = ""
    intDecimal = 0
    intBinary = 0
    intHexadecimal = 0
    i = 0
    intRemainder = 0
    intHalf = 0
    strHexaOutput = ""
    
End Sub

