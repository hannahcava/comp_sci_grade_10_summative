VERSION 5.00
Begin VB.Form frmLogicGates 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Logic Gates"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Form"
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
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
   End
   Begin VB.ComboBox cmbGate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmLogicGates.frx":0000
      Left            =   4800
      List            =   "frmLogicGates.frx":0016
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame fraInputTwo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
      Begin VB.OptionButton optFalseTwo 
         BackColor       =   &H00FFFFC0&
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optTrueTwo 
         BackColor       =   &H00FFFFC0&
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraInputOne 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
      Begin VB.OptionButton optFalseOne 
         BackColor       =   &H00FFFFC0&
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optTrueOne 
         BackColor       =   &H00FFFFC0&
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lblTrueFalse 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image imgXOrGate 
      Height          =   1335
      Left            =   4800
      Picture         =   "frmLogicGates.frx":0057
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image imgXNorGate 
      Height          =   1335
      Left            =   4800
      Picture         =   "frmLogicGates.frx":2F79
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image imgOrGate 
      Height          =   1335
      Left            =   4800
      Picture         =   "frmLogicGates.frx":5E9B
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Image imgNorGate 
      Height          =   1335
      Left            =   4800
      Picture         =   "frmLogicGates.frx":8DBD
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Image imgNandGate 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   4800
      Picture         =   "frmLogicGates.frx":BCDF
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Line linSix 
      BorderWidth     =   5
      X1              =   4080
      X2              =   4920
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line linFive 
      BorderWidth     =   5
      X1              =   4080
      X2              =   4080
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line linFour 
      BorderWidth     =   5
      X1              =   2280
      X2              =   4080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line linThree 
      BorderWidth     =   5
      X1              =   4080
      X2              =   4920
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Image imgAndGate 
      Height          =   1335
      Left            =   4800
      Picture         =   "frmLogicGates.frx":EC01
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Line linTwo 
      BorderWidth     =   5
      X1              =   4080
      X2              =   4080
      Y1              =   3840
      Y2              =   3000
   End
   Begin VB.Line linOne 
      BorderWidth     =   5
      X1              =   2280
      X2              =   4080
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblLogicGatesTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Logic Gates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmLogicGates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbGate_Click()
    
    'The correct gate image will appear when selected
    If cmbGate.Text = "And Gate" Then
        imgAndGate.Visible = True
        imgNandGate.Visible = False
        imgOrGate.Visible = False
        imgNorGate.Visible = False
        imgXOrGate.Visible = False
        imgXNorGate.Visible = False
    ElseIf cmbGate.Text = "Nand Gate" Then
        imgAndGate.Visible = False
        imgNandGate.Visible = True
        imgOrGate.Visible = False
        imgNorGate.Visible = False
        imgXOrGate.Visible = False
        imgXNorGate.Visible = False
    ElseIf cmbGate.Text = "Or Gate" Then
        imgAndGate.Visible = False
        imgNandGate.Visible = False
        imgOrGate.Visible = True
        imgNorGate.Visible = False
        imgXOrGate.Visible = False
        imgXNorGate.Visible = False
    ElseIf cmbGate.Text = "Nor Gate" Then
        imgAndGate.Visible = False
        imgNandGate.Visible = False
        imgOrGate.Visible = False
        imgNorGate.Visible = True
        imgXOrGate.Visible = False
        imgXNorGate.Visible = False
    ElseIf cmbGate.Text = "XOr Gate" Then
        imgAndGate.Visible = False
        imgNandGate.Visible = False
        imgOrGate.Visible = False
        imgNorGate.Visible = False
        imgXOrGate.Visible = True
        imgXNorGate.Visible = False
    ElseIf cmbGate.Text = "XNor Gate" Then
        imgAndGate.Visible = False
        imgNandGate.Visible = False
        imgOrGate.Visible = False
        imgNorGate.Visible = False
        imgXOrGate.Visible = False
        imgXNorGate.Visible = True
    End If
    
End Sub

Private Sub cmdBack_Click()
    
    'The user can return to the main form
    Unload frmLogicGates
    
End Sub

Private Sub cmdCalculate_Click()
    
      'Depending on which gate is selected, the output is calculated by a given set of rules
      'This happens when the Calculate button is clicked
      If cmbGate.Text = "And Gate" Then
        If optTrueOne.Value = True And optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "True"
        Else
            lblTrueFalse.Caption = "False"
        End If
    ElseIf cmbGate.Text = "Nand Gate" Then
        If optTrueOne.Value = True And optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "False"
        Else
            lblTrueFalse.Caption = "True"
        End If
    ElseIf cmbGate.Text = "Or Gate" Then
        If optTrueOne.Value = True Or optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "True"
        Else
            lblTrueFalse.Caption = "False"
        End If
    ElseIf cmbGate.Text = "Nor Gate" Then
         If optTrueOne.Value = True Or optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "False"
        Else
            lblTrueFalse.Caption = "True"
        End If
    ElseIf cmbGate.Text = "XOr Gate" Then
        If optTrueOne.Value = True And optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "False"
        ElseIf optTrueOne.Value = True And optFalseTwo.Value = True Then
            lblTrueFalse.Caption = "True"
        ElseIf optFalseOne.Value = True And optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "True"
        ElseIf optFalseOne.Value = True And optFalseTwo.Value = True Then
            lblTrueFalse.Caption = "False"
        End If
    ElseIf cmbGate.Text = "XNor Gate" Then
        If optTrueOne.Value = True And optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "True"
        ElseIf optTrueOne.Value = True And optFalseTwo.Value = True Then
            lblTrueFalse.Caption = "False"
        ElseIf optFalseOne.Value = True And optTrueTwo.Value = True Then
            lblTrueFalse.Caption = "False"
        ElseIf optFalseOne.Value = True And optFalseTwo.Value = True Then
            lblTrueFalse.Caption = "True"
        Else
            lblTrueFalse.Caption = "False"
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    'All the gates will be invisable when the form loads
    imgAndGate.Visible = False
    imgNandGate.Visible = False
    imgOrGate.Visible = False
    imgNorGate.Visible = False
    imgXOrGate.Visible = False
    imgXNorGate.Visible = False
    
    'All options are deselected when the form loads
    optTrueOne.Value = False
    optFalseOne.Value = False
    optTrueTwo.Value = False
    optFalseTwo.Value = False
    
End Sub



