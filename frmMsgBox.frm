VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Message Box Generator V2!"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Styles"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4455
      Begin VB.Frame Frame3 
         Caption         =   "Default Button"
         Height          =   615
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton optDefaultButton4 
            Caption         =   "4"
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton optDefaultButton3 
            Caption         =   "3"
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton optDefaultButton2 
            Caption         =   "2"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton optDefaultButton1 
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   375
         End
      End
      Begin VB.CheckBox chkCenterForm 
         Caption         =   "Center Form"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtButton4 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Text            =   "Help"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtButton3 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Text            =   "Ignore"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtButton2 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "Retry"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtButton1 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "Abort"
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkHelpButton 
         Caption         =   "Help Button"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton opt1Button 
         Caption         =   "1 Button"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opt2Buttons 
         Caption         =   "2 Buttons"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opt3Buttons 
         Caption         =   "3 Buttons"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Icon"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4335
      Begin VB.OptionButton optIconNone 
         Caption         =   "None"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   4335
      End
      Begin VB.OptionButton optIconCritical 
         Caption         =   "Critical"
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optIconExclamation 
         Caption         =   "Exclamation"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optIconQuestion 
         Caption         =   "Question"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optIconInformation 
         Caption         =   "Information"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtMessage 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmMsgBox.frx":0000
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtMessageTitle 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Message"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblIcon 
      AutoSize        =   -1  'True
      Caption         =   "Icon:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblStyles 
      AutoSize        =   -1  'True
      Caption         =   "Styles:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "MsgBox Message:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label lblMsgBoxTitle 
      AutoSize        =   -1  'True
      Caption         =   "MsgBox Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdGenerate_Click()
  Dim Style As VbMsgBoxStyle
  
  If optIconNone.Value = True Then
    Style = Style '+ Nothing
  ElseIf optIconInformation.Value = True Then
    Style = Style + vbInformation
  ElseIf optIconQuestion.Value = True Then
    Style = Style + vbQuestion
  ElseIf optIconExclamation.Value = True Then
    Style = Style + vbExclamation
  ElseIf optIconCritical.Value = True Then
    Style = Style + vbCritical
  End If

  If opt1Button.Value = True Then
    Style = Style + vbOKOnly
  ElseIf opt2Buttons.Value = True Then
    Style = Style + vbOKCancel
  ElseIf opt3Buttons.Value = True Then
    Style = Style + vbAbortRetryIgnore
  End If
    
  If chkHelpButton.Value = 1 Then
    Style = Style + vbMsgBoxHelpButton
  End If
  
  If optDefaultButton1.Value = True Then
    Style = Style + vbDefaultButton1
  ElseIf optDefaultButton2.Value = True Then
    Style = Style + vbDefaultButton2
  ElseIf optDefaultButton3.Value = True Then
    Style = Style + vbDefaultButton3
  ElseIf optDefaultButton4.Value = True Then
    Style = Style + vbDefaultButton4
  End If

  MsgBoxEx txtMessage.Text, Style, txtMessageTitle.Text, , , chkCenterForm.Value, Me.hwnd, txtButton1.Text, txtButton2.Text, txtButton3.Text, txtButton4.Text
End Sub
