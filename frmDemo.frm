VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MsgBox V2 Demonstration"
   ClientHeight    =   1335
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
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "MsgBox Generator"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExitDemo 
      Cancel          =   -1  'True
      Caption         =   "E&xit Demo"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   2160
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   1680
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse Dialog"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select a Directory"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReallyQuit As Boolean

Private Const BTN_SELECT = vbOK
Private Const BTN_OK = vbCancel

Private Const BTN_SEARCH = vbAbort
Private Const BTN_WAIT = vbRetry
Private Const BTN_CANCEL = vbIgnore

Private Const BTN_QUIT = vbOK

Private Sub cmdExitDemo_Click()
  If MsgBoxEx("Can't quit this way either!  Well I guess you can, choose from the choices below:", vbQuestion + vbOKCancel, "Quit?", , , False, , "&Quit", "&Cancel") = BTN_QUIT Then End
End Sub

Private Sub Command1_Click()
  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer

  Label1.Caption = ""

  'Fill the BROWSEINFO structure with the
  'needed data. To accomodate comments, the
  'With/End With sytax has not been used, though
  'it should be your 'final' version.

  'hwnd of the window that receives messages
  'from the call. Can be your application
  'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

  'Pointer to the item identifier list specifying
  'the location of the "root" folder to browse from.
  'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

  'message to be displayed in the Browse dialog
  bi.lpszTitle = "Select your Windows\System\ directory"

  'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS

  'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)

  'the dialog has closed, so parse & display the
  'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)

  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
    pos = InStr(path, Chr$(0))
    Label1.Caption = Left(path, pos - 1)
  End If

  Call CoTaskMemFree(pidl)
End Sub

Private Sub Command2_Click()
  Timer1.Enabled = True

  Select Case MsgBoxEx("Select a directory:" + vbCrLf + vbCrLf + vbCrLf, vbOKCancel + vbDefaultButton1, "TEST", , , True, Me.hwnd, "Select...", "OK")
    Case BTN_SELECT
      Dim bi As BROWSEINFO
      Dim pidl As Long
      Dim path As String
      Dim pos As Integer

      Label1.Caption = ""

      'Fill the BROWSEINFO structure with the
      'needed data. To accomodate comments, the
      'With/End With sytax has not been used, though
      'it should be your 'final' version.

      'hwnd of the window that receives messages
      'from the call. Can be your application
      'or the handle from GetDesktopWindow().
      bi.hOwner = Me.hwnd

      'Pointer to the item identifier list specifying
      'the location of the "root" folder to browse from.
      'If NULL, the desktop folder is used.
      bi.pidlRoot = 0&

      'message to be displayed in the Browse dialog
      bi.lpszTitle = "Select a directory"

      'the type of folder to return.
      bi.ulFlags = BIF_RETURNONLYFSDIRS

      'show the browse for folders dialog
      pidl = SHBrowseForFolder(bi)

      'the dialog has closed, so parse & display the
      'user's returned folder selection contained in pidl
      path = Space$(MAX_PATH)

      If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        pos = InStr(path, Chr$(0))
        Label1.Caption = Left(path, pos - 1)
      End If

      Call CoTaskMemFree(pidl)

      While Label1.Caption <> ""
        Timer2.Enabled = True

        Select Case MsgBoxEx("Select a directory:" + vbCrLf + vbCrLf + Label1.Caption, &HF0& + vbOKCancel, "TEST", , , True, Me.hwnd, "Select...", "OK")
          Case BTN_OK
            If Label1.Caption <> "" Then
              Select Case MsgBoxEx(Label1.Caption + " has been selected.  What would you like to do?", vbQuestion + vbAbortRetryIgnore + vbMsgBoxHelpButton, "Message", , , True, Me.hwnd, "&Search", "&Wait", "&Cancel", "HELP??")
                Case BTN_SEARCH
                  ' ... Search was pressed, put your code here.
                  'MsgBox "Search Pressed"
                Case BTN_WAIT
                  ' ... Wait was pressed, put your code here.
                  'MsgBox "Wait Pressed"
                Case BTN_CANCEL
                  ' ... Cancel was pressed, put your code here.
                  'MsgBox "Cancel Pressed"
              End Select
            End If
        End Select

        Exit Sub
      Wend
  End Select
End Sub

Private Sub Command3_Click()
  frmMsgBox.Show
  ReallyQuit = True
  Unload Me
End Sub

Private Sub Form_Load()
  With Font
    .Name = "Arial"
    .Bold = True
  End With

  Dim IFont As IFont
  Set IFont = Font

  g_hBoldFont = IFont.hFont
  Set IFont = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If ReallyQuit = False Then
    Cancel = 1
    MsgBoxEx "Can't quit that way!  ;)", vbCritical + vbOKCancel, "Message", , , True, Me.hwnd, "&Grr!", "&Close"
    cmdExitDemo.Caption = "E&xit this way!"
  End If
End Sub

Private Sub Timer1_Timer()
  hMsgBox = FindWindow("#32770", MsgBox_Title)

  If hMsgBox Then
    Dim hStatic&, hButton&, stMsgBoxRect2 As RECT
    Dim stStaticRect As RECT, stButtonRect As RECT

    hStatic = FindWindowEx(hMsgBox, API_FALSE, "Static", MSGBOXTEXT)
    hButton = FindWindowEx(hMsgBox, API_FALSE, "Button", "OK")

    If hButton Then
      EnableWindow hButton, 0
      'Call SendMessage(hButton, WM_SETFONT, g_hBoldFont, ByVal API_TRUE)
    End If

    Timer1.Enabled = False
  End If
End Sub

Private Sub Timer2_Timer()
  hMsgBox = FindWindow("#32770", MsgBox_Title)

  If hMsgBox Then
    Dim hStatic&, hButton&, stMsgBoxRect2 As RECT
    Dim stStaticRect As RECT, stButtonRect As RECT

    hStatic = FindWindowEx(hMsgBox, API_FALSE, "Static", MSGBOXTEXT)
    hButton = FindWindowEx(hMsgBox, API_FALSE, "Button", "Select...")

    If hButton Then
      EnableWindow hButton, 0
      'Call SendMessage(hButton, WM_SETFONT, g_hBoldFont, ByVal API_TRUE)
    End If

    Timer2.Enabled = False
  End If
End Sub
