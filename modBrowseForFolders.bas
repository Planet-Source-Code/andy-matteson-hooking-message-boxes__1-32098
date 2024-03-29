Attribute VB_Name = "modBrowseForFolders"
'--------------------------------------------------------------
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------

Public Type BROWSEINFO
  hOwner           As Long
  pidlRoot         As Long
  pszDisplayName   As String
  lpszTitle        As String
  ulFlags          As Long
  lpfn             As Long
  lParam           As Long
  iImage           As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const MAX_PATH = 260

Public Declare Function SHGetPathFromIDList Lib "shell32" _
    Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32" _
    Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" _
    (ByVal pv As Long)

