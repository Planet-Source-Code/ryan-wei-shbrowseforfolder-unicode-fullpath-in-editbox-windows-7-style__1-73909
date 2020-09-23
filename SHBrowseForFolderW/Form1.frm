VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" (lpbi As BrowseInfo) As Long
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long


Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_EDITBOX = &H10
Private Const BIF_BROWSEINCLUDEURLS = &H80

Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
Private Const BIF_UAHINT = &H100
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Const BIF_NOTRANSLATETARGETS = &H400
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const BIF_SHAREABLE = &H8000
Private Const BIF_BROWSEFILEJUNCTIONS = &H10000


Private Function BrowseForFolder(TitleInfo As String) As String
  Dim lpIDList As Long
  Dim szTitleInfo() As Byte
'  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
'  m_CurrentDirectory = StartDir & vbNullChar
  szTitleInfo = TitleInfo & vbNullChar
'  szTitle = Title
  With tBrowseInfo
    .hWndOwner = hwnd
    .lpszTitle = VarPtr(szTitleInfo(0))
'.lpszTitle = lstrcat(StrPtr(szTitle), StrPtr(""))   'Invalid pointer, not recommended
'    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT 'old style
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_USENEWUI + BIF_NONEWFOLDERBUTTON
    .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH_UNICODE)
    SHGetPathFromIDList lpIDList, StrPtr(sBuffer)
    CoTaskMemFree lpIDList 'Clean it
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
  
End Function
Private Function GetAddressofFunction(Add As Long) As Long
  GetAddressofFunction = Add
End Function


Private Sub Command1_Click()

Me.Caption = BrowseForFolder("Select a folder or driver")

End Sub

