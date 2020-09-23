Attribute VB_Name = "Module1"
'by ryanwei2005@gmail.com
Option Explicit

Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Const MAX_PATH = 260&
Public Const MAX_PATH_UNICODE = 2 * MAX_PATH - 1

Private Const BFFM_INITIALIZED = 1&
Private Const BFFM_SELCHANGED = 2&
Private Const WM_USER = &H400
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const WM_SETTEXT = &HC
Private Const EM_SETREADONLY = &HCF
Private Const EM_NOSETFOCUS = (&H1500 + 7)
Private Const WM_KILLFOCUS = &H8


Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim lpIDList As Long
  Dim lRet As Long
  Dim sBuffer As String
  Dim Fhwnd As Long
  Dim sysDir As String
  Dim szPath() As Byte
  
  sysDir = Environ("SystemDrive") & "\"
  
  On Error GoTo errhandler
  Select Case uMsg
    Case BFFM_INITIALIZED
      Call SendMessage(hwnd, BFFM_SETSELECTION, True, ByVal sysDir)
      Fhwnd = FindWindowEx(hwnd, 0, "Edit", vbNullString)
      Call SendMessage(Fhwnd, WM_SETTEXT, 0, ByVal sysDir)
'      EnableWindow Fhwnd, 0& ' below is an internal message from microsoft only supported by windows 7 so alternatively you can use EnableWindow
      Call SendMessage(Fhwnd, EM_SETREADONLY, True, ByVal 0&)
      Call SendMessage(Fhwnd, EM_NOSETFOCUS, 0&, ByVal 0&) ' internal message, not supported  by xp,windows 2000 etc. or maybe future OS
      
      
    Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH_UNICODE)
      
      lRet = SHGetPathFromIDList(lParam, StrPtr(sBuffer))
      If lRet = 1 Then
'        Call SendMessageT(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
        Fhwnd = FindWindowEx(hwnd, 0, "Edit", vbNullString)
        szPath = sBuffer
        Call SendMessageLong(Fhwnd, WM_SETTEXT, 0, VarPtr(szPath(0)))
        Call SendMessage(Fhwnd, WM_KILLFOCUS, 0&, ByVal 0&)
    
      End If
      
  End Select

errhandler:
  BrowseCallbackProc = 0
End Function


