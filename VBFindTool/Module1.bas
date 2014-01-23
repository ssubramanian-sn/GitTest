Attribute VB_Name = "Module1"

Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 3
Public Const DRIVE_CDROM As Long = 5
Private Const BIF_VALIDATE As Long = &H20


Const WM_USER = &H400
Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const MAX_PATH = 260

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Declare Function SHBrowseForFolder _
Lib "shell32" (lpbi As BrowseInfo) As Long

Public Declare Function GetDriveType Lib "kernel32" _
   Alias "GetDriveTypeA" _
  (ByVal nDrive As String) As Long

Private Declare Function SHGetPathFromIDList _
Lib "shell32" (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    
Public objFso As New FileSystemObject



Public Function OpenDirectoryTV(objForm As VB.Form, Optional odtvTitle As String) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = odtvTitle
    With tBrowseInfo
           .hwndOwner = objForm.hWnd
           .lpszTitle = lstrcat(szTitle, "")
           .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_VALIDATE
          .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
        End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        OpenDirectoryTV = sBuffer
    End If
End Function




Private Function GetAddressofFunction(Functionaddress As Long) As Long
   On Error GoTo FunctionError
   
   GetAddressofFunction = Functionaddress
   
   Exit Function
FunctionError:
   MsgBox "Genral error in mwEventLink.modEventLinkProcessing.GetAddressofFunction "
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   On Error GoTo FunctionError
   Dim spath As String
   Dim bFlag  As Long

   Select Case uMsg
      Case BFFM_INITIALIZED
         Call SendMessage(hWnd, BFFM_ENABLEOK, 0, ByVal 0&)

'      Call SetWindowText(hWnd, "Browse for CD/DVD Drive")

      Case BFFM_SELCHANGED
         spath = Space$(MAX_PATH)
        If SHGetPathFromIDList(lParam, spath) Then

           'if drive type is CD set True othewise False
'            bFlag = (GetDriveType(spath) = DRIVE_CDROM)
            If False = objFso.FolderExists(spath) Then
               Call SendMessage(hWnd, BFFM_ENABLEOK, 0, False)
            End If
         
         End If

   End Select
'   If uMsg = BFFM_INITIALIZED Then
'      SendMessage hWnd, BFFM_SETSELECTION, 1, "D:\MWS_DEV"
'   End If

   Exit Function
FunctionError:
   MsgBox "Genral error in mwEventLink.modEventLinkProcessing.BrowseCallbackProc "
End Function

