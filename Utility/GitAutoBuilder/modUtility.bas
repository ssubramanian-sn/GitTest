Attribute VB_Name = "modUtility"
' modUtility - Maritime Workstation GUI Utility Functions
' 7/2002 ms
'

Option Explicit

Private Const SW_SHOWNORMAL = 1

'
'   HIDE              0 Hides the window and passes activation to another window.
'   SHOWNORMAL        1 Activates and displays a window.
'                       If the window is minimized or maximized, Windows restores it to its original
'                       size and position (same as RESTORE).
'   SHOWMINIMIZED     2 Activates a window and displays it as an icon.
'   SHOWMAXIMIZED     3 Activates a window and displays it as a maximized window.
'   SHOWMINNOACTIVATE 4 Displays a window in its most recent size and position.
'                       The window that is currently active remains active.
'   SHOW              5 Activates a window and displays it in its current size and position.
'   MINIMIZE          6 Minimizes the specified window and activates the top-level window in the system's list.
'   SHOWMINNOACTIVE   7 Displays a window as an icon. The window that is currently active remains active.
'   SHOWNA            8 Displays a window in its current state. The window that is currently active remains active.
'   RESTORE           9 Activates and displays a window. If the window is minimized or maximized,
'                       Windows restores it to its original size and position (same as SHOWNORMAL).



Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess _
As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const SYNCHRONIZE As Long = &H100000
Const INFINITE As Long = &HFFFFFFFF

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal _
  lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long

Declare Function WinQueryPerformanceCounter Lib "kernel32" _
    Alias "QueryPerformanceCounter" (lpPerformanceCount As _
    Currency) As Long
    
Declare Function WinQueryPerformanceFrequency Lib "kernel32" _
    Alias "QueryPerformanceFrequency" (lpFrequency As Currency) _
    As Long
Public TimerFrequency    As Currency

Public Const RegHKey = HKEY_LOCAL_MACHINE
Public Const RegSection = "Software\Maritime Systems\GitBuilder"

Public Function RunBat(sFile As String) As Long
   Dim lresp As Long
   Dim dummy As Long
   Dim Parms As String
   
   On Error GoTo FunctionError
   dummy = 1

   lresp = ShellExecute(dummy, "open", sFile, vbNullString, vbNullString, SW_SHOWNORMAL)
   
   RunBat = lresp
   Exit Function
FunctionError:
'   moParent.RaiseError "Error in mwSession.mwAPI.RunShellExecute", Err.Number, Err.Description
   RunBat = lresp
End Function

Public Function GetRegVal(ByVal ValueName As String, DefaultValue As Variant, _
    Optional KeyType As Integer = REG_SZ) As Variant

   If CheckRegistryKey(RegHKey, RegSection) = False Then
      CreateRegistryKey RegHKey, RegSection
   End If

   GetRegVal = GetRegistryValue(RegHKey, RegSection, ValueName, KeyType, DefaultValue)

End Function
Public Sub SetRegVal(ByVal ValueName As String, value As Variant, Optional KeyType As Integer = REG_SZ)
   
   If CheckRegistryKey(RegHKey, RegSection) = False Then
      CreateRegistryKey RegHKey, RegSection
   End If
   
   SetRegistryValue RegHKey, RegSection, ValueName, KeyType, value
   
End Sub



Function TimeGetTime() As Single

    Static Overhead     As Currency
    Dim CurrentTime     As Currency
    Dim Time2           As Currency

    If 0 = TimerFrequency Then
        Call WinQueryPerformanceFrequency(TimerFrequency)
        'TimerFrequency = (CurrentTime * 10000) / 1000
        'TimeGetTime = 0
        Call WinQueryPerformanceCounter(CurrentTime)
        Call WinQueryPerformanceCounter(Time2)
        Overhead = Time2 - CurrentTime
   End If
   Call WinQueryPerformanceCounter(CurrentTime)
   TimeGetTime = CurrentTime

End Function

Public Function StartTimer() As Single
   
   StartTimer = TimeGetTime

End Function

Public Function EndTimer(StartTime As Single, Optional msg As String = "") As Single
   Dim Elapsed As Currency
   Dim CurrTime As Currency
   
   If StartTime > -1 Then
      CurrTime = TimeGetTime
      EndTimer = (CurrTime - StartTime) / TimerFrequency
   Else
      EndTimer = 0
   End If
      
   If Len(msg) > 0 Then
      Debug.Print msg & " " & Format(EndTimer, "####.00")
   End If
   
End Function


Public Function KillObject(ByRef obj As Object)
   If Not obj Is Nothing Then Set obj = Nothing
End Function



Public Function GetIniParameter(sIniFile As String, sIniHeading As String, sIniParameter As String) As String
'/*** Get a parameter setting from an INI file
Dim nIniFile As Integer
Dim IsHeaderFound As Boolean
Dim IsParameterFound As Boolean
Dim sInLine As String
    
On Error GoTo Error_Trap

   ' Set the variable defaults
   If sIniHeading = "" Then
      IsHeaderFound = True
   Else
      IsHeaderFound = False
   End If
   
   IsParameterFound = False
   
   nIniFile = FreeFile
    
   ' Open the INI file
   Open sIniFile For Input As #nIniFile
    
    ' Read all lines of the INI file, until the search criteria has been found
   Do While (Not EOF(nIniFile)) And (IsParameterFound = False)
      Line Input #nIniFile, sInLine
      Let sInLine = Trim(sInLine)
      
      ' If the line is not empty and not a remark line
      If (sInLine <> "") And (Left(sInLine, 1) <> ";") Then
        
         ' If the first char is a '[' then it is s new header section
         If Left(sInLine, 1) = "[" Then IsHeaderFound = False
        
         ' Check if the INI section heading has been found
         If Not (IsHeaderFound) Then
            If sIniHeading = "" Then
               IsHeaderFound = True
            Else
               If Left(UCase(sInLine), Len(sIniHeading) + 2) = _
                  "[" & UCase(sIniHeading) & "]" Then
                     IsHeaderFound = True
               End If
            End If
         Else
            If Left(UCase(sInLine), Len(sIniParameter)) = _
               UCase(sIniParameter) Then
                  IsParameterFound = True
          
                  GetIniParameter = Trim(Right(sInLine, Len(sInLine) - _
                    (Len(sIniParameter) + 1)))
            End If
         End If
      End If
   Loop
    
   Close #nIniFile
    
   If Not IsParameterFound Then GetIniParameter = ""
    
   Exit Function

Error_Trap:
    MsgBox "Error retrieving INI parameters!" & vbCrLf & "Error #" & _
            Str(Err.Number) & "(" & Err.Description & ")" & vbCrLf & _
            "Raised by: " & Err.Source, vbCritical + vbOKOnly, "Error"
    Exit Function
End Function


Public Function PutIniParameter(strIniFile As String, strIniParameter As String, strParamValue As String) As Boolean
'/*** Update a parameter setting in an INI file
Dim intIniFile As Integer
Dim blnHeaderFound As Boolean
Dim blnParameterFound As Boolean
Dim strInLine As String
Dim strIniLine(5000) As String
Dim intIniPtr As Integer
Dim intIniLine As Integer
Dim i As Integer
    
    On Error GoTo Error_Trap

    ' Set the variable defaults
    blnHeaderFound = False
    blnParameterFound = False
    intIniFile = FreeFile
    intIniPtr = 0
    
    ' Read the INI file into an array
    Open strIniFile For Input As #intIniFile
    
    ' Read all lines of the INI file, until the search criteria has been found
    Do While (Not EOF(intIniFile))
      Let intIniPtr = intIniPtr + 1
      Line Input #intIniFile, strInLine
      Let strInLine = Trim(strInLine)
      Let strIniLine(intIniPtr) = strInLine
      
      ' If the line is not empty and not a remark line
      If (strInLine <> "") And (Left(strInLine, 1) <> ";") Then
        
          If Left(UCase(strInLine), Len(strIniParameter)) = _
             UCase(strIniParameter) Then
                blnParameterFound = True
            intIniLine = intIniPtr
          End If
      End If
    Loop
    
    Close #intIniFile
    
    If Not blnParameterFound Then
        If strIniParameter = "VersionComments" Then
            intIniLine = AddIniParameter(strIniFile, "RevisionVer", strIniParameter, strParamValue)
        End If
        If strIniParameter = "VersionFileDescription" Then
            intIniLine = AddIniParameter(strIniFile, "VersionComments", strIniParameter, strParamValue)
        End If
      If intIniLine = 0 Then
        PutIniParameter = False
        Exit Function
      End If
    End If
    
    Open strIniFile For Output As #intIniFile
    
    For i = 1 To intIniPtr
      If i = intIniLine Then
        Print #intIniFile, strIniParameter; "="; strParamValue
      Else
        Print #intIniFile, strIniLine(i)
      End If
    Next i
    
    Close #intIniFile
    
    PutIniParameter = True
    
    Exit Function


Error_Trap:
    MsgBox "Error reading/writing INI parameters!" & vbCrLf & "Error #" & _
            Str(Err.Number) & "(" & Err.Description & ")" & vbCrLf & _
            "Raised by: " & Err.Source, vbCritical + vbOKOnly, "Error"
    Exit Function
End Function



Public Function RunBatInHideMode(sFile As String) As Long
   Dim lresp As Long
   Dim dummy As Long
   Dim Parms As String
   
   On Error GoTo FunctionError
   dummy = 1

   lresp = ShellExecute(dummy, "open", sFile, vbNullString, vbNullString, SW_SHOWNORMAL)
   
   RunBatInHideMode = lresp
   Exit Function
FunctionError:
'   moParent.RaiseError "Error in mwSession.mwAPI.RunShellExecute", Err.Number, Err.Description
   RunBatInHideMode = lresp
End Function



Public Sub ShellAndWait(ByVal program_name As String, ByVal window_style As VbAppWinStyle)
Dim process_id As Long
Dim process_handle As Long
Dim lngdb  As Long

    ' Start the program.
    process_id = Shell(program_name, window_style)
    'process_id = ShellExecute(Me.hWnd, "open", program_name, vbNullString, App.Path, 1)


    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
         DoEvents
    End If
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
End Sub


Public Function AddIniParameter(strIniFile As String, strNextToIniParameter As String, strIniParameter As String, strParamValue As String) As Integer
'/*** Update a parameter setting in an INI file
Dim intIniFile As Integer
Dim blnHeaderFound As Boolean
Dim blnParameterFound As Boolean
Dim strInLine As String
Dim strIniLine(5000) As String
Dim intIniPtr As Integer
Dim intIniLine As Integer
Dim i As Integer
    
    On Error GoTo Error_Trap

    ' Set the variable defaults
    blnHeaderFound = False
    blnParameterFound = False
    intIniFile = FreeFile
    intIniPtr = 0
    
    ' Read the INI file into an array
    Open strIniFile For Input As #intIniFile
    
    ' Read all lines of the INI file, until the search criteria has been found
    Do While (Not EOF(intIniFile))
      Let intIniPtr = intIniPtr + 1
      Line Input #intIniFile, strInLine
      Let strInLine = Trim(strInLine)
      Let strIniLine(intIniPtr) = strInLine
      
      ' If the line is not empty and not a remark line
      If (strInLine <> "") And (Left(strInLine, 1) <> ";") Then
        
          If Left(UCase(strInLine), Len(strNextToIniParameter)) = _
             UCase(strNextToIniParameter) Then
                blnParameterFound = True
            intIniLine = intIniPtr
          End If
      End If
    Loop
    
    Close #intIniFile
    
    If Not blnParameterFound Then
      AddIniParameter = 0
      Exit Function
    End If
    
    Open strIniFile For Output As #intIniFile
    
    For i = 1 To intIniPtr
      If i = intIniLine Then
        Print #intIniFile, strIniLine(i)
        Print #intIniFile, strIniParameter; "="; strParamValue
        AddIniParameter = intIniLine + 1
      Else
        Print #intIniFile, strIniLine(i)
      End If
    Next i
    
    Close #intIniFile
    
    Exit Function


Error_Trap:
    MsgBox "Error reading/writing INI parameters!" & vbCrLf & "Error #" & _
            Str(Err.Number) & "(" & Err.Description & ")" & vbCrLf & _
            "Raised by: " & Err.Source, vbCritical + vbOKOnly, "Error"
    Exit Function
End Function

'Returns entry from .ini file:''(strIniFile)''[Section]'Key=Entry'
Function GetIni(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String) As String
On Error GoTo FunctionError:
   Dim lngRC As Long
   Dim strEntry As String
   strEntry = Space$(256)
   lngRC = GetPrivateProfileString(strSection, strKey, "", strEntry, Len(strEntry), strIniFile)
   strEntry = Left$(strEntry, lngRC)
   GetIni = strEntry
   Exit Function
FunctionError:
   MsgBox Err.Description
End Function

'Puts an entry to an .ini file:''(strIniFile)''[Section]'Key=Entry'
Sub PutIni(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, ByVal strEntry As String)
On Error GoTo SubError:
   Dim lngRC As Long
   lngRC = WritePrivateProfileString(strSection, strKey, strEntry, strIniFile)
   Exit Sub
SubError:
   MsgBox Err.Description
End Sub

