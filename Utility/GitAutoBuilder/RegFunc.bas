Attribute VB_Name = "RegistryAPI"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_ALL = &H1F0000

Const KEY_QUERY_VALUE = &H1
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Const ERROR_SUCCESS = 0&

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2


' Return True if the system has a math processor.

Function MathProcessor() As Boolean
    Dim hKey As Long, key As String
    key = "HARDWARE\DESCRIPTION\System\FloatingPointProcessor"
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, key, 0, KEY_READ, hKey) = 0 Then
        ' If the open operation succeeded, the key exists
        MathProcessor = True
        ' Important: close the key before exiting.
        RegCloseKey hKey
    End If
End Function

' Test if a Registry key exists.

Function CheckRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
    Dim handle As Long
    ' Try to open the key.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0 Then
        ' The key exists.
        CheckRegistryKey = True
        ' Close it before exiting.
        RegCloseKey handle
    End If
End Function

' Create a registry key, then close it.
' Returns True if the key already existed, False if it was created.

Function CreateRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
    Dim handle As Long, disposition As Long
    If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, handle, disposition) Then
        Err.Raise 1001, , "Unable to create the registry key"
    Else
        ' Return True if the key already existed.
        CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
        ' Close the key.
        RegCloseKey handle
    End If
End Function

' Delete a registry key.
' Under Windows NT it doesn't work if the key contains subkeys.

Sub DeleteRegistryKey(ByVal hKey As Long, ByVal KeyName As String)
    RegDeleteKey hKey, KeyName
End Sub

' Read a Registry value.
' Use KeyName = "" for the default value.
' Supports only DWORD, SZ, and BINARY value types.

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, ByVal KeyType As Integer, _
    Optional DefaultValue As Variant = Empty) As Variant

    Dim handle As Long, resLong As Long
    Dim resString As String, length As Long
    Dim resBinary() As Byte
    
    ' Prepare the default result.
    GetRegistryValue = DefaultValue
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    
    Select Case KeyType
        Case REG_DWORD
            ' Read the value, use the default if not found.
            If RegQueryValueEx(handle, ValueName, 0, REG_DWORD, _
                resLong, 4) = 0 Then
                GetRegistryValue = resLong
            End If
        Case REG_SZ
            length = 1024: resString = Space$(length)
            If RegQueryValueEx(handle, ValueName, 0, REG_SZ, _
                ByVal resString, length) = 0 Then
                ' If value is found, trim characters in excess.
                If length > 0 Then
                  GetRegistryValue = Left$(resString, length)
                Else
                  GetRegistryValue = ""
                End If
                
            End If
        Case REG_BINARY
            length = 4096
            ReDim resBinary(length - 1) As Byte
            If RegQueryValueEx(handle, ValueName, 0, REG_BINARY, _
                resBinary(0), length) = 0 Then
                ReDim Preserve resBinary(length - 1) As Byte
                GetRegistryValue = resBinary()
            End If
        Case Else
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    RegCloseKey handle
End Function

' Write / Create a Registry value.
' Use KeyName = "" for the default value.
' Supports only DWORD, SZ, and BINARY value types.

Sub SetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, ByVal ValueName As String, ByVal KeyType As Integer, value As Variant)
    Dim handle As Long, lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte, length As Long
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Sub
    
    Select Case KeyType
        Case REG_DWORD
            lngValue = value
            RegSetValueEx handle, ValueName, 0, KeyType, lngValue, 4
        Case REG_SZ
            strValue = value
            RegSetValueEx handle, ValueName, 0, KeyType, ByVal strValue, Len(strValue)
        Case REG_BINARY
            binValue = value
            length = UBound(binValue) - LBound(binValue) + 1
            RegSetValueEx handle, ValueName, 0, KeyType, binValue(LBound(binValue)), length
    End Select
    
    ' Close the key.
    RegCloseKey handle
End Sub

' Delete a value.

Sub DeleteRegistryValue(ByVal hKey As Long, ByVal KeyName As String, ByVal ValueName As String)
    Dim handle As Long
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Sub
    ' Delete the value.
    RegDeleteValue handle, ValueName
    ' Close the handle.
    RegCloseKey handle
End Sub

' Enumerate registry keys under a given key, returns an array of strings.

Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As String()
    Dim handle As Long, index As Long, length As Long
    ReDim result(0 To 100) As String
    Dim FileTimeBuffer(100) As Byte
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all case the subsequent functions use hKey
        hKey = handle
    End If
    
    For index = 0 To 999999
        ' Make room in the array.
        If index > UBound(result) Then
            ReDim Preserve result(index + 99) As String
        End If
        length = 260                   ' Max length for a key name.
        result(index) = Space$(length)
        If RegEnumKey(hKey, index, result(index), length) Then Exit For
        result(index) = Left$(result(index), InStr(result(index), vbNullChar) - 1)
    Next
   
    ' Close the key, if it was actually opened.
    If handle Then RegCloseKey handle
        
    ' Trim unused items in the array.
    ReDim Preserve result(index - 1) As String
    EnumRegistryKeys = result()
End Function

' Enumerate registry values under a given key,
' returns a two dimensional array of Variant (row 0 for value names, row 1 for value contents)

Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As Variant()
    Dim handle As Long, index As Long, valueType As Long
    Dim name As String, nameLen As Long
    Dim lngValue As Long, strValue As String, dataLen As Long
    
    ReDim result(0 To 1, 0 To 100) As Variant

    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all case the subsequent functions use hKey
        hKey = handle
    End If
    
    For index = 0 To 999999
        ' Make room in the array.
        If index > UBound(result, 2) Then
            ReDim Preserve result(0 To 1, index + 99) As Variant
        End If
        nameLen = 260                   ' Max length for a key name.
        name = Space$(nameLen)
        dataLen = 4096
        ReDim binValue(0 To dataLen - 1) As Byte
        If RegEnumValue(hKey, index, name, nameLen, ByVal 0&, valueType, binValue(0), dataLen) Then Exit For
        result(0, index) = Left$(name, nameLen)
        
        Select Case valueType
            Case REG_DWORD
                ' Copy the first 4 bytes into a long variable
                CopyMemory lngValue, binValue(0), 4
                result(1, index) = lngValue
            Case REG_SZ
                ' Convert the result to a string.
                result(1, index) = Left$(StrConv(binValue(), vbUnicode), dataLen - 1)
            Case Else
                ' In all other cases, just copy the array of bytes.
                ReDim Preserve binValue(0 To dataLen - 1) As Byte
                result(1, index) = binValue()
        End Select
    Next
   
    ' Close the key, if it was actually opened.
    If handle Then RegCloseKey handle
        
    ' Trim unused items in the array.
    ReDim Preserve result(0 To 1, index - 1) As Variant
    EnumRegistryValues = result()
End Function

' You can use this function to decipher error messages from the API.

Function SystemMessage(ApiErrorCode As Long) As String
    Dim buffer As String, length As Long
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    
    buffer = Space$(1024)
    length = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0, ApiErrorCode, 0, buffer, Len(buffer), ByVal 0)
    SystemMessage = Left$(buffer, length)
    
End Function
