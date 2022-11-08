Attribute VB_Name = "modINIfileManager"

    '=================================================================================='
    '                                   [modINIfileManager]                            '
    '=================================================================================='
    '
    '----------------------------------------------------------------------------------'
    ' Module Which Manage the INI file                                                 '
    '                                                                                  '
    ' List of Functions Implemented                                                    '
    ' Private Members   :                                                              '
    '                      StripTerminator                                             '
    ' ---------------------------------------------------------------------------------'





Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long


'---------------------------------------------------------------------------------------------------'
'   FUNCTION: StripTerminator                                                                       '
'---------------------------------------------------------------------------------------------------'
'   Returns a string without any zero terminator.  Typically,                                       '
'   this was a string returned by a Windows API call.                                               '
'                                                                                                   '
'   IN: [strString] - String to remove terminator from                                              '
'                                                                                                   '
'   Returns: The value of the string passed in minus any                                            '
'          terminating zero.                                                                        '
'---------------------------------------------------------------------------------------------------'
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'---------------------------------------------------------------------------------------------------'
' FUNCTION: ReadIniFile                                                                             '
'---------------------------------------------------------------------------------------------------'
'   Reads a value from the specified section/key of the                                             '
'   specified .INI file                                                                             '
'                                                                                                   '
'   IN:     [strIniFile] - name of .INI file to read                                                '
'           [strSection] - section where key is found                                               '
'           [strKey] - name of key to get the value of                                              '
'                                                                                                   '
'   Returns: non-zero terminated value of .INI file key                                             '
'---------------------------------------------------------------------------------------------------'
Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer
    
    'If successful read of .INI file, strip any trailing zero returned by the Windows API GetPrivateProfileString
    strBuffer = Space$(255)
    If GetPrivateProfileString(strSection, strKey, vbNullString, strBuffer, 255, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = vbNullString
    End If
End Function

Function WriteINIfile(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, ByVal strIniFile As String)
    Debug.Print strSection, strKey, strValue, strIniFile
    Call WritePrivateProfileString(strSection, strKey, strValue, strIniFile)
End Function

