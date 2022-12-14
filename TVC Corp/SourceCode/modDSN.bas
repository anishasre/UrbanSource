Attribute VB_Name = "modDSN"
Option Explicit
'Public Const ODBC_ADD_DSN = 1      ' Add data source
'Public Const ODBC_REMOVE_DSN = 3   ' Delete data source


'Adding and Deleting DSNs
'The following routines show how to add and delete ODBC DSNs programatically.
'Note, further information on DSNs can be found in the following download:
'----DSN Declarations--------
Public Enum eDBType
    FileBased
    ServerBased
End Enum

Public Type tDSNAttrib
    Type As eDBType                 'FileBased (eg Access) or ServerBased (eg. SQL Server)
    Server As String                'Database Server
    Description As String           'Database description
    Dsn As String                   'The DSN Name
    Driver As String                'The Drive name
    Database As String              'Name or path of database
    UserID As String                'The UserID
    PassWord As String              'The User Password
    TrustedConnection As Boolean    'If True ignore the UserID and Password as will us NT
    SystemDSN As Boolean            'If True creates a system DSN, else creates a user DSN.
End Type

Private Const ODBC_ADD_DSN = 1
Private Const ODBC_CONFIG_DSN = 2
Private Const ODBC_REMOVE_DSN = 3
Private Const ODBC_ADD_SYS_DSN = 4
Private Const ODBC_CONFIG_SYS_DSN = 5
Private Const ODBC_REMOVE_SYS_DSN = 6
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Private Declare Function SQLInstallerError Lib "ODBCCP32.DLL" (ByVal iError As Long, ByVal pfErrorCode As Long, ByVal lpszErrorMsg As String, ByVal cbErrorMsgMax As Long, pcbErrorMsg As Long) As Long
     
     
'Purpose     :  Creates a new DSN
'Inputs      :  tAttributes             A type containing the input parameters for the DSN.
'                                       Look in either "C:\WINNT\Odbc.ini" or the registry under "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI"
'                                       for typical details.
'Outputs     :  Returns an empty string if successful, else returns an error message
'Notes       :  If TrustedConnection is set to False, then you must supply a valid UID
'               and PWD (username and password), otherwise the DSN will not be created.
'               When specify a username and password (i.e. when TrustedConnection = False) the
'               connection details will be verified before the DSN is created. If the details
'               are invalid the DSN will not be created. When using a TrustedConnection, no
'               such checking is carried out before the DSN is created.
'Revisions   :
'Assumptions :

Public Function DSNCreate_B4Change(tAttributes As tDSNAttrib) As String
    Const clMaxErrors As Long = 8
    Dim lRet As Long, sError As String, lLen As Long, lErrorCode As Long
    Dim sAttributes As String, bSuccess As Boolean, lThisMessage As Long
    
    'On Error Resume Next
    If tAttributes.Type = FileBased Then
        'File based database
        sAttributes = "DBQ=" & tAttributes.Database & vbNullChar
    Else
        'Server based database
        sAttributes = "Server=" & tAttributes.Server & vbNullChar
        sAttributes = sAttributes & "DATABASE=" & tAttributes.Database & vbNullChar
    End If
    
    'Name of DSN
    sAttributes = sAttributes & "DSN=" & tAttributes.Dsn & vbNullChar
    
    If Len(tAttributes.Description) Then
        'Description
        sAttributes = sAttributes & "DESCRIPTION=" & tAttributes.Description & vbNullChar
    End If
    
    If tAttributes.TrustedConnection Then
        'Use Windows NT Authentication
        '(will only validate the username and password when connection to the database)
        sAttributes = sAttributes & "Trusted_Connection=Yes" & vbNullChar
    Else
        'Specify a username and password (must specify a valid username and password)
        If Len(tAttributes.UserID) Then
            sAttributes = sAttributes & "UID=" & tAttributes.UserID & vbNullChar
        End If
        
        If Len(tAttributes.PassWord) Then
            sAttributes = sAttributes & "PWD=" & tAttributes.PassWord & vbNullChar
        End If
    End If
    If tAttributes.SystemDSN Then
        'Create a system DSN (visible to all users and services)
        bSuccess = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, tAttributes.Driver, sAttributes)
    Else
        'Create a user DSN (visible to the current users)
        bSuccess = SQLConfigDataSource(0&, ODBC_ADD_DSN, tAttributes.Driver, sAttributes)
    End If
    
    If bSuccess = False Then
        'Failed to create DSN, return error message
        sError = String(255, 0)
        For lThisMessage = 1 To clMaxErrors
            lRet = SQLInstallerError(lThisMessage, lErrorCode, sError, 255&, lLen)
            If lRet = 0 Then
                'Add error messages together
                DSNCreate_B4Change = DSNCreate_B4Change & Left(sError, lLen) & vbNewLine
            Else
                'No more error messages
                Exit For
            End If
        Next
    Else
        'Success
        DSNCreate_B4Change = ""
    End If
End Function

Public Function DSNCreate(tAttributes As tDSNAttrib) As String
    Const clMaxErrors As Long = 8
    Dim lRet As Long, sError As String, lLen As Long, lErrorCode As Long
    Dim sAttributes As String, bSuccess As Boolean, lThisMessage As Long
    
    'On Error Resume Next
    If tAttributes.Type = FileBased Then
        'File based database
        sAttributes = "DBQ=" & tAttributes.Database & Chr$(0)
    Else
        'Server based database
        sAttributes = "Server=" & tAttributes.Server & Chr$(0)
        sAttributes = sAttributes & "DATABASE=" & tAttributes.Database & Chr$(0)
    End If
    
    'Name of DSN
    sAttributes = sAttributes & "DSN=" & tAttributes.Dsn & Chr$(0)
    
    If Len(tAttributes.Description) Then
        'Description
        sAttributes = sAttributes & "DESCRIPTION=" & tAttributes.Description & Chr$(0)
    End If
    sAttributes = sAttributes & "Trusted_Connection=No" & Chr$(0)
    
    
    '    If tAttributes.TrustedConnection Then
    '        'Use Windows NT Authentication
    '        '(will only validate the username and password when connection to the database)
    '        sAttributes = sAttributes & "Trusted_Connection=Yes" & Chr$(0)
    '    Else
    '        'Specify a username and password (must specify a valid username and password)
    '        sAttributes = sAttributes & "Trusted_Connection=No" & Chr$(0)
    '        sAttributes = sAttributes & "Uid=FAUser" & Chr(0)
    '        If Len(tAttributes.UserID) Then
    '            sAttributes = sAttributes & "UID=" & tAttributes.UserID & Chr$(0)
    '        End If
    '
    '        If Len(tAttributes.PassWord) Then
    '            sAttributes = sAttributes & "PWD=" & tAttributes.PassWord & Chr$(0)
    '        End If
    '    End If
    If tAttributes.SystemDSN Then
        'Create a system DSN (visible to all users and services)
        bSuccess = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, tAttributes.Driver, sAttributes)
    Else
        'Create a user DSN (visible to the current users)
        bSuccess = SQLConfigDataSource(0&, ODBC_ADD_DSN, tAttributes.Driver, sAttributes)
    End If
    
    If bSuccess = False Then
        'Failed to create DSN, return error message
        sError = String(255, 0)
        For lThisMessage = 1 To clMaxErrors
            lRet = SQLInstallerError(lThisMessage, lErrorCode, sError, 255&, lLen)
            If lRet = 0 Then
                'Add error messages together
                DSNCreate = DSNCreate & Left(sError, lLen) & vbNewLine
            Else
                'No more error messages
                Exit For
            End If
        Next
    Else
        'Success
        DSNCreate = ""
    End If
End Function
'Purpose     :  Deletes an existing DSN
'Inputs      :  tAttributes             A type containing the input parameters of the DSN.
'                                       Look in either "C:\WINNT\Odbc.ini" or the registry under "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI"
'                                       for typical details.
'               [bSystemDSN]            If True deletes as system DSN, else deletes a user DSN.
'Outputs     :  Returns True if successful
'Notes       :
'Revisions   :
'Assumptions :

Public Function DSNDelete(sDSN As String, sDriver As String, Optional bSystemDSN As Boolean = False) As Boolean
    Dim lRet As Long
    Dim sAttributes As String
    
    On Error Resume Next
    sAttributes = "DSN=" & sDSN & vbNullChar
    If bSystemDSN Then
        DSNDelete = SQLConfigDataSource(0&, ODBC_REMOVE_DSN, sDriver, sAttributes)
    Else
        DSNDelete = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sAttributes)
    End If
End Function

'Demonstration routine
Sub Test()
    Dim tDSNDetails As tDSNAttrib, sError As String
    
'---Add an Access DSN
    With tDSNDetails
        .Database = "C:\vbusers.mdb"
        .Driver = "Microsoft Access Driver (*.mdb)"
        .PassWord = ""
        .UserID = "Admin"
        .Dsn = "TestDSN"
        .Description = "A Test Database"
        .Type = FileBased
    End With

    sError = DSNCreate(tDSNDetails)
    If Len(sError) = 0 Then
        MsgBox "Created user DSN"
        'Delete the new DSN
        If DSNDelete(tDSNDetails.Dsn, tDSNDetails.Driver) Then
            MsgBox "Deleted New DSN"
        Else
            MsgBox "Failed to Delete New DSN"
        End If
    Else
        MsgBox "Failed to Create DSN... " & vbNewLine & sError
    End If
    
'---Add an SQL Server DSN
    With tDSNDetails
        .Database = "Pubs"
        .Driver = "SQL Server"
        .Server = "MyServer"
        .TrustedConnection = True    'Use NT authentication
        .PassWord = ""
        .UserID = ""
        .Dsn = "TestDSN2"
        .Description = "A Test Database2"
        .Type = ServerBased
        .SystemDSN = True           'Create a System DSN
    End With

    sError = DSNCreate(tDSNDetails)
    If Len(sError) = 0 Then
        MsgBox "Created system DSN"
        'Delete the new DSN
        If DSNDelete(tDSNDetails.Dsn, tDSNDetails.Driver) Then
            MsgBox "Deleted New DSN"
        Else
            MsgBox "Failed to Delete New DSN"
        End If
    Else
        MsgBox "Failed to Create DSN... " & vbNewLine & sError
    End If
End Sub
