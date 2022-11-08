VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
        
        Private mvarConnection As New ADODB.Connection
        Private mRecordAffected As Integer
        
        Public Property Let Connection(ByVal vData As ADODB.Connection)
            mvarConnection = vData
        End Property
        
        Public Property Set Connection(ByVal vData As ADODB.Connection)
            Set mvarConnection = vData
        End Property
        
        Public Property Get Connection() As ADODB.Connection
            If IsObject(mvarConnection) Then
            Set Connection = mvarConnection
            Else
            Connection = mvarConnection
            End If
        End Property
        
        Public Function SetConnection(ByRef mCnn As ADODB.Connection) As Boolean
                On Error GoTo err:
                If gbConnectionToFa.State = 0 Then
                    gbConnectionToFa.ConnectionString = "PROVIDER=MSDASQL;dsn=dsnFa;uid=FAUser;pwd=FAUser;database=DB_Finance;"
                    gbConnectionToFa.Open
                End If
                Set mCnn = gbConnectionToFa
                SetConnection = True
                Exit Function
err:
                SetConnection = False
        End Function
        
        Public Function CreateNewConnection(ByRef mCnn As ADODB.Connection, mSource As enuSourceString) As Boolean
                On Error GoTo err:
                mCnn.ConnectionString = GetConnectionString(mSource)
                mCnn.Open
                CreateNewConnection = True
                Exit Function
err:
                CreateNewConnection = False
        End Function
        
        Public Function funExecuteSP(ByVal strSPName As String, _
                                    Optional vValues As Variant, _
                                    Optional vOutput, _
                                    Optional varyOut, _
                                    Optional ByVal IsDML As Boolean = True, _
                                    Optional AdoCon As Connection = Nothing)
                                    
                    Dim adoCommand As New ADODB.Command
                    Dim adoRec As New ADODB.Recordset
                    Dim parParams As Parameters
                    Dim intCount As Integer
                    Dim adoLocalCon As Connection
                    Dim blnNewConnection As Boolean
                    
                    If Not AdoCon Is Nothing Then
                        Set adoLocalCon = AdoCon
                        blnNewConnection = False
                    Else
                        blnNewConnection = True
                        adoLocalCon.BeginTrans
                        On Error GoTo EHandler
                    End If
                    With adoCommand
                        Set .ActiveConnection = adoLocalCon
                        .CommandType = adCmdStoredProc
                        .CommandText = strSPName
                        Set parParams = .Parameters
                    End With
                    If Not IsMissing(vValues) Then
                        For intCount = 1 To UBound(vValues)
                            parParams(intCount).Value = vValues(intCount - 1)
                            Debug.Print parParams(intCount).Name & vbTab & parParams(intCount).Value
                        Next
                    End If
                    If Not IsDML Then
                        Set adoRec = adoCommand.Execute
                        If adoRec.EOF And adoRec.BOF Then
                            varyOut = Empty
                        Else
                            varyOut = adoRec.GetRows
                        End If
                    Else
                        adoCommand.Execute
                    End If
                    If Not IsMissing(vOutput) Then
                        ReDim vOutput(10)
                        For intCount = 0 To (parParams.count - 1 - (UBound(vValues) + 1))
                            ReDim Preserve vOutput(intCount + 1)
                            vOutput(intCount) = parParams(intCount + UBound(vValues) + 1)
                        Next intCount
                    End If
                    If blnNewConnection Then adoLocalCon.CommitTrans
                    GoTo Exitfunction
EHandler:
                    If blnNewConnection Then adoLocalCon.RollbackTrans
Exitfunction:
                    If blnNewConnection Then adoLocalCon.Close
        End Function
        Public Function funExecuteSPib(ByVal strSPName As String, _
                                       Optional vValues As Variant, _
                                       Optional vOutput, _
                                       Optional AdoCon As ADODB.Connection = Nothing) As ADODB.Recordset
                    
                    Dim adoCommand As New ADODB.Command
                    Dim intCount As Integer
                    Dim adoLocalCon As ADODB.Connection
                    Dim blnNewConnection As Boolean
                    Dim objdb As clsDB
                    
                    '-------------------------------------------'
                    '  Connection                               '
                    '-------------------------------------------'
                    If Not AdoCon Is Nothing Then
                        Set adoLocalCon = AdoCon
                        blnNewConnection = False
                    Else
                        Set objdb = New clsDB
                        blnNewConnection = True
                        objdb.SetConnection adoLocalCon
                        adoLocalCon.BeginTrans
                        'On Error GoTo EHandler
                    End If
                    '-------------------------------------------'
                    '  Command Object                           '
                    '-------------------------------------------'
                    With adoCommand
                        Set .ActiveConnection = adoLocalCon
                        .CommandType = adCmdStoredProc
                        .CommandText = strSPName
                    End With
                    '-------------------------------------------'
                    '  Executing                                '
                    '-------------------------------------------'
                    If Not IsMissing(vValues) Then
                        Set funExecuteSPib = adoCommand.Execute(, vValues)
                    Else
                        Set funExecuteSPib = adoCommand.Execute
                    End If
                    If blnNewConnection Then
                        adoLocalCon.CommitTrans
                    End If
                    GoTo Exitfunction
EHandler:
                    If blnNewConnection Then adoLocalCon.RollbackTrans
Exitfunction:
                    If blnNewConnection Then Set adoLocalCon = Nothing
        End Function
        
        Public Function ExecuteSP( _
                ByVal strSPName As String, _
                Optional aryInput As Variant, _
                Optional aryOutPut As Variant, _
                Optional ByVal IsDML As Boolean = True, _
                Optional adoCnn As ADODB.Connection = Nothing, _
                Optional adoCmdType As ADODB.CommandTypeEnum = ADODB.CommandTypeEnum.adCmdStoredProc _
                ) As ADODB.Recordset
                
                Dim ADOCmd As New ADODB.Command
                Dim mCount As Long
                Dim adoRec As New ADODB.Recordset
                
                If adoCnn Is Nothing Then
                    SetConnection adoCnn
                Else
                    If adoCnn.State = 0 Then
                        SetConnection adoCnn
                    End If
                End If
                Set ADOCmd.ActiveConnection = adoCnn
                ADOCmd.CommandType = adoCmdType
                ADOCmd.CommandText = strSPName
                If Not IsMissing(aryInput) Then
                    For mCount = 0 To UBound(aryInput)
                        If aryInput(mCount) = "" Or IsEmpty(aryInput(mCount)) Then
                            aryInput(mCount) = Null
                        End If
                        ADOCmd.Parameters(mCount + 1).Value = aryInput(mCount)
                        Debug.Print ADOCmd.Parameters(mCount + 1).Value
                    Next mCount
                End If
                adoRec.CursorType = adOpenDynamic
                If Not IsMissing(aryInput) Then
                    Set adoRec = ADOCmd.Execute(, ADOCmd.Parameters)
                Else
                    Set adoRec = ADOCmd.Execute
                End If
                Set ExecuteSP = adoRec
                If Not IsMissing(aryOutPut) Then
                    If (ExecuteSP.BOF = False And ExecuteSP.EOF = False) Then
                        aryOutPut = ExecuteSP.GetRows()
                    End If
                End If
        End Function
        Public Function SetExtDBConnection(ByRef mCnn As ADODB.Connection, mConString As String) As Boolean
                'On Error GoTo Err:
                'Debug.Print gbConnectionToFa.State
                mCnn.ConnectionString = mConString '"PROVIDER=MSDASQL;dsn=dsnFa;uid=FAUser;pwd=FAUser;database=DB_Finance;"
                mCnn.Open
                If mCnn.State = 0 Then
                    SetExtDBConnection = False
                Else
                    SetExtDBConnection = True
                End If
                Exit Function
err:
                SetExtDBConnection = False
        End Function
        Public Function GetConnectionString(mApplicationID As enuSourceString) As String
                Select Case mApplicationID
                    Case Is = 1 ' FINANCE (Saankhya Double Entry)
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnFa;uid=FAUser;pwd=FAUser;database=DB_Finance;timeout=0"
                    Case Is = 2 ' SANCHAYA
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnSanchaya;uid=FAUser;pwd=FAUser;database=DB_Sanchaya;"
                    Case Is = 3 ' SANCHAYA LITE  '''Database Removed From Connection String On 1/2/11
                        'If gbLocalBodyID = 171 Or gbLocalBodyID = 168 Or gbLocalBodyID = 205 Then
'                        If gbLocalBodyID = 205 Then
'                            GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNSanchayaLite;uid=FAUser;pwd=FAUser;"
'                        Else
                            GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnSanchaya;uid=FAUser;pwd=FAUser;"
'                        End If
                    Case Is = 4 ' SAANKHYA MASTERS
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNSaankhyaMaster;uid=Saankhya;pwd=san;database=DB_SaankhyaMasters;"
                    Case Is = 5 ' STHAPANA
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNSthapana;uid=SthapanaUser;pwd=Sthapana123;database=DB_Sthapana;"
                    Case Is = 6 ' SOOCHIKA URBAN
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=GWFlow1;uid=dbsoochika;pwd=urbansoochika;database=DB_SoochikaUrban;"
                    Case Is = 7 ' KMBR
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNKMBR;uid=Sanketham;pwd=Sanketham;database=DB_KMBR;"
                    Case Is = 8 ' SEVANA PENSIONS
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNPension;uid=FAUser;pwd=FAUser;database=DB_SevanaPension;"
                    Case Is = 9 ' SULEKHA
                        'GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNFormulation2009;uid=FAUser;pwd=FAUser;"
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnSulekha;uid=FAUser;pwd=FAUser;"
                    Case Is = 10
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=DSNSevanaLBMCommon;uid=LBUser;pwd=T1x+a7Pw;database=DB_Common;"
                    Case Is = 11
                        GetConnectionString = "PROVIDER=MSDASQL;DSN=DSNSevanaReg;uid=LBUser;pwd=T1x+a7Pw;database=DB_SevanaKiosk;"
                    Case Is = 12
                        'If gbLBID = 167 Or gbLBID = 168 Or gbLBID = 173 Or gbLBID = 243 Or gbLBID = 288 Or gbLBID = 248 Then
                            Dim objcn As Object 'Web32CR.clsGen1
                            Set objcn = CreateObject("Web32CR.clsGen1")
                            GetConnectionString = "PROVIDER=MSDASQL;" & objcn.gen_cnset("DSNSevanaRegn", 0) & "database=DB_SevanaRegistration;"
                        'Else
                        '    GetConnectionString = "PROVIDER=MSDASQL;DSN=DSNSevanaRegn;uid=LBUser;pwd=T1x+a7Pw;database=DB_SevanaRegistration;"
                        'End If
                    Case Is = 16
                        GetConnectionString = "PROVIDER=MSDASQL;DSN=Sugama;uid=SugamaUser;pwd=SugamaUser;database=DB_Sugama;"
                    Case Is = 17
                        GetConnectionString = "PROVIDER=MSDASQL;DSN=GWFlow1;uid=workflow;pwd=A+v378*R;database=DB_Soochika;"
                    Case Is = 93 ' S A A N K H Y A      B A C K U P
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnFABackup;uid=FAUser;pwd=FAUser;database=DB_FinanceBackUp;"
                    Case Is = 94 ' S A A N K H Y A  HO
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnFAHO;uid=FAUser;pwd=FAUser;database=DB_FinanceHO;"
                    Case Is = 95 ' SACHAYA HO
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnSanchayaHO;uid=FAUser;pwd=FAUser;database=DB_SanchayaHO;"
                    Case Is = 96 ' OLD SAANKHYA
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=Accounts;uid=Saankhya;pwd=SAN;database=DB_Accounts;"
                    Case Is = 97 ' OLD SAHATHA
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=Receipts;uid=sa;pwd=;database=Receipts;"
                    Case Is = 98 ' iSAANKHYA MASTERS
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsniSaankhyaMasters;uid=FAUser;pwd=FAUser;database=DB_iSaankhyaMasters;"
                    Case Is = 99 ' DB MASTERS (COMMON DATABASE)
                        GetConnectionString = "PROVIDER=MSDASQL;dsn=dsnMaster;uid=FAUser;pwd=FAUser;database=DB_Masters;"
                End Select
        End Function
        
        Private Function spDebug(sp As String, ary)
            If IsArray(ary) Then
                spDebug = sp & " "
                Dim i
                For Each i In ary
                    spDebug = spDebug & "'" & i & "',"
                Next
                spDebug = Left(spDebug, Len(spDebug) - 1)
                Debug.Print spDebug
            End If
        End Function
        Public Function FillGridCombo(ByVal vsGridName As VSFlexGrid, ByVal intColNo As Integer, SPName As String, Optional aryInput As Variant)
            '==================================================================='
            ' Added on  : 15-Jan-2008  By Jyothishkumar U.
            ' Input     :  FlexGridName,Column number of the grid,SP to use,Parameters to the sp
            ' Output    :  Fills the grid combo
            '==================================================================='
            Dim adoRec As New ADODB.Recordset
            Dim adoCnn As New ADODB.Connection
            Dim mItem As String
            mItem = ""
            If IsMissing(aryInput) Then
                Set adoRec = ExecuteSP(SPName, , , , adoCnn, adCmdStoredProc)
            Else
                Set adoRec = ExecuteSP(SPName, aryInput, , , adoCnn, adCmdStoredProc)
            End If
            While Not adoRec.EOF
                mItem = mItem & "|#" & CStr(adoRec.Fields(1)) & ";" & adoRec.Fields(0)
                adoRec.MoveNext
            Wend
            vsGridName.ColComboList(intColNo) = mItem
            mItem = ""
        End Function
        Public Function FillGridComboWithZeroIndex(ByVal vsGridName As VSFlexGrid, ByVal intColNo As Integer, SPName As String, Optional aryInput As Variant)
            '==================================================================='
            ' Added on  : 15-Jan-2008  By Jyothishkumar U.
            ' Input     :  FlexGridName,Column number of the grid,SP to use,Parameters to the sp
            ' Output    :  Fills the grid combo with a blank line
            '==================================================================='
            Dim adoRec As New ADODB.Recordset
            Dim adoCnn As New ADODB.Connection
            Dim mItem As String
            mItem = "#0;..."
            If IsMissing(aryInput) Then
                Set adoRec = ExecuteSP(SPName, , , , adoCnn, adCmdStoredProc)
            Else
                Set adoRec = ExecuteSP(SPName, aryInput, , , adoCnn, adCmdStoredProc)
            End If
            While Not adoRec.EOF
                mItem = mItem & "|#" & CStr(adoRec.Fields(1)) & ";" & adoRec.Fields(0)
                adoRec.MoveNext
            Wend
            vsGridName.ColComboList(intColNo) = mItem
            mItem = ""
        End Function
        
        Public Function ExecuteSPIb( _
                ByVal strSPName As String, _
                Optional aryInput As Variant, _
                Optional aryOutPut As Variant, _
                Optional ByVal IsDML As Boolean = True, _
                Optional adoCnn As ADODB.Connection = Nothing _
                ) As ADODB.Recordset
                
                Dim ADOCmd As New ADODB.Command
                Dim mCount As Long
                Dim adoRec As New ADODB.Recordset
                
                If adoCnn Is Nothing Then
                    SetConnection adoCnn
                Else
                    If adoCnn.State = 0 Then
                        SetConnection adoCnn
                    End If
                End If
                Set ADOCmd.ActiveConnection = adoCnn
                'ADOCmd.CommandType = adoCmdType
                ADOCmd.CommandText = strSPName
                If Not IsMissing(aryInput) Then
                    For mCount = 0 To UBound(aryInput)
                        If aryInput(mCount) = "" Or IsEmpty(aryInput(mCount)) Then
                            aryInput(mCount) = Null
                        End If
                        ADOCmd.Parameters(mCount + 1).Value = aryInput(mCount)
                        Debug.Print ADOCmd.Parameters(mCount + 1).Value
                    Next mCount
                End If
                adoRec.CursorLocation = adUseClient
                adoRec.CursorType = adOpenDynamic
                If Not IsMissing(aryInput) Then
                    Set adoRec = ADOCmd.Execute(mRecordAffected, ADOCmd.Parameters)
                Else
                    Set adoRec = ADOCmd.Execute(mRecordAffected, , adCmdStoredProc)
                End If
                Set ExecuteSPIb = adoRec
                If Not IsMissing(aryOutPut) Then
                    If (ExecuteSPIb.BOF = False And ExecuteSPIb.EOF = False) Then
                        aryOutPut = ExecuteSPIb.GetRows()
                    End If
                End If
        End Function
        
        
        Public Property Get RecordAffected() As Integer
            RecordAffected = mRecordAffected
        End Property
        
