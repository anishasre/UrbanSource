Attribute VB_Name = "modTest"
Option Explicit

Public Enum qryType
        RInsert = 1
        rselect = 2
        rUpdate = 3
        RDelete = 4
End Enum

Public Enum Dsn
'    DEManagement = 1
    dsnFA = 2
End Enum

Public Function gFunSetConnection(ByVal vInDSNID As Integer) As ADODB.Connection
    Dim adoCon As New ADODB.Connection
'    Select Case vInDSNID
'        Case 1
'            adocon.ConnectionString = "PROVIDER=MSDASQL;dsn=DEManagement;uid=sa;pwd=;"
'        Case 2
             adoCon.ConnectionString = "PROVIDER=MSDASQL;dsn=dsbFa;uid=FAUser;pwd=FAUser;" 'DB_PFPDEDE
'    End Select
    adoCon.Open
    Set gFunSetConnection = adoCon
End Function

Public Function ExecuteSP(ByVal strForExecute As String, _
                            ByVal qryType As Integer, _
                            ByVal ADOCmd As ADODB.CommandTypeEnum, _
                            Optional vAryIn, _
                            Optional varyOut, _
                            Optional adoConnection As ADODB.Connection)
                            
    Dim adoCon As New ADODB.Connection
    Dim adoCom As New ADODB.Command
    Dim adoRec As New ADODB.Recordset
    Dim intcnt As Integer
    Dim lpCnt1, lpCnt2 As Integer
    
    
    
    If Not IsMissing(adoConnection) Then
        Set adoCom.ActiveConnection = adoConnection
    Else: Set adoCom.ActiveConnection = gFunSetConnection(Dsn.dsnFA)
    
    End If
    
    If Not IsMissing(vAryIn) Then
        For intcnt = 0 To UBound(vAryIn)
            If vAryIn(intcnt) = "" Or IsEmpty(vAryIn(intcnt)) Then vAryIn(intcnt) = Null
        Next intcnt
    End If
    
    adoCom.CommandType = ADOCmd
    adoCom.CommandText = strForExecute
    Select Case qryType
        Case RInsert
            Set adoRec = adoCom.Execute(, vAryIn)
            If Not IsMissing(varyOut) Then
                If (adoRec.BOF = False And adoRec.EOF = False) Then
                    varyOut = adoRec.GetRows()
                End If
            End If
        Case rselect
            If IsMissing(vAryIn) Then
                Set adoRec = adoCom.Execute
            Else
                Set adoRec = adoCom.Execute(, vAryIn)
            End If
            If (adoRec.BOF = False And adoRec.EOF = False) Then varyOut = adoRec.GetRows()
            If IsArray(varyOut) Then
                For lpCnt1 = 0 To UBound(varyOut)
                    For lpCnt2 = 0 To UBound(varyOut, 2)
                        If IsNull(varyOut(lpCnt1, lpCnt2)) Then varyOut(lpCnt1, lpCnt2) = ""
                    Next lpCnt2
                Next lpCnt1
            End If
        Case rUpdate
            If IsMissing(vAryIn) Then Set adoRec = adoCom.Execute Else Set adoRec = adoCom.Execute(, vAryIn)
        Case RDelete
            If IsMissing(vAryIn) Then adoCom.Execute Else adoCom.Execute , vAryIn
    End Select
Exitfunction:
    Set adoCom = Nothing
End Function


Public Function gFunExecuteSP(ByVal strForExecute As String, ByVal qryType As Integer, ByVal ADOCmd As ADODB.CommandTypeEnum, _
                            Optional vAryIn, Optional varyOut, Optional adoConnection As ADODB.Connection)

'    Dim adoCon As New ADODB.Connection
'    Dim adoCom As New ADODB.Command
'    Dim adoRec As New ADODB.Recordset
'    Dim intcnt As Integer
'    Dim lpCnt1, lpCnt2 As Integer
'    If Not IsMissing(adoConnection) Then
'        Set adoCom.ActiveConnection = adoConnection
'    Else: Set adoCom.ActiveConnection = gFunSetConnection(Dsn.SulekhaFormulation)
'    End If
'
'    If Not IsMissing(vAryIn) Then
'        For intcnt = 0 To UBound(vAryIn)
'            If vAryIn(intcnt) = "" Or IsEmpty(vAryIn(intcnt)) Then vAryIn(intcnt) = Null
'        Next intcnt
'    End If
'
'    adoCom.CommandType = ADOCmd
'    adoCom.CommandText = strForExecute
'    Select Case qryType
'        Case RInsert
'            Set adoRec = adoCom.Execute(, vAryIn)
'            If Not IsMissing(varyOut) Then
'              If (adoRec.BOF = False And adoRec.EOF = False) Then varyOut = adoRec.GetRows()
'            End If
'        Case rselect
'            If IsMissing(vAryIn) Then Set adoRec = adoCom.Execute Else Set adoRec = adoCom.Execute(, vAryIn)
'            If (adoRec.BOF = False And adoRec.EOF = False) Then varyOut = adoRec.GetRows()
'            If IsArray(varyOut) Then
'                For lpCnt1 = 0 To UBound(varyOut)
'                    For lpCnt2 = 0 To UBound(varyOut, 2)
'                        If IsNull(varyOut(lpCnt1, lpCnt2)) Then varyOut(lpCnt1, lpCnt2) = ""
'                    Next lpCnt2
'                Next lpCnt1
'            End If
'        Case rUpdate
'            If IsMissing(vAryIn) Then Set adoRec = adoCom.Execute Else Set adoRec = adoCom.Execute(, vAryIn)
'        Case RDelete
'            If IsMissing(vAryIn) Then adoCom.Execute Else adoCom.Execute , vAryIn
'    End Select
'Exitfunction:
'    Set adoCom = Nothing
End Function






Public Function SaveRSToXML(ConnectionString As String, _
    SQLString As String, FullPath As String) As Boolean
'**************************************************
'PURPOSE: SAVE A RECORDSET TO AN XML FILE USING
'ADO 2.5

'PARAMETERS:
'ConnectionString:  Valid Connection String
'SQLString:         Valid SQL Statement for Data Source specified
'                   in ConnectionString
'FullPath:          FullPath of XMLFile to write to

'RETURNS:           True if Sucessful, false otherwise
'REQUIRES:          Installation of and reference to ADO 2.5
'EXAMPLE of SaveRsToXML and LoadRSToXML:

'Dim sConnString As String
'Dim sSQL As String
'Dim oRs As ADODB.Recordset
'Dim iCtr As Integer

'sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data
'Source=C:\MyDb.mdb"
'sSQL = "select * from MyTable"
'SaveRSToXML sConnString, sSQL, "C:\My Documents\MyRs.xml"

'Set oRs = LoadRsFromXML("C:\my documents\MyRS.xml")
'If Not oRs Is Nothing Then
'  Do While Not oRs.EOF
'    For iCtr = 0 To oRs.Fields.Count - 1
'        Debug.Print oRs.Fields(iCtr).Name & " = " _
'           & oRs.Fields(iCtr).Value & ";";
'    Next
'    Debug.Print vbCrLf
'    oRs.MoveNext
'  Loop
'  set oRs = nothing
'End If


'******************************************************

Dim oCn As New ADODB.Connection
Dim oCmd As New ADODB.Command
Dim oRs As ADODB.Recordset

On Error GoTo ErrorHandler:

oCn.ConnectionString = ConnectionString
oCn.Open
Set oCmd.ActiveConnection = oCn
oCmd.CommandText = SQLString
oCmd.CommandType = adCmdText
Set oRs = oCmd.Execute
oRs.Save FullPath, adPersistXML
SaveRSToXML = True

ErrorHandler:
    On Error Resume Next
    Set oRs = Nothing
    Set oCmd = Nothing
    If oCn.State <> 0 Then oCn.Close
    Set oCn = Nothing
    
End Function

Public Function LoadRsFromXML(FullPath As String) As _
  ADODB.Recordset

'**************************************************
'PURPOSE: LOAD A RECORDSET FROM AN XML FILE USING
'ADO 2.5.  THE XML FILE MUST HAVE BEEN SAVED
'USING SAVE METHOD OF RECORDSET OBJECT WITH adPersistXML AD
'SECOND PARAMETER

'PARAMETERS:
 'FullPath:     FullPath of XMLFile to load

'RETURNS:       Reference to a Recordset Object, or Nothing if
'               Function fails
'REQUIRES:      Installation of and reference to ADO 2.5
'EXAMPLE:       See Example for SaveRsToXML

'******************************************************

Dim oRs As New ADODB.Recordset
On Error Resume Next

If Dir(FullPath) = "" Then Exit Function
oRs.Open FullPath, "Provider=MSPersist;", adOpenForwardOnly, _
    adLockReadOnly, adCmdFile

If Err.Number = 0 Then
    Set LoadRsFromXML = oRs
End If

End Function







