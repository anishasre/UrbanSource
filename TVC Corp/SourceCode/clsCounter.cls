VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private mCounterID                  As Long
    Private mCounterNo                  As Integer
    Private mCounterDescription         As String
    Private mCounterIP                  As String
    Private mCounterConnected           As Boolean
    Private mCounterActive              As Boolean
    Private mCounterSectionID           As Integer
    Private mCounterSection             As String
    Private mCounterOperationModeID     As Integer
    
    Public Property Let CounterNo(ByVal vData As Integer)
        mCounterNo = vData
    End Property
    Public Property Get CounterNo() As Integer
        CounterNo = mCounterNo
    End Property
    Public Property Let CounterID(ByVal vData As Long)
        mCounterID = vData
    End Property
    Public Property Get CounterID() As Long
        CounterID = mCounterID
    End Property
    Public Property Let CounterDescription(ByVal vData As String)
        mCounterDescription = vData
    End Property
    Public Property Get CounterDescription() As String
        CounterDescription = mCounterDescription
    End Property
    Public Property Let CounterIP(ByVal vData As String)
        mCounterIP = vData
    End Property
    Public Property Get CounterIP() As String
        CounterIP = mCounterIP
    End Property
    Public Property Let CounterConnected(ByVal vData As Boolean)
        mCounterConnected = vData
    End Property
    Public Property Get CounterConnected() As Boolean
        CounterConnected = mCounterConnected
    End Property
    Public Property Let CounterActive(ByVal vData As Boolean)
        mCounterActive = vData
    End Property
    Public Property Get CounterActive() As Boolean
        CounterActive = mCounterActive
    End Property
    Public Property Get CounterSectionID() As Integer
        CounterSectionID = mCounterSectionID
    End Property
    Public Property Get CounterSection() As String
        CounterSection = mCounterSection
    End Property
    Public Property Get CounterOperationModeID() As String
        CounterOperationModeID = mCounterOperationModeID
    End Property
    
    Public Sub SetCounter(intCounterNo As Integer)
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim mRec As New ADODB.Recordset
        Dim mVarrIn(0) As Variant
        
        mVarrIn(0) = intCounterNo
        If (objDb.SetConnection(mCnn)) Then
            Set mRec = objDb.ExecuteSP("spGetCounterDetails", mVarrIn, , , mCnn, adCmdStoredProc)
            If Not mRec.EOF Then
                mCounterDescription = mRec!vchDescription
                mCounterIP = mRec!vchIPAddress
                mCounterID = mRec!intCounterID
                mCounterNo = mRec!intCounterNo
                If mRec!tnyConnectedFlag = 1 Then
                    mCounterConnected = True
                Else
                    mCounterConnected = False
                End If
                If mRec!tnySuspentedFlag = 1 Then
                    mCounterActive = False
                Else
                    mCounterActive = True
                End If
                mCounterSectionID = mRec!intSectionID
                mCounterSection = mRec!vchSectionName
            End If
        End If
    End Sub

    Public Sub SetCounterByIP(mIP As String)
        '==================================================================='
        ' Added on  :  10-Jan-2007   By Aiby
        ' Input     :  IP Address
        ' Output    :  Set's local object variables
        '==================================================================='
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        mSQL = "Select * From faCounters Where vchIPAddress = '" & mIP & "'"
        Set Rec = GetRecordSet(mSQL)
        If Not (Rec.BOF And Rec.EOF) Then
            mCounterNo = Rec!intCounterNo
            mCounterID = Rec!intCounterID
            mCounterDescription = Rec!vchDescription
            mCounterIP = Rec!vchIPAddress
            mCounterConnected = Rec!tnyConnectedFlag
            mCounterActive = Not Rec!tnySuspentedFlag
            mCounterSectionID = Rec!intSectionID
        End If
        Rec.Close
        Set mCnn = Nothing
        Set objDb = Nothing
    End Sub
    
    Public Sub CounterLogin(ByVal mCounterID As Integer, mLoggedIn As Boolean)
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        If objDb.SetConnection(mCnn) Then
            If mLoggedIn = True Then
                mCnn.Execute ("update faCounters set tnyConnectedFlag = 1 where intCounterID=" & mCounterID)
            Else
                mCnn.Execute ("update faCounters set tnyConnectedFlag = 0 where intCounterID=" & mCounterID)
            End If
        End If
    End Sub
