VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
        Private mvarFundID          As Variant
        Private mvarFundCode        As Variant
        Private mvarFundName        As Variant
        Private mvarMajorFundID     As Variant
        Private mvarMajorFundName   As Variant
     Private Sub ClearVariables()
        mvarFundID = Null
        mvarFundCode = Null
        mvarFundName = Null
        mvarMajorFundID = Null
        mvarMajorFundName = Null
    End Sub
    Public Sub SetFund(intID As Variant)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If intID > 0 Then
            mSQL = mSQL + " SELECT * from faFunds "
            mSQL = mSQL + " WHERE faFunds.intFundID = " & intID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFundID = Rec!intFundID
                mvarFundCode = Rec!vchFundCode
                mvarFundName = Rec!vchFund
            End If
            Rec.Close
        End If
    End Sub
    
    Public Sub SetFundByCode(mFundCode As Variant)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If mFundCode <> "" Then
            mSQL = mSQL + " SELECT * from faFunds "
            mSQL = mSQL + " WHERE vchFundCode = '" & CStr(mFundCode) & "'"
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFundID = Rec!intFundID
                mvarFundCode = Rec!vchFundCode
                mvarFundName = Rec!vchFund
            End If
            Rec.Close
        End If
    End Sub
    Public Property Let MajorFundName(ByVal vData As Variant)
        mvarMajorFundName = vData
    End Property
    
    Public Property Set MajorFundName(ByVal vData As Variant)
        Set mvarMajorFundName = vData
    End Property
    
    Public Property Get MajorFundName() As Variant
        If IsObject(mvarMajorFundName) Then
            Set MajorFundName = mvarMajorFundName
        Else
            MajorFundName = mvarMajorFundName
        End If
    End Property
    
    Public Property Let MajorFundID(ByVal vData As Variant)
        mvarMajorFundID = vData
    End Property
    
    Public Property Set MajorFundID(ByVal vData As Variant)
        Set mvarMajorFundID = vData
    End Property
    
    Public Property Get MajorFundID() As Variant
        If IsObject(mvarMajorFundID) Then
            Set MajorFundID = mvarMajorFundID
        Else
            MajorFundID = mvarMajorFundID
        End If
    End Property
    
    Public Property Let FundName(ByVal vData As Variant)
        mvarFundName = vData
    End Property
    
    Public Property Set FundName(ByVal vData As Variant)
        Set mvarFundName = vData
    End Property
    
    Public Property Get FundName() As Variant
        If IsObject(mvarFundName) Then
            Set FundName = mvarFundName
        Else
            FundName = mvarFundName
        End If
    End Property
    
    Public Property Let FundCode(ByVal vData As Variant)
        mvarFundCode = vData
    End Property
    
    Public Property Set FundCode(ByVal vData As Variant)
        Set mvarFundCode = vData
    End Property
    
    Public Property Get FundCode() As Variant
        If IsObject(mvarFundCode) Then
            Set FundCode = mvarFundCode
        Else
            FundCode = mvarFundCode
        End If
    End Property
    
    Public Property Let FundID(ByVal vData As Variant)
        mvarFundID = vData
    End Property
    
    Public Property Set FundID(ByVal vData As Variant)
        Set mvarFundID = vData
    End Property
    
    Public Property Get FundID() As Variant
        If IsObject(mvarFundID) Then
            Set FundID = mvarFundID
        Else
            FundID = mvarFundID
        End If
    End Property



