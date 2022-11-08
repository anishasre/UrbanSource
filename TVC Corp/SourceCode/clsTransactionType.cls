VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransactionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
    
    'local variable(s) to hold property value(s)
    Private mvarTransactionTypeID       As Long 'local copy
    Private mvarTransactionType         As String 'local copy
    Private mvarExternalApplicationID   As Long 'local copy
    Private mvarExternalModuleID        As Long 'local copy
    Private mvarFundID                  As Long 'local copy
    Private mvarintGroup                As Long 'local copy
    Private mvarvchGroup                As String 'local copy
    Private mvarchSubGroup              As String 'local copy
    Private mintSourceFundID            As Variant
    Private mvchSourceOfFund            As Variant
    Private mFunctionID                 As Variant
    
    Public Property Let vchGroup(ByVal vData As String)
        mvarvchGroup = vData
    End Property
    
    Public Property Get vchGroup() As String
        vchGroup = mvarvchGroup
    End Property
    
    Public Property Let intGroup(ByVal vData As Long)
        mvarintGroup = vData
    End Property
    
    Public Property Get intGroup() As Long
        intGroup = mvarintGroup
    End Property
    
    Public Property Let SourceFundID(mData As Variant)
        mintSourceFundID = mData
    End Property
    
    Public Property Get SourceFundID() As Variant
        SourceFundID = mintSourceFundID
    End Property
    
    Public Property Let SourceOfFund(mData As Variant)
        mvchSourceOfFund = mData
    End Property
    
    Public Property Get SourceOfFund() As Variant
        SourceOfFund = mvchSourceOfFund
    End Property
    
    Public Property Let FunctionID(mData As Variant)
        mFunctionID = mData
    End Property
    
    Public Property Get FunctionID() As Variant
        FunctionID = mFunctionID
    End Property
    Public Sub SetTransactionType(Optional mTranTypeID As Long)
        Dim objDb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        
        If IsNumeric(mTranTypeID) Then
            mvarTransactionTypeID = mTranTypeID
        End If
        
        If val(mvarTransactionTypeID) > 0 Then
            arrInput = Array(mvarTransactionTypeID)
            Set Rec = objDb.ExecuteSP("spGetTransactionType", arrInput)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarTransactionType = Rec!vchTransactionType
                mvarExternalApplicationID = IIf(IsNull(Rec!intExternalApplicationID), -1, Rec!intExternalApplicationID)
                mvarExternalModuleID = IIf(IsNull(Rec!intExternalModuleID), -1, Rec!intExternalModuleID)
                mvarFundID = IIf(IsNull(Rec!intFundID), -1, Rec!intFundID)
            Else
                mvarTransactionType = ""
                mvarExternalApplicationID = -1
                mvarExternalModuleID = -1
                mvarFundID = -1
            End If
            Rec.Close
        End If
        
    End Sub
    
    Public Property Let FundID(ByVal vData As Long)
        mvarFundID = vData
    End Property
    
    Public Property Get FundID() As Long
        FundID = mvarFundID
    End Property
    
    Public Property Let ExternalModuleID(ByVal vData As Long)
        mvarExternalModuleID = vData
    End Property
    
    Public Property Get ExternalModuleID() As Long
        ExternalModuleID = mvarExternalModuleID
    End Property
    
    Public Property Let ExternalApplicationID(ByVal vData As Long)
        mvarExternalApplicationID = vData
    End Property
    
    Public Property Get ExternalApplicationID() As Long
        ExternalApplicationID = mvarExternalApplicationID
    End Property
    
    Public Property Let TransactionType(ByVal vData As String)
        mvarTransactionType = vData
    End Property
    
    Public Property Get TransactionType() As String
        TransactionType = mvarTransactionType
    End Property
    
    Public Property Let TransactionTypeID(ByVal vData As Long)
        mvarTransactionTypeID = vData
    End Property
    
    Public Property Get TransactionTypeID() As Long
        TransactionTypeID = mvarTransactionTypeID
    End Property
    
    Public Property Let SubGroupID(ByVal vData As Long)
        mvarchSubGroup = vData
    End Property
    
    Public Property Get SubGroupID() As Long
        SubGroupID = mvarchSubGroup
    End Property
    
    Public Function GetTransactionAccountHeads(Optional ByVal intTransactionTypeID As Integer) As Recordset
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mRecAccHead As New ADODB.Recordset
        Dim mVarrIn(0) As Variant
        
        If Not IsMissing(intTransactionTypeID) Then
            mvarTransactionTypeID = intTransactionTypeID
        End If
        mVarrIn(0) = mvarTransactionTypeID
        If (objDb.SetConnection(mCnn)) Then
            Set mRecAccHead = objDb.ExecuteSP("spGetTransactionAccountHeads", mVarrIn, , , mCnn, adCmdStoredProc)
        Else
            Set mRecAccHead = Nothing
        End If
       Set GetTransactionAccountHeads = mRecAccHead
    End Function
    
    Public Function GetTransactionTypeID(mTransactionType As String) As Variant
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        
        mSql = "Select * From faTransactionType where vchTransactionType = '" & mTransactionType & "'"
        objDb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            GetTransactionTypeID = Rec!intTransactionTypeID
        Else
            GetTransactionTypeID = Null
        End If
        Rec.Close
        Set Rec = Nothing
        Set mCnn = Nothing
    End Function
    
    Public Sub SetSourceOfFund(intTrTypeID As Integer)
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDb As New clsDB
            Dim mSql As String
            
            If objDb.SetConnection(mCnn) Then
                mSql = "Select * from faTransactionType "
                mSql = mSql + " Inner Join suSourceOfFund On suSourceOfFund.intSourceFundID = faTransactionType.intSourceFundID "
                mSql = mSql + " Where intTransactionTypeID = " & intTrTypeID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    SourceFundID = IIf(IsNull(Rec!intSourceFundID), Null, Rec!intSourceFundID)
                    SourceOfFund = IIf(IsNull(Rec!vchSourceFundName), Null, Rec!vchSourceFundName)
                    mFunctionID = IIf(IsNull(Rec!intFunctionID), Null, Rec!intFunctionID)
                End If
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
