VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
        Private mvarAccountHeadID           As Long     'local copy
        Private mvarAccountCode             As String   'local copy
        Private mvarAccountHead             As String   'local copy
        Private mvarAccountTypeID           As Integer
        Private mvarSecondaryAccountHead    As Boolean  'local copy
        Private mvarOpeningBalance          As Double   'local copy
        
        Private mvarMajorAccountHeadID      As Long     'local copy
        Private mvarMajorAccountHeadCode    As String   'local copy
        Private mvarMajorAccountHead        As String   'local copy
        Private mvarMajorAccountTypeID      As Integer
        
        Private mvarMinorAccountHeadID      As Long 'local copy
        Private mvarMinorAccountHeadCode    As String 'local copy
        Private mvarMinorAccountHead        As String 'local copy
        Private mvarMinorAccountTypeID      As Integer
        
        Private mvarGroupID                 As Long
        Private mvarGroup                   As String
        Private mvarType                    As Integer
        Private mvarLocalBodyID             As Long
        Private mvarFinancialYearID         As Long
        Private mvarDebitOrCredit           As Boolean
        Private mvarAlias                   As String
        Private mvarSecondaryHeadName       As String
        Private mprimaryHeadID              As Integer
     
        Friend Sub subClearVariables()
            mvarAccountHeadID = -1
            mvarAccountCode = ""
            mvarAccountHead = ""
            mvarSecondaryAccountHead = False
            mvarOpeningBalance = 0#
            mvarMajorAccountHeadID = -1
            mvarMajorAccountHeadCode = ""
            mvarMajorAccountHead = ""
            mvarMajorAccountTypeID = -1
            mvarMinorAccountHeadID = -1
            mvarMinorAccountHead = ""
            
            mvarGroupID = -1
            mvarGroup = ""
            
            mvarType = 0
            mvarLocalBodyID = -1
            mvarFinancialYearID = -1
            mvarDebitOrCredit = False
            mvarAlias = ""
            
            mvarSecondaryHeadName = ""
            mprimaryHeadID = -1
        End Sub
        
        Public Function LedgerPosting(ArrTran As Variant, ArrChild As Variant, intTransactionID As Long)
            'Positing
        End Function
        
        Public Function GetLedgerBalance(mAcHeadID As Long, Optional mToDate = Null, Optional mFunctionaryID = Null, Optional mFunctionId = Null)
            Dim objdb As New clsDB
            Dim mCn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim arrInput As Variant
            
            objdb.SetConnection mCn
            Rec.CursorLocation = adUseClient
            arrInput = Array(mAcHeadID, mToDate, mFunctionaryID, mFunctionId)
            
            Set Rec = objdb.ExecuteSP("spGetLedgerBalance", arrInput, , , mCn, adCmdStoredProc)
            If Not (Rec.BOF And Rec.EOF) Then
                GetLedgerBalance = Rec!Balance
            End If
            Rec.Close
        End Function
        Public Sub SetMajorAccountHead(mMajorID As Long)
            Dim mDB As New clsDB
            Dim mCn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            
            subClearVariables
            mDB.SetConnection mCn
            Rec.CursorLocation = adUseClient
            Set Rec = GetRecordSet("spGetMajorAccountHeadDetails " & mMajorID)
            If Not (Rec.BOF And Rec.EOF) Then
                mvarMajorAccountHeadID = Rec!intMajorAccountHeadID
                mvarMajorAccountHeadCode = Rec!vchMajorAccountHeadCode
                mvarMajorAccountHead = Rec!vchMajorAccountHead
                mvarMajorAccountTypeID = Rec!tinType
            End If
            Rec.Close
        End Sub
        
        Public Sub SetMinorAccountHead(mMinorID As Long)
            Dim mDB As New clsDB
            Dim mCn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            
            subClearVariables
            mDB.SetConnection mCn
            
            Rec.CursorLocation = adUseClient
            Set Rec = GetRecordSet("spGetMinorAccountHeadDetails " & mMinorID)
            If Not (Rec.BOF And Rec.EOF) Then
                mvarMinorAccountHeadID = Rec!intMinorAccountHeadID
                mvarMinorAccountHeadCode = Rec!vchMinorAccountHeadCode
                mvarMinorAccountHead = Rec!vchMinorAccountHead
                mvarMajorAccountHeadID = Rec!intMajorAccountHeadID
                mvarMinorAccountTypeID = Rec!tinType
            End If
            Rec.Close
            
        End Sub
        
        Friend Sub SetAccounts(mAccHeadId As Variant)
            '-------------------------------------------'
            '  mAccHeadID will be either AccHeadID or
            '  AccountCode
            '-------------------------------------------'
            
            Dim mDB As New clsDB
            Dim mCn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            
            subClearVariables
            mDB.SetConnection mCn
            Rec.CursorLocation = adUseClient
            
            'Modified by Aswathi on 03/03/2008
            
            If VarType(mAccHeadId) = vbString Then
                    mSql = "SELECT * "
                    mSql = mSql & " FROM faAccountHeads LEFT JOIN "
                    mSql = mSql & " faMajorAccountHeads ON faAccountHeads.intMajorAccountHeadID= faMajorAccountHeads.intMajorAccountHeadID LEFT JOIN "
                    mSql = mSql & " faMinorAccountHeads ON faAccountHeads.intMinorAccountHeadID= faMinorAccountHeads.intMinorAccountHeadID  "
                    mSql = mSql & " Where faAccountHeads.vchAccountHeadCode = '" & mAccHeadId & "'"
            Else
                    mSql = "SELECT * "
                    mSql = mSql & " FROM faAccountHeads LEFT JOIN "
                    mSql = mSql & " faMajorAccountHeads ON faAccountHeads.intMajorAccountHeadID= faMajorAccountHeads.intMajorAccountHeadID LEFT JOIN "
                    mSql = mSql & " faMinorAccountHeads ON faAccountHeads.intMinorAccountHeadID= faMinorAccountHeads.intMinorAccountHeadID  "
                    mSql = mSql & " Where faAccountHeads.intAccountHeadID = " & mAccHeadId
            End If
            
            Rec.Open mSql, mCn, adOpenKeyset, adLockReadOnly
            If Rec.RecordCount > 0 Then
                mvarAccountHeadID = Rec!intAccountHeadID
                mvarAccountCode = Rec!vchAccountHeadCode
                mvarAccountHead = Rec!vchAccountHead
                mvarAccountTypeID = Rec!tinType
                mvarSecondaryAccountHead = Rec!tinSecondaryAccountFlag
                If IsNull(Rec!fltOpeningBalance) Then
                    mvarOpeningBalance = 0
                Else
                    mvarOpeningBalance = Rec!fltOpeningBalance
                End If
                mvarMajorAccountHeadID = Rec!intMajorAccountHeadID
                mvarMinorAccountHeadID = Rec!intMinorAccountHeadID
                mvarMajorAccountHead = Rec!vchMajorAccountHead
                mvarMajorAccountHeadCode = Rec!vchMajorAccountHeadCode
                mvarMinorAccountHead = Rec!vchMinorAccountHead
                mvarMinorAccountHeadCode = Rec!vchMinorAccountHeadCode
                If IsNull(Rec!intGroupID) Then
                    mvarGroupID = 0
                Else
                    mvarGroupID = Rec!intGroupID
                End If
'                If IsNull(Rec!vchGroup) Then
'                    mvarGroup = ""
'                Else
'                    mvarGroup = Rec!vchGroup
'                End If
                If IsNull(Rec!tinType) Then
                    mvarType = 0
                Else
                    mvarType = Rec!tinType
                End If
                If IsNull(Rec!intLocalBodyID) Then
                    mvarLocalBodyID = 0
                Else
                    mvarLocalBodyID = Rec!intLocalBodyID
                End If
                If IsNull(Rec!intFinancialYearID) Then
                    mvarFinancialYearID = 0
                Else
                    mvarFinancialYearID = Rec!intFinancialYearID
                End If
                If IsNull(Rec!tinDebitOrCredit) Then
                    mvarDebitOrCredit = False
                Else
                    mvarDebitOrCredit = Rec!tinDebitOrCredit
                End If
                If IsNull(Rec!vchAlias) Then
                    mvarAlias = ""
                Else
                    mvarAlias = Rec!vchAlias
                End If
                If IsNull(Rec!intPrimaryHeadID) Then
                    mprimaryHeadID = 0
                Else
                    mprimaryHeadID = Rec!intPrimaryHeadID
                End If
            End If
            Rec.Close
            
            
        End Sub

        Public Function FindAccountByHead(mAccHead As String) As Boolean

            '-------------------------------------------'
            '  mAccHeadID will be either AccHeadID or
            '  AccountCode
            '-------------------------------------------'
            Dim mDB As New clsDB
            Dim mCn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            
            subClearVariables
            mDB.SetConnection mCn
            Rec.CursorLocation = adUseClient
            mSql = "Select * From faAccountHeads  Where vchAccountHead = '" & mAccHead & "'"
            Rec.Open mSql, mCn, adOpenStatic, adLockReadOnly
            If Rec.RecordCount > 0 Then
                mvarAccountCode = Rec!vchAccountHeadCode
                mvarAccountHead = Rec!vchAccountHead
                mvarAccountHeadID = Rec!intAccountHeadID
                FindAccountByHead = True
            Else
                FindAccountByHead = False
            End If
            
            Rec.Close
        End Function
        
        Public Sub SetAccountID(ByVal AccHeadID As Long)
            SetAccounts AccHeadID
        End Sub
        
        Public Function SetAccountCode(ByVal AccHeadCode As String)
            SetAccounts AccHeadCode
        End Function

        Public Property Let MinorAccountHeadID(ByVal vData As Long)
            mvarMinorAccountHeadID = vData
        End Property
       
        Public Property Get MinorAccountHeadID() As Long
            MinorAccountHeadID = mvarMinorAccountHeadID
        End Property
        
        Public Property Let MinorAccountHeadCode(ByVal vData As String)
            mvarMinorAccountHeadCode = vData
        End Property
        
        Public Property Get MinorAccountHeadCode() As String
            MinorAccountHeadCode = mvarMinorAccountHeadCode
        End Property
        
        Public Property Let MinorAccountHead(ByVal vData As String)
            mvarMinorAccountHead = vData
        End Property
        
        Public Property Get MinorAccountHead() As String
            MinorAccountHead = mvarMinorAccountHead
        End Property
        
        Public Property Let AccountType(ByVal vData As Long)
            mvarAccountTypeID = vData
        End Property
        
        Public Property Get AccountType() As Long
           AccountType = mvarAccountTypeID
        End Property
        Public Property Let MinorAccountType(ByVal vData As Long)
            mvarMinorAccountTypeID = vData
        End Property
        
        Public Property Get MinorAccountType() As Long
            MinorAccountType = mvarMinorAccountTypeID
        End Property
        
        Public Property Let MajorAccountHeadID(ByVal vData As Long)
            mvarMajorAccountHeadID = vData
        End Property
        
        Public Property Get MajorAccountHeadID() As Long
            MajorAccountHeadID = mvarMajorAccountHeadID
        End Property
        
        Public Property Let MajorAccountHeadCode(ByVal vData As String)
            mvarMajorAccountHeadCode = vData
        End Property
        
        Public Property Get MajorAccountHeadCode() As String
            MajorAccountHeadCode = mvarMajorAccountHeadCode
        End Property
        
        Public Property Let MajorAccountHead(ByVal vData As String)
            mvarMajorAccountHead = vData
        End Property
        
        Public Property Get MajorAccountHead() As String
            MajorAccountHead = mvarMajorAccountHead
        End Property
        
        Public Property Let MajorAccountType(ByVal vData As Long)
            mvarMajorAccountTypeID = vData
        End Property
        
        Public Property Get MajorAccountType() As Long
            MajorAccountType = mvarMajorAccountTypeID
        End Property
        
        Public Property Let OpeningBalance(ByVal vData As Double)
            mvarOpeningBalance = vData
        End Property
        
        Public Property Get OpeningBalance() As Double
            OpeningBalance = mvarOpeningBalance
        End Property
        
        Public Property Let SecondaryAccountHead(ByVal vData As Boolean)
            mvarSecondaryAccountHead = vData
        End Property
        
        Public Property Get SecondaryAccountHead() As Boolean
            SecondaryAccountHead = mvarSecondaryAccountHead
        End Property
        
        Public Property Let AccountHead(ByVal vData As String)
            mvarAccountHead = vData
        End Property
        
        Public Property Get AccountHead() As String
            AccountHead = mvarAccountHead
        End Property
        
        Public Property Let AccountCode(ByVal vData As String)
            mvarAccountCode = vData
        End Property
        
        Public Property Get AccountCode() As String
            AccountCode = mvarAccountCode
        End Property
        
        Public Property Let AccountHeadID(ByVal vData As Long)
            mvarAccountHeadID = vData
        End Property
        
        Public Property Get AccountHeadID() As Long
            AccountHeadID = mvarAccountHeadID
        End Property
        Public Property Let GroupID(ByVal vData As Long)
            mvarGroupID = vData
        End Property
        
        Public Property Get GroupID() As Long
            GroupID = mvarGroupID
        End Property
        
        Public Property Let Group(ByVal vData As String)
            mvarGroup = vData
        End Property
        
        Public Property Get Group() As String
            Group = mvarGroup
        End Property
        
        Public Property Let mType(ByVal vData As Integer)
            mvarType = vData
        End Property
        
        Public Property Get mType() As Integer
            mType = mvarType
        End Property
        Public Property Let LocalBodyID(ByVal vData As Long)
            mvarLocalBodyID = vData
        End Property
        
        Public Property Get LocalBodyID() As Long
            LocalBodyID = mvarLocalBodyID
        End Property

        Public Property Let FinancialYearID(ByVal vData As Long)
            mvarFinancialYearID = vData
        End Property
        
        Public Property Get FinancialYearID() As Long
            FinancialYearID = mvarFinancialYearID
        End Property

        Public Property Let DebitOrCredit(ByVal vData As Boolean)
            mvarDebitOrCredit = vData
        End Property
        
        Public Property Get DebitOrCredit() As Boolean
            DebitOrCredit = mvarDebitOrCredit
        End Property
        
        Public Property Let Alias(ByVal vData As String)
            mvarAlias = vData
        End Property
        
        Public Property Get Alias() As String
            Alias = mvarAlias
        End Property

        Public Property Let primaryID(ByVal vData As Integer)
            mprimaryHeadID = vData
        End Property
        
        Public Property Get primaryID() As Integer
            primaryID = mprimaryHeadID
        End Property
