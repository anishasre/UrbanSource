VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBudgetCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
   
Option Explicit
    
    Private mvarBudgetCentreID      As Long
    Private mvarBudgetCentreCode    As String
    Private mvarBudgetCentre        As String
    Private mvarFunctionaryID        As Long
    Private mvarFunctionaryCode     As String
    Private mvarFunctionaryName     As String
    Private mvarFunctionID          As Long
    Private mvarFunctionCode        As String
    Private mvarFunctionName        As String
    Private mvarFieldID             As Long
    Private mvarFieldCode           As String
    Private mvarFieldName           As String
    
    Private mvarFundID             As Long
    Private mvarFundCode           As String
    Private mvarFundName           As String
    
    Private mvarFinancialYearID     As Long
    Private mvarLocalBodyID         As Variant
    Private mVarEstimatedAmount     As Double
    Private mVarUtilisedAmount      As Double
    Private mVarBalanceAmount       As Double
    
    Private Sub ClearVariables()
            mvarBudgetCentreID = -1
            mvarBudgetCentreCode = ""
            mvarBudgetCentre = ""
            mvarFunctionaryID = -1
            mvarFunctionaryCode = ""
            mvarFunctionaryName = ""
            mvarFunctionID = -1
            mvarFunctionCode = ""
            mvarFunctionName = ""
            mvarFieldID = -1
            mvarFieldCode = ""
            mvarFieldName = ""
            mvarFundID = -1
            mvarFundCode = ""
            mvarFundName = ""
            mvarFinancialYearID = -1
            mvarLocalBodyID = -1
            mVarEstimatedAmount = -1
            mVarUtilisedAmount = -1
            mVarBalanceAmount = -1
    End Sub
    
    Public Function GetBudgetAmount(ByVal intFunctionID As Long, ByVal intFunctionaryID As Long, ByVal intAccountHeadID As Long)
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        mSQL = " Select * From faBudgetCentres Inner Join"
        mSQL = mSQL & " faBudgetAccountHeads On faBudgetAccountHeads.intBudgetCentreID = faBudgetCentres.intBudgetCentreID"
        mSQL = mSQL & " Where intFunctionID = " & intFunctionID & " AND intFunctionaryID = " & intFunctionaryID
        mSQL = mSQL & " AND intAccountHeadID = " & intAccountHeadID
        
        Set Rec = GetRecordSet(mSQL)
        If Rec.RecordCount > 0 Then
            GetBudgetAmount = Format(Rec!fltEstimatedAmount, "0.00")
        Else
            GetBudgetAmount = 0#
        End If
        Rec.Close
        Set Rec = Nothing
        Set mCn = Nothing
    End Function
    
    Public Function SetBudgetCentre(ByVal strBudgetCentreCode As String) As Boolean
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        mSQL = "SELECT * "
        mSQL = mSQL & " FROM faBudgetCentres LEFT JOIN "
        mSQL = mSQL & " faFunctionaries ON faBudgetCentres.intFunctionaryID = faFunctionaries.intFunctionaryID LEFT JOIN "
        mSQL = mSQL & " faFinancialYear ON faBudgetCentres.intFinancialYearID = faFinancialYear.intFinancialYearID LEFT JOIN "
        mSQL = mSQL & " faFunctions ON faBudgetCentres.intFunctionID = faFunctions.intFunctionID LEFT JOIN "
        mSQL = mSQL & " faFields ON faBudgetCentres.intFieldID = faFields.intFieldID LEFT JOIN "
        mSQL = mSQL & " faFunds ON faBudgetCentres.intFundID = faFunds.intFundID LEFT JOIN "
        mSQL = mSQL & " faLocalBody ON faBudgetCentres.intLocalBodyID = faLocalBody.intLocalBodyID "
        mSQL = mSQL & " WHERE faBudgetCentres.vchBudgetCentreCode = '" & strBudgetCentreCode & "'"
        Rec.CursorLocation = adUseClient
        
        Set Rec = GetRecordSet(mSQL)
        If Rec.RecordCount > 0 Then
            mvarBudgetCentreID = Rec!intBudgetCentreID
            mvarBudgetCentreCode = Rec!vchBudgetCentreCode
            mvarBudgetCentre = Rec!vchBudgetCentre
            
            mvarFinancialYearID = IIf(IsNull(Rec!intFinancialYearID), -1, Rec!intFinancialYearID)
            
            mvarFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), -1, Rec!intFunctionaryID)
            mvarFunctionaryCode = IIf(IsNull(Rec!vchFunctionaryCode), "", Rec!vchFunctionaryCode)
            mvarFunctionaryName = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            
            mvarFunctionID = IIf(IsNull(Rec!intFunctionID), -1, Rec!intFunctionID)
            mvarFunctionCode = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode)
            mvarFunctionName = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            
            mvarFieldID = IIf(IsNull(Rec!intFieldID), -1, Rec!intFieldID)
            mvarFieldCode = IIf(IsNull(Rec!vchFieldCode), "", Rec!vchFieldCode)
            mvarFieldName = IIf(IsNull(Rec!vchField), "", Rec!vchField)
            
            mvarFundID = IIf(IsNull(Rec!intFundID), -1, Rec!intFundID)
            mvarFundCode = IIf(IsNull(Rec!vchFundCode), "", Rec!vchFundCode)
            mvarFundName = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
            
            mvarFinancialYearID = IIf(IsNull(Rec!intFinancialYearID), -1, Rec!intFinancialYearID)
            mvarLocalBodyID = IIf(IsNull(Rec!intLocalBodyID), -1, Rec!intLocalBodyID)
        Else
            ClearVariables
        End If
    End Function
    Public Function GetAccountHeads(ByVal intBudgetCentreID As Long) As ADODB.Recordset
        Dim Rec As New ADODB.Recordset
        Dim objDb As New clsDB
        'Set GetAccountHeads = objDb.ExecuteSP("spGetBudgetAccountHeads", Array(intBudgetCentreID))
        Set GetAccountHeads = objDb.ExecuteSP("spGetBudgetRevisionHistoryDetails", Array(intBudgetCentreID))
    End Function
    Public Function SetBudgetCentreByID(ByVal intBudgetCentreID As Long) As Boolean
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        mSQL = "SELECT * "
        mSQL = mSQL & " FROM faBudgetCentres LEFT JOIN "
        mSQL = mSQL & " faFunctionaries ON faBudgetCentres.intFunctionaryID = faFunctionaries.intFunctionaryID LEFT JOIN "
        mSQL = mSQL & " faFunctions ON faBudgetCentres.intFunctionID = faFunctions.intFunctionID LEFT JOIN "
        mSQL = mSQL & " faFields ON faBudgetCentres.intFieldID = faFields.intFieldID LEFT JOIN "
        mSQL = mSQL & " faLocalBody ON faBudgetCentres.intLocalBodyID = faLocalBody.intLocalBodyID "
        mSQL = mSQL & " WHERE faBudgetCentres.intBudgetCentreID = " & intBudgetCentreID
        
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSQL)
        If Rec.RecordCount > 0 Then
            mvarBudgetCentreID = Rec!intBudgetCentreID
            mvarBudgetCentreCode = Rec!vchBudgetCentreCode
            mvarBudgetCentre = Rec!vchBudgetCentre
            
            mvarFunctionaryID = IIf(IsNull(Rec!intFunctionID), mvarFunctionaryID, Rec!intFunctionID)
            mvarFunctionaryCode = IIf(IsNull(Rec!vchFunctionaryCode), mvarFunctionaryCode, Rec!vchFunctionaryCode)
            mvarFunctionaryName = IIf(IsNull(Rec!vchFunctionary), mvarFunctionaryName, Rec!vchFunctionary)
            
            mvarFunctionID = IIf(IsNull(Rec!intFunctionID), mvarFunctionID, Rec!intFunctionID)
            mvarFunctionCode = IIf(IsNull(Rec!vchFunctionCode), mvarFunctionCode, Rec!vchFunctionCode)
            mvarFunctionName = IIf(IsNull(Rec!vchFunction), mvarFunctionName, Rec!vchFunction)
            
            mvarFieldID = IIf(IsNull(Rec!intFieldID), mvarFieldID, Rec!intFieldID)
            mvarFieldCode = IIf(IsNull(Rec!vchFieldCode), mvarFieldCode, Rec!vchFieldCode)
            mvarFieldName = IIf(IsNull(Rec!vchField), mvarFieldName, Rec!vchField)
            
            mvarFinancialYearID = IIf(IsNull(Rec!intFinancialYearID), mvarFinancialYearID, Rec!intFinancialYearID)
            mvarLocalBodyID = IIf(IsNull(Rec!intLocalBodyID), mvarLocalBodyID, Rec!intLocalBodyID)
        Else
            ClearVariables
        End If
    End Function
    Public Function GetProjectExpenditure(numProjID As Double) As Boolean
        
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim objDb As New clsDB
        Dim mArrIn As Variant
        GetProjectExpenditure = False
        Return
        
        
        'NOT COMPLETED
        'Method to find Project Wise Utilization and Balance Amount
        
    End Function
    Public Function SetBudgetAccountHead(intAccountHeadID As Integer, Optional intFunctionaryID As Variant = Null, Optional intFunctionID As Variant = Null)
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim objDb As New clsDB
        Dim mArrIn As Variant
        Dim mFunctionary As Variant
        Dim mFunction As Variant
        mFunctionary = intFunctionaryID
        mFunction = intFunctionID
        If IsNull(intFunctionaryID) Then
            mFunctionary = CStr(mvarFunctionaryID)
        End If
        If IsNull(intFunctionID) Then
            mFunction = CStr(mvarFunctionID)
        End If
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            Exit Function
        End If
        mArrIn = Array(mFunctionary, _
                        mFunction, _
                        CStr(gbFinancialYearID), _
                        intAccountHeadID)
         Set Rec = objDb.ExecuteSP("spGetBudgetAccountHeadsAmount", mArrIn, , , mCnn)
        If Not (Rec.BOF And Rec.EOF) Then
            mvarFunctionaryID = Rec!intFunctionID
            mvarFunctionaryCode = Rec!vchFunctionaryCode
            mvarFunctionaryName = Rec!vchFunctionary
            mvarFunctionID = Rec!intFunctionID
            mvarFunctionCode = Rec!vchFunctionCode
            mvarFunctionName = Rec!vchFunction
            mvarBudgetCentreID = Rec!intBudgetCentreID
            mvarBudgetCentreCode = Rec!vchBudgetCentreCode
            mvarBudgetCentre = Rec!vchBudgetCentre
            mVarEstimatedAmount = Rec!fltEstimatedAmount
            mVarUtilisedAmount = IIf(IsNull(Rec!UtilisedAmount), 0, Rec!UtilisedAmount)
            mVarBalanceAmount = mVarEstimatedAmount - mVarUtilisedAmount
            mvarLocalBodyID = IIf(IsNull(Rec!intLBID), "", Rec!intLBID)
            mvarFinancialYearID = gbFinancialYearID
        End If
    End Function
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
    Public Property Let FieldName(ByVal vData As String)
        mvarFieldName = vData
    End Property
    Public Property Get FieldName() As String
        FieldName = mvarFieldName
    End Property
    Public Property Let FieldCode(ByVal vData As String)
        mvarFieldCode = vData
    End Property
    Public Property Get FieldCode() As String
        FieldCode = mvarFieldCode
    End Property
    Public Property Let FieldID(ByVal vData As Long)
        mvarFieldID = vData
    End Property
    Public Property Get FieldID() As Long
        FieldID = mvarFieldID
    End Property
    Public Property Let FunctionName(ByVal vData As String)
        mvarFunctionName = vData
    End Property
    Public Property Get FunctionName() As String
        FunctionName = mvarFunctionName
    End Property
    Public Property Let FunctionCode(ByVal vData As String)
        mvarFunctionCode = vData
    End Property
    Public Property Get FunctionCode() As String
        FunctionCode = mvarFunctionCode
    End Property
    Public Property Let FunctionID(ByVal vData As Long)
        mvarFunctionID = vData
    End Property
    Public Property Get FunctionID() As Long
        FunctionID = mvarFunctionID
    End Property
    Public Property Let FunctionaryName(ByVal vData As String)
        mvarFunctionaryName = vData
    End Property
    Public Property Get FunctionaryName() As String
        FunctionaryName = mvarFunctionaryName
    End Property
    Public Property Let FunctionaryCode(ByVal vData As String)
        mvarFunctionaryCode = vData
    End Property
    Public Property Get FunctionaryCode() As String
        FunctionaryCode = mvarFunctionaryCode
    End Property
    Public Property Let FunctionaryID(ByVal vData As Long)
        mvarFunctionaryID = vData
    End Property
    Public Property Get FunctionaryID() As Long
        FunctionaryID = mvarFunctionaryID
    End Property
    Public Property Let BudgetCentre(ByVal vData As String)
        mvarBudgetCentre = vData
    End Property
    Public Property Get BudgetCentre() As String
        BudgetCentre = mvarBudgetCentre
    End Property
    Public Property Let BudgetCentreCode(ByVal vData As String)
        mvarBudgetCentreCode = vData
    End Property
    Public Property Get BudgetCentreCode() As String
        BudgetCentreCode = mvarBudgetCentreCode
    End Property
    Public Property Let BudgetCentreID(ByVal vData As Long)
        mvarBudgetCentreID = vData
    End Property
    Public Property Get BudgetCentreID() As Long
        BudgetCentreID = mvarBudgetCentreID
    End Property
    Private Sub Class_Initialize()
        Call ClearVariables
    End Sub
    Public Property Let FundName(ByVal vData As String)
        mvarFundName = vData
    End Property
    Public Property Get FundName() As String
        FundName = mvarFundName
    End Property
    Public Property Let FundCode(ByVal vData As String)
        mvarFundCode = vData
    End Property
    Public Property Get FundCode() As String
        FundCode = mvarFundCode
    End Property
    Public Property Let FundID(ByVal vData As Long)
        mvarFundID = vData
    End Property
    Public Property Get FundID() As Long
        FundID = mvarFundID
    End Property
    Public Property Get EstimatedAmount() As Double
        EstimatedAmount = mVarEstimatedAmount
    End Property
    Public Property Get UtilisedAmount() As Double
        UtilisedAmount = mVarUtilisedAmount
    End Property
    Public Property Get BalanceAmount() As Double
        BalanceAmount = mVarBalanceAmount
    End Property
