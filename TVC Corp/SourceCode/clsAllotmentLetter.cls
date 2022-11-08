VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAllotmentLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    
    Private intAllotmentID              As Variant
    Private vchAllotmentNo              As Variant
    Private dtAllotmentDate             As Variant
    Private intImplementingOfficersID   As Variant
    Private vchImplementingOfficer      As Variant
    Private intSourceOfFundID           As Variant
    Private intCategoryID               As Variant
    Private intTreasuryID               As Variant
    Private intInstalmentNo             As Variant
    Private intCrAccountHeadID          As Variant
    Private vchTreasuryCode             As Variant
    Private vchTreasuryName             As Variant
    Private intFunctionaryID            As Variant
    Private intFunctionID               As Variant
    Private intGrossAccountHeadID       As Variant
    Private fltAmount                   As Variant
    Private fltRequisitionAmount        As Variant
    Private numProjectID                As Variant
    Private vchPublicGrantHead          As Variant
    Private vchPublicBudgetHead         As Variant
    Private numIssuingAuthority         As Variant
    Private numPreparedBy               As Variant
    Private dtOfEntry                   As Variant
    Private intAgreementID              As Variant
    Private intLocalBodyID              As Variant
    Private intFinancialYearID          As Variant
    Private tnyStatus                   As Variant
    Private tnyCancelledFlag            As Variant
    Private vchSourceFundName           As Variant
    Private vchCategoryName             As Variant
    Private chvEngSecType               As Variant
    Private vchFunctionary              As Variant
    Private vchFunction                 As Variant
    Private vchProjectNo                As Variant
    Private chvLBCode                   As Variant
    Private chvLBNameEnglish            As Variant
    Private tnyTypeID                   As Variant
    
    Public Property Let AllotmentID(ByVal mData As Long)
        intAllotmentID = mData
    End Property
    Public Property Get AllotmentID() As Long
        AllotmentID = intAllotmentID
    End Property
    
    Public Property Let AllotmentNo(ByVal mData As String)
        vchAllotmentNo = mData
    End Property
    Public Property Get AllotmentNo() As String
        AllotmentNo = vchAllotmentNo
    End Property
    
    Public Property Let AllotmentDate(ByVal mData As Date)
        dtAllotmentDate = mData
    End Property
    Public Property Get AllotmentDate() As Date
        AllotmentDate = dtAllotmentDate
    End Property
    
    Public Property Let ImplementingOfficersID(ByVal mData As Long)
        intImplementingOfficersID = mData
    End Property
    Public Property Get ImplementingOfficersID() As Long
        ImplementingOfficersID = intImplementingOfficersID
    End Property
    
     Public Property Let ImplementingOfficer(ByVal mData As String)
        vchImplementingOfficer = mData
    End Property
    Public Property Get ImplementingOfficer() As String
        ImplementingOfficer = vchImplementingOfficer
    End Property
    
    Public Property Let SourceOfFundID(ByVal mData As Long)
    intSourceOfFundID = mData
    End Property
    Public Property Get SourceOfFundID() As Long
        SourceOfFundID = intSourceOfFundID
    End Property
    
    Public Property Let SourceOfFund(ByVal mData As String)
        vchSourceFundName = mData
    End Property
    Public Property Get SourceOfFund() As String
        SourceOfFund = vchSourceFundName
    End Property
    Public Property Let CategoryID(ByVal mData As Long)
        intCategoryID = mData
    End Property
    Public Property Get CategoryID() As Long
        CategoryID = intCategoryID
    End Property
    
    Public Property Let mNewModeID(ByVal mData As Long)
        intTreasuryID = mData
    End Property
    Public Property Get mNewModeID() As Long
        mNewModeID = intTreasuryID
    End Property
    
    Public Property Let Category(ByVal mData As String)
        vchCategoryName = mData
    End Property
    Public Property Get Category() As String
        Category = vchCategoryName
    End Property
    Public Property Let InstalmentNo(ByVal mData As Long)
        intInstalmentNo = mData
    End Property
    Public Property Get InstalmentNo() As Long
        InstalmentNo = intInstalmentNo
    End Property
    
    Public Property Let CrAccountHeadID(ByVal mData As Long)
        intCrAccountHeadID = mData
    End Property
    Public Property Get CrAccountHeadID() As Long
        CrAccountHeadID = intCrAccountHeadID
    End Property
    
    Public Property Let TreasuryCode(ByVal mData As String)
        vchTreasuryCode = mData
    End Property
    Public Property Get TreasuryCode() As String
        TreasuryCode = vchTreasuryCode
    End Property
    
    Public Property Let TreasuryName(ByVal mData As String)
        vchTreasuryName = mData
    End Property
    Public Property Get TreasuryName() As String
        TreasuryName = vchTreasuryName
    End Property
    
    Public Property Let FunctionaryID(ByVal mData As Long)
        intFunctionaryID = mData
    End Property
    Public Property Get FunctionaryID() As Long
        FunctionaryID = intFunctionaryID
    End Property
    
    Public Property Let Functionary(ByVal mData As String)
        vchFunctionary = mData
    End Property
    Public Property Get Functionary() As String
        Functionary = vchFunctionary
    End Property
    
    Public Property Let FunctionID(ByVal mData As Long)
        intFunctionID = mData
    End Property
    Public Property Get FunctionID() As Long
        FunctionID = intFunctionID
    End Property
    
    Public Property Let GrossAccountHeadID(ByVal mData As Long)
        intGrossAccountHeadID = mData
    End Property
    Public Property Get GrossAccountHeadID() As Long
        GrossAccountHeadID = intGrossAccountHeadID
    End Property
    
    Public Property Let Amount(ByVal mData As Double)
        fltAmount = mData
    End Property
    Public Property Get Amount() As Double
        Amount = fltAmount
    End Property
    
    Public Property Let RequisitionAmount(ByVal mData As Double)
        fltRequisitionAmount = mData
    End Property
    
    Public Property Get RequisitionAmount() As Double
        RequisitionAmount = fltRequisitionAmount
    End Property
    
    Public Property Let ProjectID(ByVal mData As Double)
        numProjectID = mData
    End Property
    Public Property Get ProjectID() As Double
        ProjectID = numProjectID
    End Property
    
    Public Property Let ProjectNo(ByVal mData As Double)
        vchProjectNo = mData
    End Property
    Public Property Get ProjectNo() As Double
        ProjectNo = vchProjectNo
    End Property
    
    Public Property Let PublicGrantHead(ByVal mData As String)
        vchPublicGrantHead = mData
    End Property
    Public Property Get PublicGrantHead() As String
        PublicGrantHead = vchPublicGrantHead
    End Property
    
    Public Property Let PublicBudgetHead(ByVal mData As String)
        vchPublicBudgetHead = mData
    End Property
    Public Property Get PublicBudgetHead() As String
        PublicBudgetHead = vchPublicBudgetHead
    End Property
    
    Public Property Let IssuingAuthority(ByVal mData As Variant)
        numIssuingAuthority = mData
    End Property
    Public Property Get IssuingAuthority() As Variant
        IssuingAuthority = numIssuingAuthority
    End Property
    
    Public Property Let PreparedBy(ByVal mData As Double)
        numPreparedBy = mData
    End Property
    Public Property Get PreparedBy() As Double
        PreparedBy = numPreparedBy
    End Property
    
    Public Property Let DateOfEntry(ByVal mData As Date)
        dtOfEntry = mData
    End Property
    Public Property Get DateOfEntry() As Date
        DateOfEntry = dtOfEntry
    End Property
    
    Public Property Let AgreementID(ByVal mData As Variant)
        intAgreementID = mData
    End Property
    Public Property Get AgreementID() As Variant
        AgreementID = intAgreementID
    End Property
    
    Public Property Let LocalBodyID(ByVal mData As Long)
        intLocalBodyID = mData
    End Property
    Public Property Get LocalBodyID() As Long
        LocalBodyID = intLocalBodyID
    End Property
    
    Public Property Let FinancialYearID(ByVal mData As Long)
        intFinancialYearID = mData
    End Property
    Public Property Get FinancialYearID() As Long
        FinancialYearID = intFinancialYearID
    End Property
    
    Public Property Let Status(ByVal mData As Integer)
        tnyStatus = mData
    End Property
    Public Property Get Status() As Integer
        Status = tnyStatus
    End Property
    Public Property Let TypeID(ByVal mData As Integer)
        tnyTypeID = mData
    End Property
    Public Property Get TypeID() As Integer
        TypeID = tnyTypeID
    End Property
    Public Property Let CancelledFlag(ByVal mData As Integer)
        tnyCancelledFlag = mData
    End Property
    Public Property Get CancelledFlag() As Integer
        CancelledFlag = tnyCancelledFlag
    End Property
    
    Public Property Let SectorType(ByVal mData As String)
        chvEngSecType = mData
    End Property
    Public Property Get SectorType() As String
        SectorType = chvEngSecType
    End Property
    
    Public Property Let LBCode(ByVal mData As String)
        chvLBCode = mData
    End Property
    Public Property Get LBCode() As String
        LBCode = chvLBCode
    End Property
    
    Public Property Let LBName(ByVal mData As String)
        chvLBNameEnglish = mData
    End Property
    Public Property Get LBName() As String
        LBName = chvLBNameEnglish
    End Property
    
    Public Function SetAllotmentLetter(AllotmentID As Integer)
     'On Error GoTo Err:
     Dim mCnn    As New ADODB.Connection
     Dim objDB   As New clsDB
     Dim Rec     As New ADODB.Recordset
     Dim mSQL    As String
     
     objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
     mSQL = "Select * From faAllotmentLetters"
     mSQL = mSQL + " Inner Join suSourceOfFund On faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID"
     mSQL = mSQL + " Inner Join faFunctionaries On faAllotmentLetters.intImplementingOfficersID = faFunctionaries.intFunctionaryID"
     mSQL = mSQL + " Inner Join suProjectDetails On faAllotmentLetters.numProjectID = suProjectDetails.decProjectID"
     mSQL = mSQL + " Inner Join suSectorType On suProjectDetails.intSectorTypeID = suSectorType.intSecTypeID"
     mSQL = mSQL + " Inner Join DB_Masters..GM_LocalBody On faAllotmentLetters.intLocalBodyID = DB_Masters..GM_LocalBody.intLBID"
     mSQL = mSQL + " Where intAllotmentID = " & AllotmentID
     
     mSQL = "Select * From faAllotmentLetters"
     mSQL = mSQL + " Inner Join suSourceOfFund On faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID"
     mSQL = mSQL + " LEFT Join faFunctionaries On faAllotmentLetters.intImplementingOfficersID = faFunctionaries.intFunctionaryID"
     
     'mSQL = mSQL + " Inner Join suProjectDetails On faAllotmentLetters.numProjectID = suProjectDetails.decProjectID"
     'mSQL = mSQL + " LEFT JOIN suSectorType On suProjectDetails.intSectorTypeID = suSectorType.intSecTypeID"
     'mSQL = mSQL + " Inner Join DB_Masters..GM_LocalBody On faAllotmentLetters.intLocalBodyID = DB_Masters..GM_LocalBody.intLBID"
     mSQL = mSQL + " Where intAllotmentID = " & AllotmentID
     
     Rec.Open mSQL, mCnn
     If Not (Rec.EOF And Rec.BOF) Then
         intAllotmentID = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
         AllotmentNo = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
         AllotmentDate = IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate)
         'ImplementingOfficersID = -1
         ImplementingOfficersID = IIf(IsNull(Rec!intImplementingOfficersID), -1, Rec!intImplementingOfficersID)
         SourceOfFundID = IIf(IsNull(Rec!intSourceOfFundID), "", Rec!intSourceOfFundID)
         SourceOfFund = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
         InstalmentNo = IIf(IsNull(Rec!intInstalmentNo), "-1", Rec!intInstalmentNo)
         CrAccountHeadID = IIf(IsNull(Rec!intCrAccountHeadID), "", Rec!intCrAccountHeadID)
         TreasuryCode = IIf(IsNull(Rec!vchTreasuryCode), "", Rec!vchTreasuryCode)
         TreasuryName = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
         FunctionaryID = IIf(IsNull(Rec!intFunctionaryID), "-1", Rec!intFunctionaryID)
         Functionary = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
         FunctionID = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
         GrossAccountHeadID = IIf(IsNull(Rec!intGrossAccountHeadID), "", Rec!intGrossAccountHeadID)
         Amount = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
         ProjectID = IIf(IsNull(Rec!numProjectID), "-1", Rec!numProjectID)
         PublicGrantHead = IIf(IsNull(Rec!vchPublicGrantHead), "", Rec!vchPublicGrantHead)
         PublicBudgetHead = IIf(IsNull(Rec!vchPublicBudgetHead), "", Rec!vchPublicBudgetHead)
         IssuingAuthority = IIf(IsNull(Rec!numIssuingAuthority), "", Rec!numIssuingAuthority)
         PreparedBy = IIf(IsNull(Rec!numPreparedBy), "", Rec!numPreparedBy)
         DateOfEntry = IIf(IsNull(Rec!dtOfEntry), "", Rec!dtOfEntry)
         AgreementID = IIf(IsNull(Rec!intAgreementID), "", Rec!intAgreementID)
         LocalBodyID = IIf(IsNull(Rec!intLocalBodyID), "", Rec!intLocalBodyID)
         FinancialYearID = IIf(IsNull(Rec!intFinancialYearID), "", Rec!intFinancialYearID)
         Status = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
         CancelledFlag = IIf(IsNull(Rec!tnyCancelledFlag), "", Rec!tnyCancelledFlag)
         SectorType = "" 'IIf(IsNull(Rec!chvEngSecType), "", Rec!chvEngSecType)
         LBCode = "" 'IIf(IsNull(Rec!chvLBCode), "", Rec!chvLBCode)
         LBName = gbLBName  'IIf(IsNull(Rec!chvLBNameEnglish), "", Rec!chvLBNameEnglish)
     End If
     Rec.Close
'Err:
'     Debug.Print (Error$)
    End Function
    
    Public Function SetAllotment(AllotmentID As Integer)
     'On Error GoTo Err:
     Dim mCnn    As New ADODB.Connection
     Dim objDB   As New clsDB
     Dim Rec     As New ADODB.Recordset
     Dim mSQL    As String
     
     objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
     mSQL = "Select *,faAllotments.intID[AllotmentID], faAllotments.intAccountHeadID GrossHeadID From faAllotments"
     mSQL = mSQL + " Inner Join faSubsidiaryAccountHeads On faAllotments.intImplementingOfficersID = faSubsidiaryAccountHeads.intSubsidiaryAccountHeadID"
     mSQL = mSQL + " Inner Join suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID"
     mSQL = mSQL + " Left Join faTransactionCategory On faAllotments.intFundCategoryID = faTransactionCategory.intCategoryID"
    ' mSQL = mSQL + " Inner Join faFunctionaries On faAllotments.intImplementingOfficersID = faFunctionaries.intFunctionaryID"
     mSQL = mSQL + " LEFT Join suProjectDetails On faAllotments.numProjectID = suProjectDetails.decProjectID"
     mSQL = mSQL + " LEFT Join suSectorType On suProjectDetails.intSectorTypeID = suSectorType.intSecTypeID"
     'mSQL = mSQL + " Inner Join DB_Masters..GM_LocalBody On faAllotmentLetters.intLocalBodyID = DB_Masters..GM_LocalBody.intLBID"
     mSQL = mSQL + " Where faAllotments.intID = " & AllotmentID
     Rec.Open mSQL, mCnn
     If Not (Rec.EOF And Rec.BOF) Then
         intAllotmentID = IIf(IsNull(Rec!AllotmentID), "", Rec!AllotmentID)
         AllotmentNo = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
         dtAllotmentDate = IIf(IsNull(Rec!dtAllotmentDate), Null, Rec!dtAllotmentDate)
         ImplementingOfficersID = IIf(IsNull(Rec!intImplementingOfficersID), "", Rec!intImplementingOfficersID)
         ImplementingOfficer = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
         SourceOfFundID = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
         SourceOfFund = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
         
         CategoryID = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
         Category = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
         
         InstalmentNo = IIf(IsNull(Rec!tnyInstallmentNo), 0, Rec!tnyInstallmentNo)
        ' CrAccountHeadID = IIf(IsNull(rec!intAccountHeadID), 0, rec!intAccountHeadID)
         TreasuryCode = IIf(IsNull(Rec!vchTreasuryCode), "", Rec!vchTreasuryCode)
         TreasuryName = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
         FunctionaryID = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
         'Functionary = IIf(IsNull(rec!vchFunctionary), "", rec!vchFunctionary)
         FunctionID = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
         GrossAccountHeadID = IIf(IsNull(Rec!GrossHeadID), "", Rec!GrossHeadID)
         fltAmount = IIf(IsNull(Rec!fltAuthorizedAmt), 0, Rec!fltAuthorizedAmt)
         fltRequisitionAmount = IIf(IsNull(Rec!fltRequestedAmt), 0, Rec!fltRequestedAmt)
         ProjectID = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
         'PublicGrantHead = IIf(IsNull(rec!vchPublicGrantHead), "", rec!vchPublicGrantHead)
         'PublicBudgetHead = IIf(IsNull(rec!vchPublicBudgetHead), "", rec!vchPublicBudgetHead)
         'IssuingAuthority = IIf(IsNull(rec!numIssuingAuthority), "", rec!numIssuingAuthority)
         'PreparedBy = IIf(IsNull(rec!numPreparedBy), "", rec!numPreparedBy)
         'DateOfEntry = IIf(IsNull(rec!dtOfEntry), "", rec!dtOfEntry)
         'AgreementID = IIf(IsNull(rec!intAgreementID), "", rec!intAgreementID)
         'LocalBodyID = IIf(IsNull(rec!intLocalBodyID), "", rec!intLocalBodyID)
         FinancialYearID = IIf(IsNull(Rec!intFinancialYearID), "", Rec!intFinancialYearID)
         Status = IIf(IsNull(Rec!tnyStatus), "-1", Rec!tnyStatus)
         'CancelledFlag = IIf(IsNull(rec!tnyCancelledFlag), "", rec!tnyCancelledFlag)
         SectorType = IIf(IsNull(Rec!chvEngSecType), "", Rec!chvEngSecType)
'         LBCode = IIf(IsNull(rec!chvLBCode), "", rec!chvLBCode)
'         LBName = IIf(IsNull(rec!chvLBNameEnglish), "", rec!chvLBNameEnglish)
          TypeID = IIf(IsNull(Rec!tnyTypeID), "0", Rec!tnyTypeID)
          mNewModeID = IIf(IsNull(Rec!intTreasuryID), "0", Rec!intTreasuryID)
     End If
     Rec.Close
'Err:
'     Debug.Print (Error$)
    End Function

