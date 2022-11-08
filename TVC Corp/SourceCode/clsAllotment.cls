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

   Public Property Let AllotmentID(ByVal mData As Long)
       intAllotmentID = mData
   End Property
   Public Property Get AllotmentID() As Long
   AllotmentID = mData
   End Property

   Public Property Let AllotmentNo(ByVal mData As String)
       vchAllotmentNo = mData
   End Property
   Public Property Get AllotmentNo() As String
   AllotmentNo = mData
   End Property

   Public Property Let llotmentDate(ByVal mData As Date)
       dtAllotmentDate = mData
   End Property
   Public Property Get llotmentDate() As Date
   llotmentDate = mData
   End Property

   Public Property Let ImplementingOfficersID(ByVal mData As Long)
       intImplementingOfficersID = mData
   End Property
   Public Property Get ImplementingOfficersID() As Long
   ImplementingOfficersID = mData
   End Property

   Public Property Let SourceOfFundID(ByVal mData As Long)
       intSourceOfFundID = mData
   End Property
   Public Property Get SourceOfFundID() As Long
   SourceOfFundID = mData
   End Property

   Public Property Let InstalmentNo(ByVal mData As Long)
       intInstalmentNo = mData
   End Property
   Public Property Get InstalmentNo() As Long
   InstalmentNo = mData
   End Property

   Public Property Let CrAccountHeadID(ByVal mData As Long)
       intCrAccountHeadID = mData
   End Property
   Public Property Get CrAccountHeadID() As Long
   CrAccountHeadID = mData
   End Property

   Public Property Let TreasuryCode(ByVal mData As String)
       vchTreasuryCode = mData
   End Property
   Public Property Get TreasuryCode() As String
   TreasuryCode = mData
   End Property

   Public Property Let TreasuryName(ByVal mData As String)
       vchTreasuryName = mData
   End Property
   Public Property Get TreasuryName() As String
   TreasuryName = mData
   End Property

   Public Property Let FunctionaryID(ByVal mData As Long)
       intFunctionaryID = mData
   End Property
   Public Property Get FunctionaryID() As Long
   FunctionaryID = mData
   End Property

   Public Property Let FunctionID(ByVal mData As Long)
       intFunctionID = mData
   End Property
   Public Property Get FunctionID() As Long
   FunctionID = mData
   End Property

   Public Property Let GrossAccountHeadID(ByVal mData As Long)
       intGrossAccountHeadID = mData
   End Property
   Public Property Get GrossAccountHeadID() As Long
   GrossAccountHeadID = mData
   End Property

   Public Property Let Amount(ByVal mData As Double)
       fltAmount = mData
   End Property
   Public Property Get Amount() As Double
   Amount = mData
   End Property

   Public Property Let ProjectID(ByVal mData As Double)
       numProjectID = mData
   End Property
   Public Property Get ProjectID() As Double
   ProjectID = mData
   End Property

   Public Property Let PublicGrantHead(ByVal mData As String)
       vchPublicGrantHead = mData
   End Property
   Public Property Get PublicGrantHead() As String
   PublicGrantHead = mData
   End Property

   Public Property Let PublicBudgetHead(ByVal mData As String)
       vchPublicBudgetHead = mData
   End Property
   Public Property Get PublicBudgetHead() As String
   PublicBudgetHead = mData
   End Property

   Public Property Let IssuingAuthority(ByVal mData As Double)
       numIssuingAuthority = mData
   End Property
   Public Property Get IssuingAuthority() As Double
   IssuingAuthority = mData
   End Property

   Public Property Let PreparedBy(ByVal mData As Double)
       numPreparedBy = mData
   End Property
   Public Property Get PreparedBy() As Double
   PreparedBy = mData
   End Property

   Public Property Let fEntry(ByVal mData As Date)
       dtOfEntry = mData
   End Property
   Public Property Get fEntry() As Date
   fEntry = mData
   End Property

   Public Property Let AgreementID(ByVal mData As Long)
       intAgreementID = mData
   End Property
   Public Property Get AgreementID() As Long
   AgreementID = mData
   End Property

   Public Property Let LocalBodyID(ByVal mData As Long)
       intLocalBodyID = mData
   End Property
   Public Property Get LocalBodyID() As Long
   LocalBodyID = mData
   End Property

   Public Property Let FinancialYearID(ByVal mData As Long)
       intFinancialYearID = mData
   End Property
   Public Property Get FinancialYearID() As Long
   FinancialYearID = mData
   End Property

   Public Property Let Status(ByVal mData As Integer)
       tnyStatus = mData
   End Property
   Public Property Get Status() As Integer
   Status = mData
   End Property

   Public Property Let CancelledFlag(ByVal mData As Integer)
       tnyCancelledFlag = mData
   End Property
   Public Property Get CancelledFlag() As Integer
   CancelledFlag = mData
   End Property

