VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    
    Private intAgreementID          As Variant
    Private vchAgreementNo          As Variant
    Private dtAgreementDate         As Variant
    Private numProjectID            As Variant
    Private vchProjectName          As Variant
    Private vchProjectSlNo          As Variant
    Private dtDueDateToStart        As Variant
    Private dtActualStartedDate     As Variant
    Private dtDueDateOfCompletion   As Variant
    Private dtActualCompletedDate   As Variant
    Private intOrderID              As Variant
    Private vchOrderNo              As Variant
    Private dtOrderDate             As Variant
    Private vchWorkTitle            As Variant
    Private dtWorkDate              As Variant
    Private intAssetID              As Variant
    Private intAssetHeadID          As Variant
    Private vchAssetName            As Variant
    Private fltPAC                  As Variant
    Private intSubLedgerTypeID      As Variant
    Private vchSubLedgerType        As Variant
    Private numSubLedgerID          As Variant
    Private vchSubLedger            As Variant
    Private intLBID                 As Variant
    Private intFinancialYearID      As Variant
    Private numUserID               As Variant
    Private numApproverID           As Variant
    Private dtApprovedDate          As Variant
    Private tnyStatus               As Variant

    Public Property Let AgreementID(ByVal mData As Variant)
        intAgreementID = mData
    End Property
    Public Property Get AgreementID() As Variant
        AgreementID = intAgreementID
    End Property
    Public Property Let AgreementNo(ByVal mData As Variant)
        vchAgreementNo = mData
    End Property
     
    Public Property Get AgreementNo() As Variant
        AgreementNo = vchAgreementNo
    End Property
    Public Property Let AgreementDate(ByVal mData As Variant)
        dtAgreementDate = mData
    End Property
    Public Property Get AgreementDate() As Variant
        AgreementDate = dtAgreementDate
    End Property
    Public Property Let ProjectID(ByVal mData As Variant)
        numProjectID = mData
    End Property
    Public Property Get ProjectID() As Variant
        ProjectID = numProjectID
    End Property
    Public Property Let ProjectName(ByVal mData As Variant)
        vchProjectName = mData
    End Property
    Public Property Get ProjectName() As Variant
        ProjectName = vchProjectName
    End Property
    Public Property Let ProjectSlNo(ByVal mData As Variant)
        vchProjectSlNo = mData
    End Property
    Public Property Get ProjectSlNo() As Variant
        ProjectSlNo = vchProjectSlNo
    End Property
    
    Public Property Let DueDateToStart(ByVal mData As Variant)
        dtDueDateToStart = mData
    End Property
    Public Property Get DueDateToStart() As Variant
        DueDateToStart = dtDueDateToStart
    End Property
    Public Property Let ActualStartedDate(ByVal mData As Variant)
        dtActualStartedDate = mData
    End Property
    Public Property Get ActualStartedDate() As Variant
        ActualStartedDate = dtActualStartedDate
    End Property
    Public Property Let DueDateOfCompletion(ByVal mData As Variant)
        DueDateOfCompletion = mData
    End Property
    Public Property Get DueDateOfCompletion() As Variant
        DueDateOfCompletion = dtDueDateOfCompletion
    End Property
    Public Property Let ActualCompletedDate(ByVal mData As Variant)
        dtActualCompletedDate = mData
    End Property
    Public Property Get ActualCompletedDate() As Variant
        ActualCompletedDate = dtActualCompletedDate
    End Property
    Public Property Let OrderID(ByVal mData As Variant)
        intOrderID = mData
    End Property
    Public Property Get OrderID() As Variant
        OrderID = intOrderID
    End Property
    Public Property Let OrderNo(ByVal mData As Variant)
        vchOrderNo = mData
    End Property
    Public Property Get OrderNo() As Variant
        OrderNo = vchOrderNo
    End Property
    Public Property Let OrderDate(ByVal mData As Variant)
        dtOrderDate = mData
    End Property
    Public Property Get OrderDate() As Variant
       OrderDate = dtOrderDate
    End Property
    Public Property Let WorkTitle(ByVal mData As Variant)
        vchWorkTitle = mData
    End Property
    Public Property Get WorkTitle() As Variant
       WorkTitle = vchWorkTitle
    End Property
    
    Public Property Let WorkDate(ByVal mData As Variant)
        dtWorkDate = mData
    End Property
    Public Property Get WorkDate() As Variant
       WorkDate = dtWorkDate
    End Property
    Public Property Let AssetID(ByVal mData As Variant)
        intAssetID = mData
    End Property
    Public Property Get AssetID() As Variant
       AssetID = intAssetID
    End Property
    Public Property Let AssetHeadID(ByVal mData As Variant)
        intAssetHeadID = mData
    End Property
    Public Property Get AssetHeadID() As Variant
       AssetHeadID = intAssetHeadID
    End Property
    Public Property Let AssetName(ByVal mData As Variant)
        vchAssetName = mData
    End Property
    Public Property Get AssetName() As Variant
       AssetName = vchAssetName
    End Property
    Public Property Let PAC(ByVal mData As Variant)
        fltPAC = mData
    End Property
    Public Property Get PAC() As Variant
       PAC = fltPAC
    End Property
    Public Property Let SubLedgerTypeID(ByVal mData As Variant)
        intSubLedgerTypeID = mData
    End Property
    Public Property Get SubLedgerTypeID() As Variant
       SubLedgerTypeID = intSubLedgerTypeID
    End Property
    Public Property Let SubLedgerType(ByVal mData As Variant)
        vchSubLedgerType = mData
    End Property
    Public Property Get SubLedgerType() As Variant
       SubLedgerType = vchSubLedgerType
    End Property
    Public Property Let SubLedgerID(ByVal mData As Variant)
        numSubLedgerID = mData
    End Property
    Public Property Get SubLedgerID() As Variant
       SubLedgerID = numSubLedgerID
    End Property
    Public Property Let SubLedger(ByVal mData As Variant)
        vchSubLedger = mData
    End Property
    Public Property Get SubLedger() As Variant
       SubLedger = vchSubLedger
    End Property
    Public Property Let LBID(ByVal mData As Variant)
        intLBID = mData
    End Property
    Public Property Get LBID() As Variant
       LBID = intLBID
    End Property
    Public Property Let FinancialYearID(ByVal mData As Variant)
        intFinancialYearID = mData
    End Property
    Public Property Get FinancialYearID() As Variant
       FinancialYearID = intFinancialYearID
    End Property
    Public Property Let UserID(ByVal mData As Variant)
        numUserID = mData
    End Property
    Public Property Get UserID() As Variant
       UserID = numUserID
    End Property
    Public Property Let ApproverID(ByVal mData As Variant)
        numApproverID = mData
    End Property
    Public Property Get ApproverID() As Variant
       ApproverID = numApproverID
    End Property
    Public Property Let ApprovedDate(ByVal mData As Variant)
        dtApprovedDate = mData
    End Property
    Public Property Get ApprovedDate() As Variant
       ApprovedDate = dtApprovedDate
    End Property
    Public Property Let Status(ByVal mData As Variant)
        tnyStatus = mData
    End Property
    Public Property Get Status() As Variant
       Status = tnyStatus
    End Property
    Public Sub SetAgreements(AgreementID As Double)
    On Error GoTo err:
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New Connection
        Dim mSql        As String
        If objdb.SetConnection(mCnn) Then
            mSql = "Select faAgreements.*,chvProjectName as vchProjectName,vchTitle as SubLedger,vchSubLedgerType,chvProjectSlNo "
            mSql = mSql + " From faAgreements  Inner Join suProjectDetails ON faAgreements.numProjectID=suProjectDetails.decProjectID"
            mSql = mSql + " Inner Join faSubSidiaryAccountHeads ON faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID=faAgreements.numSubledgerID"
            mSql = mSql + " Inner Join faSubLedgerTypes On faSubLedgerTypes.intSubLedgerTypeID=faAgreements.intSubLedgerTypeID"
            mSql = mSql + " Where intAgreementID=" & AgreementID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                intAgreementID = Rec!intAgreementID
                vchAgreementNo = Rec!vchAgreementNo
                dtAgreementDate = IIf(IsNull(Rec!dtAgreementDate), "", Rec!dtAgreementDate)
                numProjectID = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
                vchProjectName = IIf(IsNull(Rec!vchProjectName), "", Rec!vchProjectName)
                vchProjectSlNo = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
                dtDueDateToStart = IIf(IsNull(Rec!dtDueDateToStart), "", Rec!dtDueDateToStart)
                dtActualStartedDate = IIf(IsNull(Rec!dtActualStartedDate), "", Rec!dtActualStartedDate)
                dtDueDateOfCompletion = IIf(IsNull(Rec!dtDueDateOfCompletion), "", Rec!dtDueDateOfCompletion)
                dtActualCompletedDate = IIf(IsNull(Rec!dtActualCompletedDate), "", Rec!dtActualCompletedDate)
                intOrderID = IIf(IsNull(Rec!intOrderID), "", Rec!intOrderID)
                vchOrderNo = IIf(IsNull(Rec!vchOrderNo), "", Rec!vchOrderNo)
                dtOrderDate = IIf(IsNull(Rec!dtOrderDate), "", Rec!dtOrderDate)
                vchWorkTitle = IIf(IsNull(Rec!vchWorkTitle), "", Rec!vchWorkTitle)
                dtWorkDate = IIf(IsNull(Rec!dtWorkDate), "", Rec!dtWorkDate)
                intAssetID = IIf(IsNull(Rec!intAssetID), "", Rec!intAssetID)
                intAssetHeadID = IIf(IsNull(Rec!intAssetHeadID), "", Rec!intAssetHeadID)
                vchAssetName = IIf(IsNull(Rec!vchAssetName), "", Rec!vchAssetName)
                fltPAC = IIf(IsNull(Rec!fltPAC), "", Rec!fltPAC)
                intSubLedgerTypeID = IIf(IsNull(Rec!intSubLedgerTypeID), "", Rec!intSubLedgerTypeID)
                vchSubLedgerType = IIf(IsNull(Rec!vchSubLedgerType), "", Rec!vchSubLedgerType)
                numSubLedgerID = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                vchSubLedger = IIf(IsNull(Rec!SubLedger), "", Rec!SubLedger)
                intLBID = IIf(IsNull(Rec!intLBID), "", Rec!intLBID)
                intFinancialYearID = IIf(IsNull(Rec!intFinancialYearID), "", Rec!intFinancialYearID)
                numUserID = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
                numApproverID = IIf(IsNull(Rec!numApproverID), "", Rec!numApproverID)
                dtApprovedDate = IIf(IsNull(Rec!dtApprovedDate), "", Rec!dtApprovedDate)
                tnyStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
                
            End If
        End If
        Exit Sub
err:
        Debug.Print (Error$)
    End Sub

