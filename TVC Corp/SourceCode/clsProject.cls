VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private decProjectID            As Variant
Private intLBID                 As Variant
Private intYearID               As Variant
Private intProjectSlNo          As Variant
Private chvProjectSerialNo      As Variant
Private chvProjectName          As Variant
Private chvProjectnameEnglish   As Variant
Private intProjCatID            As Variant
Private chvDPCOrderNo           As Variant
Private dtDPCOrderDate          As Variant
Private intPlanID               As Variant
Private vchSector               As Variant
Private intSectorTypeID         As Variant

Private vchCategory             As Variant
Private intCategoryID           As Variant
Private intSourceOfFundID       As Variant
Private vchSourceOfFund         As Variant
Private fltProjectCost          As Variant
' NEW FIELDS
Private intSubSectorID          As Variant
Private vchSubSector            As Variant
Private tnyStatus               As Variant

Private intSchemeID             As Variant

Private mSourceAsCollection As New Collection
Private mCol As Collection


'Public Property Let PropertyName (ByVal mData As String)
'    localvariable = mData
'End Property
'Public Property Get PropertyName() As String
'    PropertyName = localvariable
'End Property

Private Sub Class_Initialize()
    '
    ' C O N S T R U C T O R
    '
    decProjectID = 0
    intLBID = 0
    intYearID = 0
    intProjectSlNo = 0
    chvProjectSerialNo = ""
    chvProjectName = ""
    chvProjectnameEnglish = ""
    intProjCatID = 0
    chvDPCOrderNo = ""
    dtDPCOrderDate = Null
    intPlanID = 0
    vchSector = ""
    intSectorTypeID = 0
    vchCategory = "" 'Null
    intCategoryID = 0 'Null
    intSourceOfFundID = 0
    vchSourceOfFund = ""
    fltProjectCost = 0
    intSubSectorID = 0
    vchSubSector = ""
    tnyStatus = 0
    intSchemeID = 0
End Sub

Public Property Let ProjectID(ByVal mData As Variant)
    decProjectID = mData
End Property
Public Property Get ProjectID() As Variant
    ProjectID = decProjectID
End Property

Public Property Let LBID(ByVal mData As Long)
    intLBID = mData
End Property
Public Property Get LBID() As Long
    LBID = intLBID
End Property

Public Property Let YearID(ByVal mData As Long)
    intYearID = mData
End Property
Public Property Get YearID() As Long
    YearID = intYearID
End Property

Public Property Let ProjectSlNo(ByVal mData As Long)
    intProjectSlNo = mData
End Property
Public Property Get ProjectSlNo() As Long
    ProjectSlNo = intProjectSlNo
End Property

Public Property Let ProjectSerialNo(ByVal mData As String)
    chvProjectSerialNo = mData
End Property
Public Property Get ProjectSerialNo() As String
    ProjectSerialNo = chvProjectSerialNo
End Property

Public Property Let ProjectName(ByVal mData As String)
    chvProjectName = mData
End Property
Public Property Get ProjectName() As String
    ProjectName = chvProjectName
End Property

Public Property Let ProjectNameEnglish(ByVal mData As String)
    chvProjectnameEnglish = mData
End Property
Public Property Get ProjectNameEnglish() As String
    ProjectNameEnglish = chvProjectnameEnglish
End Property

Public Property Let ProjCatID(ByVal mData As Long)
    intProjCatID = mData
End Property
Public Property Get ProjCatID() As Long
    ProjCatID = intProjCatID
End Property

Public Property Let DPCOrderNo(ByVal mData As String)
    chvDPCOrderNo = mData
End Property
Public Property Get DPCOrderNo() As String
    DPCOrderNo = chvDPCOrderNo
End Property

Public Property Let PCOrderDate(ByVal mData As Variant)
    dtDPCOrderDate = mData
End Property
Public Property Get PCOrderDate() As Variant
    PCOrderDate = dtDPCOrderDate
End Property

Public Property Let SectorTypeID(ByVal mData As Long)
    intSectorTypeID = mData
End Property
Public Property Get SectorTypeID() As Long
    SectorTypeID = intSectorTypeID
End Property
Public Property Let PlanID(ByVal mData As Long)
    intPlanID = mData
End Property
Public Property Let ProjectCost(ByVal mData As Long)
    fltProjectCost = mData
End Property

Public Property Get PlanID() As Long
    PlanID = intPlanID
End Property
Public Property Get Sector() As String
    Sector = vchSector
End Property

Public Property Get Category() As String
    Category = vchCategory
End Property
Public Property Get CategoryID() As String
    CategoryID = intCategoryID
End Property

Public Property Get SourceOfFund() As String
    SourceOfFund = vchSourceOfFund
End Property
Public Property Get SourceOfFundID() As Long
    SourceOfFundID = intSourceOfFundID
End Property

Public Property Get ProjectCost() As Long
    ProjectCost = fltProjectCost
End Property

Public Property Let SubSector(ByVal mData As String)
    vchSubSector = mData
End Property
Public Property Get SubSector() As String
    SubSector = vchSubSector
End Property

Public Property Let SubSectorID(ByVal mData As String)
    intSubSectorID = mData
End Property
Public Property Get SubSectorID() As String
    SubSectorID = intSubSectorID
End Property
Public Property Let Status(ByVal mData As Variant)
    tnyStatus = mData
End Property
Public Property Get Status() As Variant
    Status = tnyStatus
End Property
Public Property Let SchemeID(ByVal mData As String)
    intSchemeID = mData
End Property
Public Property Get SchemeID() As String
    SchemeID = intSchemeID
End Property
Friend Sub setClearProjectVariable()
    decProjectID = ""
    intLBID = ""
    intYearID = ""
    intProjectSlNo = ""
    chvProjectSerialNo = ""
    chvProjectName = ""
    chvProjectnameEnglish = ""
    intProjCatID = ""
    chvDPCOrderNo = ""
    dtDPCOrderDate = ""
    intPlanID = ""
    intSectorTypeID = ""
    vchSector = ""
    vchCategory = ""
    intCategoryID = ""
    intSubSectorID = ""
    intSchemeID = ""
End Sub
Public Sub SetProject(numProjID As Variant, Optional YearID As Integer)
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim mLoop As Integer
    Dim mYearID As Integer
    
    
    If YearID = 0 Then 'IsNull(YearID)
        mYearID = gbFinancialYearID
    Else
        mYearID = YearID
    End If
    
    If mYearID < 2013 Then
        mSQL = " SELECT * FROM ProjectDetails "
        mSQL = mSQL + " INNER JOIN M_SubSector ON M_SubSector.intSubSecID = ProjectDetails.intSubSecID"
        mSQL = mSQL + " Where ProjectDetails.intYearID = " & mYearID & " And ProjectDetails.decProjectID = " & numProjID
        mSQL = mSQL + " Order by chvProjectSlNo"
    Else
        mSQL = " SELECT * FROM ProjectDetails "
        mSQL = mSQL + " INNER JOIN M_SubSector2013 ON M_SubSector2013.intSubSecID = ProjectDetails.intSubSecID"
        mSQL = mSQL + " Where ProjectDetails.intYearID = " & mYearID & " And ProjectDetails.decProjectID = " & numProjID
        mSQL = mSQL + " Order by chvProjectSlNo"
    End If
    If objDB.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
        Rec.Open mSQL, mCnn, adOpenStatic, adLockReadOnly
        If Not (Rec.BOF And Rec.EOF) Then
        
            decProjectID = Rec!decProjectID
            intLBID = Rec!intLBID
            intYearID = Rec!intYearID
            intProjectSlNo = Rec!intProjectSlNo
            chvProjectSerialNo = Rec!chvProjectSlNo
            chvProjectName = Rec!chvProjectName
            chvProjectnameEnglish = Rec!chvProjectNameEng
            intProjCatID = Rec!intProjCatID
            
            chvDPCOrderNo = Null
            dtDPCOrderDate = Null
            intPlanID = Null
            
            intSectorTypeID = Rec!intMajorSecID
            vchSector = Rec!intSecID
            Select Case Rec!intProjCatID
                Case Is = 1: vchCategory = "General"
                Case Is = 2: vchCategory = "SCP"
                Case Is = 3: vchCategory = "TSP"
                Case Else
                vchCategory = "General"
                
            End Select
            
            
            intCategoryID = IIf(IsNull(Rec!intProjCatID), 0, Rec!intProjCatID)
            intSubSectorID = IIf(IsNull(Rec!intSubSecID), 0, Rec!intSubSecID)
            vchSubSector = IIf(IsNull(Rec!chrSubSecCodeEng), "", Rec!chrSubSecCodeEng) 'nchvSubSector
      
            intSchemeID = IIf(IsNull(Rec!intCSSID), 0, Rec!intCSSID)
            
        End If
        Rec.Close
    Else
        MsgBox "Connection to DBSulekha doesnot exist", vbInformation, "Saankhya"
        Exit Sub
    End If
    
End Sub

Public Sub SetProject_(mdecProjectID As Variant)
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim arrInput As Variant
    
    arrInput = Array(mdecProjectID)
    objDB.SetConnection mCnn
    Set Rec = objDB.ExecuteSP("spGetProjectDetails", arrInput, , , mCnn, adCmdStoredProc)
    If Not (Rec.BOF And Rec.EOF) Then
        decProjectID = IIf(IsNull(Rec!decProjectID), "", Rec!decProjectID)
        intLBID = IIf(IsNull(Rec!intLBID), "", Rec!intLBID)
        intYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
        intProjectSlNo = IIf(IsNull(Rec!intProjectSlNo), "", Rec!intProjectSlNo)
        chvProjectSerialNo = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
        chvProjectName = IIf(IsNull(Rec!chvProjectName), "", Rec!chvProjectName)
        chvProjectnameEnglish = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
        intProjCatID = IIf(IsNull(Rec!intProjCatID), "", Rec!intProjCatID)
        chvDPCOrderNo = IIf(IsNull(Rec!chvDPCOrderNo), "", Rec!chvDPCOrderNo)
        dtDPCOrderDate = IIf(IsNull(Rec!dtDPCOrderDate), "", Rec!dtDPCOrderDate)
        intSectorTypeID = IIf(IsNull(Rec!intSectorTypeID), "", Rec!intSectorTypeID)
        intPlanID = IIf(IsNull(Rec!intPlanID), "", Rec!intPlanID)
        vchSector = IIf(IsNull(Rec!chvEngSecType), "", Rec!chvEngSecType)
        vchCategory = IIf(IsNull(Rec!vchProjectCategory), "", Rec!vchProjectCategory)
        intCategoryID = IIf(IsNull(Rec!intProjCatID), "", Rec!intProjCatID)
    End If
    Rec.Close
    Rec.Open "Select Sum(fltEstAmt) As ProjCost From suEstimation Where intYearID=" & gbFinancialYearID & "  And decProjectID = " & arrInput(0), mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        fltProjectCost = IIf(IsNull(Rec!ProjCost), 0, Rec!ProjCost)
    End If
    Rec.Close
    Rec.Open "Select tnyStatus From suProjectDetails Where decProjectID = " & arrInput(0), mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        tnyStatus = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
    End If
    Rec.Close
End Sub

Public Function FindSourceOfFund(mSourceOfFundID As Variant) As String
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    
    objDB.SetConnection mCnn
    mSQL = "Select * From suSourceOfFund Where intSourceFundID =  " & mSourceOfFundID
    Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not (Rec.BOF And Rec.EOF) Then
        intSourceOfFundID = Rec!intSourceFundID
        vchSourceOfFund = Rec!vchSourceFundName
    End If
    Rec.Close
    FindSourceOfFund = vchSourceOfFund
End Function



Public Function GetFundDetails(mYearID As Integer, mProjID As Double) As Collection
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    Dim mLoop As Integer
    Dim objProjFund As New clsProjectFund
    
    If IsMissing(mYearID) Then
        mYearID = intYearID
        If mYearID < 2000 Then
            mYearID = gbFinancialYearID
        End If
    End If
    
    If IsMissing(mProjID) Then
        mProjID = decProjectID
        If Not IsNumeric(mProjID) Then
            mProjID = 0
        End If
    End If
    
     
    
    mSQL = "SELECT  FundDetails.intFundSrcID, "
    mSQL = mSQL + "    M_FundSource.chvFundSource, "
    mSQL = mSQL + "    M_FundSource.chvCode, "
    mSQL = mSQL + "    FundDetails.intYearID, "
    mSQL = mSQL + "    FundDetails.intSlNo, "
    mSQL = mSQL + "    FundDetails.fltAmt "
    mSQL = mSQL + " From FundDetails "
    mSQL = mSQL + " INNER JOIN M_FundSource ON FundDetails.intFundSrcID = M_FundSource.intFundSrcID"
    mSQL = mSQL + " Where FundDetails.intYearID = " & mYearID & " And FundDetails.decProjectID = " & mProjID
    mSQL = mSQL + "  AND ISNULL(tnyDelete,0)<>1"

    If objDB.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
    Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not (Rec.BOF And Rec.EOF) Then
    mLoop = 0
    While Not Rec.EOF
        objProjFund.SourceOfFundID = Rec!intFundSrcID
        objProjFund.SourceCode = Rec!chvCode
        objProjFund.SourceName = Rec!chvFundSource
        objProjFund.AllocatedYearID = Rec!intYearID
        objProjFund.SlNo = Rec!intSlNo
        objProjFund.SourceWiseAmount = Rec!fltAmt
        objProjFund.SourceWiseUtilisedAmount = 0
        
        'On Error GoTo SkipSulekha:
        'Call Add(objProjFund, str(Rec!intSlNo))
        mSourceAsCollection.Add objProjFund, str(Rec!intSlNo)
        Set objProjFund = Nothing
        Rec.MoveNext
    Wend
    End If
    Rec.Close
    End If
    Set GetFundDetails = mSourceAsCollection 'mCol
   Exit Function

SkipSulekha:
    frmRequisition.FundErSulekha = 1
    'MsgBox "Some Mistakes in Ported data from Sulekha"
End Function

' Option Explicit
' local variable to hold collection
' Private mCol As Collection

Private Sub Add(Item As Variant, Optional sKey As String, Optional Before As Variant, Optional After As Variant)

    'create a new object
    Dim objNewMember() As Variant

    'set the properties passed into the method
    ReDim objNewMember(0 To 1)
    If IsObject(Item) Then
        Set objNewMember(0) = Item
    Else
        objNewMember(0) = Item
    End If


    objNewMember(1) = sKey
    If Len(sKey) = 0 Then
        mCol.Add objNewMember, , Before, After
    Else
        mCol.Add objNewMember, sKey, Before, After
    End If

End Sub

Public Property Get Item(vntIndexKey As Variant) As Variant
    Dim vItem() As Variant
    vItem = mCol(vntIndexKey)

    If IsObject(vItem(0)) Then
        Set Item = vItem(0)
    Else
        Item = vItem(0)
    End If
End Property

Public Property Get ItemKEY(vntIndexKey As Variant) As String
    ItemKEY = mCol(vntIndexKey)(1)
End Property

Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    count = mCol.count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

'Private Sub Class_Initialize()
'    'creates the collection when this class is created
'    Set mCol = New Collection
'End Sub
'
'Private Sub Class_Terminate()
'    'destroys collection when this class is terminated
'    Set mCol = Nothing
'End Sub
'
