VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
    Private intSubsidiaryAccountHeadID  As Variant
    Private intSubLedgerTypeID          As Variant
    Private vchSubLedgerCode            As Variant
    Private vchTitle                    As Variant
    Private vchSubTitle                 As Variant
    Private intAccountHeadID            As Variant
    Private vchAccountHeadCode          As Variant
    Private intMinorAccountHeadID       As Variant
    
    Private vchMinorAccountHeadCode     As Variant
    Private fltOpeningBalance           As Variant
    Private fltClosingBalance           As Variant
    Private numEmpID                    As Variant
    Private vchReferenceCode            As Variant
    Private vchReferenceCode2           As Variant
    Private vchNameOfSubLedger          As Variant
    Private vchHouseOrOffice            As Variant
    Private vchStreet                   As Variant
    Private vchLocalPlace               As Variant
    Private vchMainPlace                As Variant
    Private vchPostOffice               As Variant
    Private vchPinCode                  As Variant
    Private vchPhone                    As Variant
    Private numWardNo                   As Variant
    Private intDoorNo                   As Variant
    Private vchDoorNo2                  As Variant
    Private intLBID                     As Variant
    Private intGroupID                  As Variant
    Private tnyDeleted                  As Variant
    Private dblDDOCode                  As Variant
    Private vchDesignation              As Variant
    Private vchDepartment               As Variant
    
    Private vchSubLedgerType            As Variant
    
    
    
'    Private intSubsidiaryAccountHeadID  As Double
'    Private intSubLedgerTypeID          As Integer
'    Private vchSubLedgerCode            As String
'    Private vchTitle                    As String
'    Private vchSubTitle                 As String
'    Private intAccountHeadID            As Long
'    Private vchAccountHeadCode          As String
'    Private intMinorAccountHeadID       As Long
'
'    Private vchMinorAccountHeadCode     As String
'    Private fltOpeningBalance           As Double
'    Private fltClosingBalance           As Double
'    Private numEmpID                    As Double
'    Private vchReferenceCode            As String
'    Private vchReferenceCode2           As String
'    Private vchName                     As String
'    Private vchHouseOrOffice            As String
'    Private vchStreet                   As String
'    Private vchLocalPlace               As String
'    Private vchMainPlace                As String
'    Private vchPostOffice               As String
'    Private vchPinCode                  As String
'    Private vchPhone                    As String
'    Private numWardNo                   As Double
'    Private intDoorNo                   As Long
'    Private vchDoorNo2                  As String
'    Private intLBID                     As Long
'    Private intGroupID                  As Long
'    Private tnyDeleted                  As Integer
'
    Public Property Let SubsidiaryAccountHeadID(ByVal mData As Variant)
        intSubsidiaryAccountHeadID = mData
    End Property
    Public Property Get SubsidiaryAccountHeadID() As Variant
        SubsidiaryAccountHeadID = intSubsidiaryAccountHeadID
    End Property
    
    Public Property Let SubLedgerTypeID(ByVal mData As Variant)
        intSubLedgerTypeID = mData
    End Property
    Public Property Get SubLedgerTypeID() As Variant
        SubLedgerTypeID = intSubLedgerTypeID
    End Property
    
    Public Property Let SubLedgerCode(ByVal mData As Variant)
        vchSubLedgerCode = mData
    End Property
    Public Property Get SubLedgerCode() As Variant
        SubLedgerCode = vchSubLedgerCode
    End Property
    
    Public Property Let Title(ByVal mData As Variant)
        vchTitle = mData
    End Property
    Public Property Get Title() As Variant
        Title = vchTitle
    End Property
    
    Public Property Let SubTitle(ByVal mData As Variant)
        vchSubTitle = mData
    End Property
    Public Property Get SubTitle() As Variant
        SubTitle = vchSubTitle
    End Property
    
    Public Property Let AccountHeadID(ByVal mData As Variant)
        intAccountHeadID = mData
    End Property
    Public Property Get AccountHeadID() As Variant
        AccountHeadID = intAccountHeadID
    End Property
    
    Public Property Let AccountHeadCode(ByVal mData As Variant)
        vchAccountHeadCode = mData
    End Property
    Public Property Get AccountHeadCode() As Variant
        AccountHeadCode = vchAccountHeadCode
    End Property
    
    Public Property Let MinorAccountHeadID(ByVal mData As Variant)
        intMinorAccountHeadID = mData
    End Property
    Public Property Get MinorAccountHeadID() As Variant
        MinorAccountHeadID = intMinorAccountHeadID
    End Property
    
    Public Property Let MinorAccountHeadCode(ByVal mData As Variant)
        vchMinorAccountHeadCode = mData
    End Property
    Public Property Get MinorAccountHeadCode() As Variant
        MinorAccountHeadCode = vchMinorAccountHeadCode
    End Property
    
    Public Property Let OpeningBalance(ByVal mData As Variant)
        fltOpeningBalance = mData
    End Property
    Public Property Get OpeningBalance() As Variant
        OpeningBalance = fltOpeningBalance
    End Property
    
    Public Property Let ClosingBalance(ByVal mData As Variant)
        fltClosingBalance = mData
    End Property
    Public Property Get ClosingBalance() As Variant
        ClosingBalance = fltClosingBalance
    End Property
    
    Public Property Let EmpID(ByVal mData As Variant)
        numEmpID = mData
    End Property
    Public Property Get EmpID() As Variant
        EmpID = numEmpID
    End Property
    
    Public Property Let ReferenceCode(ByVal mData As Variant)
        vchReferenceCode = mData
    End Property
    Public Property Get ReferenceCode() As Variant
        ReferenceCode = vchReferenceCode
    End Property
    
    Public Property Let ReferenceCode2(ByVal mData As Variant)
        vchReferenceCode2 = mData
    End Property
    Public Property Get ReferenceCode2() As Variant
        ReferenceCode2 = vchReferenceCode2
    End Property
    
    Public Property Let NameOfSubLedger(ByVal mData As Variant)
        vchNameOfSubLedger = mData
    End Property
    Public Property Get NameOfSubLedger() As Variant
        NameOfSubLedger = vchNameOfSubLedger
    End Property
    
    Public Property Let HouseOrOffice(ByVal mData As Variant)
        vchHouseOrOffice = mData
    End Property
    Public Property Get HouseOrOffice() As Variant
        HouseOrOffice = vchHouseOrOffice
    End Property
    
    Public Property Let Street(ByVal mData As Variant)
        vchStreet = mData
    End Property
    Public Property Get Street() As Variant
        Street = vchStreet
    End Property
    
    Public Property Let LocalPlace(ByVal mData As Variant)
        vchLocalPlace = mData
    End Property
    Public Property Get LocalPlace() As Variant
        LocalPlace = vchLocalPlace
    End Property
    
    Public Property Let MainPlace(ByVal mData As Variant)
        vchMainPlace = mData
    End Property
    Public Property Get MainPlace() As Variant
        MainPlace = vchMainPlace
    End Property
    
    Public Property Let PostOffice(ByVal mData As Variant)
        vchPostOffice = mData
    End Property
    Public Property Get PostOffice() As Variant
        PostOffice = vchPostOffice
    End Property
    
    Public Property Let PinCode(ByVal mData As Variant)
        vchPinCode = mData
    End Property
    Public Property Get PinCode() As Variant
        PinCode = vchPinCode
    End Property
    
    Public Property Let Phone(ByVal mData As Variant)
        vchPhone = mData
    End Property
    Public Property Get Phone() As Variant
        Phone = vchPhone
    End Property
    
    Public Property Let WardNo(ByVal mData As Variant)
        numWardNo = mData
    End Property
    Public Property Get WardNo() As Variant
        WardNo = numWardNo
    End Property
    
    Public Property Let DoorNo(ByVal mData As Variant)
        DoorNo = intDoorNo
    End Property
    Public Property Get DoorNo() As Variant
        DoorNo = intDoorNo
    End Property
    
    Public Property Let DoorNo2(ByVal mData As Variant)
        vchDoorNo2 = mData
    End Property
    Public Property Get DoorNo2() As Variant
        DoorNo2 = vchDoorNo2
    End Property
    
    Public Property Let LBID(ByVal mData As Variant)
        intLBID = mData
    End Property
    Public Property Get LBID() As Variant
        LBID = intLBID
    End Property
    
    Public Property Let GroupID(ByVal mData As Variant)
        intGroupID = mData
    End Property
    Public Property Get GroupID() As Variant
        GroupID = intGroupID
    End Property
    
    Public Property Let Deleted(ByVal mData As Variant)
        tnyDeleted = mData
    End Property
    
    Public Property Get Deleted() As Variant
        Deleted = tnyDeleted
    End Property
    
    Public Property Let SubLedgerType(ByVal mData As Variant)
        vchSubLedgerType = mData
    End Property
    
    Public Property Get SubLedgerType() As Variant
        SubLedgerType = vchSubLedgerType
    End Property
    
    Public Property Let DDOCode(ByVal mData As Variant)
        dblDDOCode = mData
    End Property
    
    Public Property Get DDOCode() As Variant
        DDOCode = dblDDOCode
    End Property
    
    Public Property Let Designation(ByVal mData As Variant)
        vchDesignation = mData
    End Property
    
    Public Property Get Designation() As Variant
        Designation = vchDesignation
    End Property
    
    Public Property Let Department(mData As Variant)
        vchDepartment = mData
    End Property
    
    Public Property Get Department() As Variant
        Department = vchDepartment
    End Property
    
    
    Private Function SetSubLedgerDetails_()
            On Error GoTo Err:
            Dim objDB As New clsDB
            Dim mSQL As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            
            If objDB.SetConnection(mCnn) Then
                If IsNumeric(SubsidiaryAccountHeadID) Then
                    mSQL = "Select * from faSubSidiaryAccountHeads Where intSubsidiaryAccountHeadID = " & SubsidiaryAccountHeadID
                ElseIf SubLedgerCode <> "" Then
                    mSQL = "Select * from faSubSidiaryAccountHeads Where vchSubLedgerCode = '" & SubLedgerCode & "'"
                ElseIf NameOfSubLedger <> "" Then
                    mSQL = "Select * from faSubSidiaryAccountHeads Where vchName = '" & NameOfSubLedger & "'"
                End If
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    SubsidiaryAccountHeadID = Rec!intSubsidiaryAccountHeadID
                    SubLedgerTypeID = Rec!intSubLedgerTypeID
                    SubLedgerCode = Rec!vchSubLedgerCode
                    Title = Rec!vchTitle
                    SubTitle = Rec!vchSubTitle
                    AccountHeadID = Rec!intAccountHeadID
                    AccountHeadCode = Rec!vchAccountHeadCode
                    MinorAccountHeadID = Rec!intMinorAccountHeadID
                    MinorAccountHeadCode = Rec!vchMinorAccountHeadCode
                    OpeningBalance = Rec!fltOpeningBalance
                    ClosingBalance = Rec!fltClosingBalance
                    EmpID = Rec!numEmpID
                    ReferenceCode = Rec!vchReferenceCode
                    ReferenceCode2 = Rec!vchReferenceCode2
                    NameOfSubLedger = Rec!vchName
                    HouseOrOffice = Rec!vchHouseOrOffice
                    Street = Rec!vchStreet
                    LocalPlace = Rec!vchLocalPlace
                    MainPlace = Rec!vchMainPlace
                    PostOffice = Rec!vchPostOffice
                    PinCode = Rec!vchPinCode
                    Phone = Rec!vchPhone
                    WardNo = Rec!numWardNo
                    DoorNo = Rec!intDoorNo
                    DoorNo2 = Rec!vchDoorNo2
                    LBID = Rec!intLBID
                    GroupID = Rec!intGroupID
                    Deleted = Rec!tnyDeleted
                    DDOCode = Rec!vchReferenceCode
                    Designation = Rec!vchDesignation
                    Department = Rec!vchDepartment
                End If
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    
    End Function
    
    Public Sub SetSubLedgerDetails(SubLedgerID As Double)
        On Error GoTo Err:
            Dim objDB As New clsDB
            Dim mSQL As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Call Class_Initialize
            If objDB.SetConnection(mCnn) Then
                mSQL = "Select * From faSubSidiaryAccountHeads "
                mSQL = mSQL + " Inner Join faSubLedgerTypes On faSubSidiaryAccountHeads.intSubLedgerTypeID = faSubLedgerTypes.intSubLedgerTypeID "
                mSQL = mSQL + " Where intSubsidiaryAccountHeadID = " & SubLedgerID
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    intSubsidiaryAccountHeadID = Rec!intSubsidiaryAccountHeadID
                    intSubLedgerTypeID = Rec!intSubLedgerTypeID
                    vchSubLedgerCode = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
                    vchTitle = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
                    vchSubTitle = IIf(IsNull(Rec!vchSubTitle), "", Rec!vchSubTitle)
                    intAccountHeadID = IIf(IsNull(Rec!intAccountHeadID), -1, Rec!intAccountHeadID)
                    vchAccountHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    intMinorAccountHeadID = IIf(IsNull(Rec!intMinorAccountHeadID), -1, Rec!intMinorAccountHeadID)
                    vchMinorAccountHeadCode = IIf(IsNull(Rec!vchMinorAccountHeadCode), "", Rec!vchMinorAccountHeadCode)
                    fltOpeningBalance = IIf(IsNull(Rec!fltOpeningBalance), 0, Rec!fltOpeningBalance)
                    fltClosingBalance = IIf(IsNull(Rec!fltClosingBalance), 0, Rec!fltClosingBalance)
                    numEmpID = IIf(IsNull(Rec!numEmpID), -1, Rec!numEmpID)
                    vchReferenceCode = IIf(IsNull(Rec!vchReferenceCode), "", Rec!vchReferenceCode)
                    vchReferenceCode2 = IIf(IsNull(Rec!vchReferenceCode2), "", Rec!vchReferenceCode2)
                    vchNameOfSubLedger = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    If vchNameOfSubLedger = "" Then
                        vchNameOfSubLedger = vchTitle
                    End If
                    vchHouseOrOffice = IIf(IsNull(Rec!vchHouseOrOffice), "", Rec!vchHouseOrOffice)
                    vchStreet = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                    vchLocalPlace = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                    vchMainPlace = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                    vchPostOffice = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
                    vchPinCode = IIf(IsNull(Rec!vchPinCode), "", Rec!vchPinCode)
                    vchPhone = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                    numWardNo = IIf(IsNull(Rec!numWardNo), -1, Rec!numWardNo)
                    intDoorNo = IIf(IsNull(Rec!intDoorNo), -1, Rec!intDoorNo)
                    vchDoorNo2 = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                    intLBID = IIf(IsNull(Rec!intLBID), -1, Rec!intLBID)
                    intGroupID = IIf(IsNull(Rec!intGroupID), -1, Rec!intGroupID)
                    tnyDeleted = IIf(IsNull(Rec!tnyDeleted), -1, Rec!tnyDeleted)
                    vchSubLedgerType = IIf(IsNull(Rec!vchSubLedgerType), "", Rec!vchSubLedgerType)
                    DDOCode = IIf(IsNull(Rec!vchReferenceCode), "", Rec!vchReferenceCode)
                    Designation = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
                    Department = IIf(IsNull(Rec!vchDepartment), "", Rec!vchDepartment)
                End If
            End If
        Exit Sub
Err:
        Debug.Print (Error$)
    
    End Sub

Private Sub Class_Initialize()
    intSubsidiaryAccountHeadID = Null
    intSubLedgerTypeID = Null
    vchSubLedgerCode = Null
    vchTitle = Null
    vchSubTitle = Null
    intAccountHeadID = Null
    vchAccountHeadCode = Null
    intMinorAccountHeadID = Null
    
    vchMinorAccountHeadCode = Null
    fltOpeningBalance = Null
    fltClosingBalance = Null
    numEmpID = Null
    vchReferenceCode = Null
    vchReferenceCode2 = Null
    vchNameOfSubLedger = Null
    vchHouseOrOffice = Null
    vchStreet = Null
    vchLocalPlace = Null
    vchMainPlace = Null
    vchPostOffice = Null
    vchPinCode = Null
    vchPhone = Null
    numWardNo = Null
    intDoorNo = Null
    vchDoorNo2 = Null
    intLBID = Null
    intGroupID = Null
    tnyDeleted = Null
    vchSubLedgerType = Null
    
    dblDDOCode = Null
    vchDesignation = Null
    vchDepartment = Null
End Sub
