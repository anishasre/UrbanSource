VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
            
    Private mvarAccountHeadID   As Long
    Private mvarAccountHeadCode As String
    Private mvarBankName        As String
    Private mvarBranch          As String
    Private mvarBankCode        As String
    Private mvarBranchCode      As String
    Private mvarAccountNumber   As String
    Private mvarAddress1        As String
    Private mvarAddress2        As String
    Private mvarPhone           As String
    Private mvarFax             As String
    Private mvarEmail           As String
    Private mvarOpening         As Double
    Private mvarBankID          As Long
    Private mvarFundID          As Variant
    Private mvarCrDrFlag        As Boolean
    Private mvarMinorAccountHead As String
    Private mvarMinorAccountHeadID As Integer
    Private mvarReconStartDate As Variant
    Private mvarReconLastDate  As Variant
    
    Public Property Let CrDrFlag(ByVal vData As Boolean)
        mvarCrDrFlag = vData
    End Property
    
    Public Property Get CrDrFlag() As Boolean
        CrDrFlag = mvarCrDrFlag
    End Property
    
    Public Property Let Opening(ByVal vData As Double)
        mvarOpening = vData
    End Property
    
    Public Property Get Opening() As Double
        Opening = mvarOpening
    End Property
     Public Property Let FundID(ByVal vData As Variant)
        mvarFundID = vData
    End Property
    
    Public Property Get FundID() As Variant
       FundID = mvarFundID
    End Property
    
    Public Property Let Email(ByVal vData As String)
        mvarEmail = vData
    End Property
    
    Public Property Get Email() As String
        Email = mvarEmail
    End Property
    
    Public Property Let Fax(ByVal vData As String)
        mvarFax = vData
    End Property
    
    Public Property Get Fax() As String
        Fax = mvarFax
    End Property
    
    Public Property Let Phone(ByVal vData As String)
        mvarPhone = vData
    End Property
    
    Public Property Get Phone() As String
        Phone = mvarPhone
    End Property
    
    Public Property Let Address2(ByVal vData As String)
        mvarAddress2 = vData
    End Property
    
    Public Property Get Address2() As String
        Address2 = mvarAddress2
    End Property
    
    Public Property Let Address1(ByVal vData As String)
        mvarAddress1 = vData
    End Property
    
    Public Property Get Address1() As String
        Address1 = mvarAddress1
    End Property
    
    Public Property Let AccountNumber(ByVal vData As String)
        mvarAccountNumber = vData
    End Property
    
    Public Property Get AccountNumber() As String
        AccountNumber = mvarAccountNumber
    End Property
    
    Public Property Let Branch(ByVal vData As String)
        mvarBranch = vData
    End Property
    
    Public Property Get Branch() As String
        Branch = mvarBranch
    End Property
    
    Public Property Let BankCode(ByVal vData As String)
        mvarBankCode = vData
    End Property
    
    Public Property Get BankCode() As String
        BankCode = mvarBankCode
    End Property
    
    Public Property Let BranchCode(ByVal vData As String)
        mvarBranchCode = vData
    End Property
    
    Public Property Get BranchCode() As String
        BranchCode = mvarBranchCode
    End Property
    Public Property Let BankName(ByVal vData As String)
        mvarBankName = vData
    End Property
    
    Public Property Get BankName() As String
        BankName = mvarBankName
    End Property
    
    Public Property Get BankAccountHeadID() As Long
        BankAccountHeadID = mvarAccountHeadID
    End Property
    
    Public Property Let BankAccountHeadID(ByVal vData As Long)
        mvarAccountHeadID = vData
    End Property
    
    Public Property Get BankID() As Long
        BankID = mvarBankID
    End Property
    
    Public Property Let BankID(ByVal vData As Long)
        mvarBankID = vData
    End Property
    
    Public Property Get BankAccountHeadCode() As String
        BankAccountHeadCode = mvarAccountHeadCode
    End Property
    
    Public Property Let BankAccountHeadCode(ByVal vData As String)
        mvarAccountHeadCode = vData
    End Property
    
    Public Property Let MinorAccountHeadID(ByVal vData As Integer)
        mvarMinorAccountHeadID = vData
    End Property
    
    Public Property Get MinorAccountHeadID() As Integer
        MinorAccountHeadID = mvarMinorAccountHeadID
    End Property
    
    Public Property Let MinorAccountHead(ByVal vData As String)
        mvarMinorAccountHead = vData
    End Property
    
    Public Property Get MinorAccountHead() As String
        MinorAccountHead = mvarMinorAccountHead
    End Property
    
    Public Property Get ReconciliationStartDate() As String
        If IsDate(mvarReconStartDate) Then
            ReconciliationStartDate = DdMmmYy(CDate(mvarReconStartDate))
        Else
            ReconciliationStartDate = ""
        End If
    End Property
    
    Public Property Get ReconciliationLastDate() As String
        If IsDate(mvarReconLastDate) Then
            ReconciliationLastDate = DdMmmYy(CDate(mvarReconLastDate))
        Else
            ReconciliationLastDate = ""
        End If
    End Property
    
    Private Sub ClearVariables()
            mvarBankID = -1
            mvarAccountHeadID = -1
            mvarAccountHeadCode = ""
            mvarBankName = ""
            mvarBranch = ""
            mvarBankCode = ""
            mvarBranchCode = ""
            mvarAccountNumber = ""
            mvarAddress1 = ""
            mvarAddress2 = ""
            mvarPhone = ""
            mvarFax = ""
            mvarEmail = ""
            mvarOpening = "0.00"
            mvarFundID = Null
            mvarMinorAccountHead = ""
            mvarMinorAccountHeadID = -1
            mvarReconStartDate = Null
            mvarReconLastDate = Null
    End Sub
    
    Public Sub SetBankInfo(mBankID As Double)
            
            Dim mCn As ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            
            Call ClearVariables
            
            mSQL = "SELECT * "
            mSQL = mSQL & " FROM faBanks "
            mSQL = mSQL & " LEFT JOIN faAccountHeads ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID "
            mSQL = mSQL & " LEFT JOIN faMinorAccountHeads ON faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID "
            mSQL = mSQL & " LEFT JOIN faFunds ON faBanks.intFundID = faFunds.intFundID "
            mSQL = mSQL & " WHERE faBanks.intBankID = " & mBankID
            
            
            Rec.CursorLocation = adUseClient
            Set Rec = GetRecordSet(mSQL)
            If Rec.RecordCount > 0 Then
                '-----------------------------------------------'
                ' To over ride Null Value Errors if any         '
                '-----------------------------------------------'
                  On Error Resume Next
                '-----------------------------------------------'
                
                mvarAccountHeadID = Rec!intAccountHeadID
                mvarAccountHeadCode = Rec!vchAccountHeadCode
                mvarBankID = Rec!intBankID
                mvarBankName = Rec!vchBankName
                mvarBranch = Rec!vchBranch
                mvarBankCode = Rec!vchBankCode
                mvarBranchCode = Rec!vchBranchCode
                mvarAccountNumber = Rec!vchAccountNumber
                mvarAddress1 = Rec!vchAddress1
                mvarAddress2 = Rec!vchAddress2
                mvarPhone = Rec!vchPhone
                mvarFax = Rec!vchFax
                mvarEmail = Rec!vchEmail
                mvarOpening = Rec!fltOpening
                mvarFundID = Rec!intFundID
                mvarCrDrFlag = Rec!tinDebitOrCreditFlag
                mvarMinorAccountHeadID = Rec!intMinorAccountHeadID
                mvarMinorAccountHead = Rec!vchMinorAccountHead
                
                If IsDate(Rec!dtReconStartDate) Then
                    mvarReconStartDate = Rec!dtReconStartDate
                Else
                    mvarReconStartDate = Null
                End If
                
                If IsDate(Rec!dtReconEndDate) Then
                    mvarReconLastDate = Rec!dtReconEndDate
                Else
                    mvarReconLastDate = Null
                End If
                
                On Error GoTo 0
            End If
            Rec.Close
    End Sub
    
    Public Sub SetBankInfoByAccID(mAccID As Long)
            
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            
            Call ClearVariables
            mSQL = "SELECT * "
            mSQL = mSQL & " FROM faBanks "
            mSQL = mSQL & " LEFT JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faBanks.intAccountHeadID "
            mSQL = mSQL & " LEFT JOIN faMinorAccountHeads ON faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID "
            mSQL = mSQL & " LEFT JOIN faFunds ON faFunds.intFundID = faBanks.intFundID "
            mSQL = mSQL & " WHERE faBanks.intAccountHeadID = " & mAccID
            
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                '-----------------------------------------------'
                ' To over ride Null Value Errors if any         '
                '-----------------------------------------------'
                  On Error Resume Next
                '-----------------------------------------------'
                mvarAccountHeadID = Rec!intAccountHeadID
                mvarAccountHeadCode = Rec!vchAccountHeadCode
                mvarBankID = Rec!intBankID
                mvarBankName = Rec!vchBankName
                mvarBranch = Rec!vchBranch
                mvarBankCode = Rec!vchBankCode
                mvarBranchCode = Rec!vchBranchCode
                mvarAccountNumber = Rec!vchAccountNumber
                mvarAddress1 = Rec!vchAddress1
                mvarAddress2 = Rec!vchAddress2
                mvarPhone = Rec!vchPhone
                mvarFax = Rec!vchFax
                mvarEmail = Rec!vchEmail
                mvarOpening = Rec!fltOpening
                mvarFundID = Rec!intFundID
                mvarCrDrFlag = Rec!tinDebitOrCreditFlag
                
                mvarMinorAccountHeadID = Rec!intMinorAccountHeadID
                mvarMinorAccountHead = Rec!vchMinorAccountHead
                
                If IsDate(Rec!dtReconStartDate) Then
                    mvarReconStartDate = Rec!dtReconStartDate
                Else
                    mvarReconStartDate = Null
                End If
                
                If IsDate(Rec!dtReconEndDate) Then
                    mvarReconLastDate = Rec!dtReconEndDate
                Else
                    mvarReconLastDate = Null
                End If
                
                On Error GoTo 0
            End If
            Rec.Close
            
    End Sub
   
    Public Function GetNeWChequeNumber() As Variant
            Dim mCn As ADODB.Connection
            Dim Rec As New ADODB.Recordset
            
            Set Rec = GetRecordSet("spGetNewChequeNo " & mvarAccountHeadID)
            
            If Not (Rec.BOF And Rec.EOF) Then
                GetNeWChequeNumber = IIf(IsNull(Rec.Fields(0).value), "", Rec.Fields(0).value)
            End If
            Rec.Close
    End Function
    Public Function GetReconciliationStatus(intBankID As Integer, mDate As Date) As Boolean
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mMonthID As Integer
        Dim mFinYear As Integer
       
        mMonthID = Month(mDate)
        mFinYear = Year(mDate)
        If mMonthID < 4 Then
            mFinYear = mFinYear - 1
        End If
        
        mSQL = "Select * From faBankReconcile "
        mSQL = mSQL + " Where tnyReconStatus=1 And intBankAccountHeadID = " & intBankID
        mSQL = mSQL + " AND intMonthID=" & mMonthID
        mSQL = mSQL + " AND intYearID=" & mFinYear
               
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSQL)
            If Rec.RecordCount > 0 Then
                GetReconciliationStatus = True
            Else
                GetReconciliationStatus = False
            End If
    End Function
    
     Public Function GetLastReconDate(intBankID As Integer) As Date
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mMonthID As Integer
        Dim mFinYear As Integer
               
        mSQL = "Select * From faBanks Where intAccountHeadID=" & intBankID
        
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                GetLastReconDate = IIf(IsNull(Rec!dtReconEndDate), gbStartingDate, Rec!dtReconEndDate)
            End If
        Rec.Close
    End Function
