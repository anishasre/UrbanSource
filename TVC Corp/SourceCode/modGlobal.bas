Attribute VB_Name = "modGlobal"
    
    Option Explicit
    
    Public gbTopMargin As Integer           'Added by Aiby
    Public gbBottomMargin As Integer        'Added by Aiby
    Public gbLeftMargin As Integer          'Added by Aiby
    Public gbRightMargin As Integer         'Added by Aiby
    Public gbPageWidth As Long              'Added by Aiby
    
    Public gbNoOfLinesPerPage As Integer    'Total no of line/page  'Added by Aiby
    Public gbNoOfPrintableLines As Integer  'Total no of line perpage - margins  'Added by Aiby
    Public gbTextArea As Integer            'Printable No of characters in a page after margins (Added by Aiby)
    Public gbFileNO As Integer              'free file handler
    
    Public gbPrinterPort As String          'gets the priter port to print
    Public gbFileName As String             'stors the text filename
    
    Public gbReportID   As Integer          '1=Day Book:  2=Cash Book
    Public gbSearchProductCode As String 'Searching Code
    Private Declare Function SetParent Lib "user32" (ByVal frmChild As Long, ByVal frmParent As Long) As Long
    
    Private Const LOCALE_SLONGDATE As Long = &H20  'long date format string
    Private Const LOCALE_SSHORTDATE = &H1F  'short date format string
    Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
                    ByVal lpLCData As String, ByVal cchData As Long) As Long
    '-------------------------------------------------'
    ' Global Environment Variable for Application     '
    '-------------------------------------------------'
        Public gbCnn As New ADODB.Connection
        Public gbSelectedAccount As String
        Public gbSelectedListItem As String
        Public gbSelectedListIndex As Long
        Public gbAppKey As String
    '-------------------------------------------------'
    
'      '------------Ptax-------------------------------'
'        Public mRowCount As Integer
'        Public TaxPenal As Double
'        Public PenalPTax As Double
'     '-------------------------------------------------'
        Public gbDate As Date
        Public gbServerdate As Date
        Public gbServerPath As String
        Public gbSessionID As Long
    '-------------------------------------------------'
    '   Primary Ledger Groups ID
    '-------------------------------------------------'
    
        Public Const faCash = 1
        Public Const faBank = 2
        
        Public Const faPurchase = 3
        Public Const faPurchaseReturn = 4
        
        Public Const faSales = 5
        Public Const faSalesReturn = 6
        
        Public Const faSundryDebtors = 7
        Public Const faSundryCreditors = 8
        
        Public Const faDirectExpences = 9
        Public Const faIndirectExpences = 10
        
        Public Const faIncome = 11
        Public Const faIndirectIncome = 12
        
        Public Const faCurrentAsset = 13
        Public Const faFixedAsset = 14
        
        Public Const faCurrentLiability = 15
        Public Const faLiability = 16
        
        Public Const faCapital = 17
        Public Const faDrawing = 18
        Public Const faStock = 19
        Public Const faPandL = 20
      
        
        Public gbComputerName As String
    
    '*************************************************'
    Public Function SetConnection(mvarConnection As ADODB.Connection) As Boolean
        On Error GoTo ErrOpen
        'mvarConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
        'mvarConnection.ConnectionString = "C:\Documents and Settings\Administrator\My Documents\My Projects\iJewellary\iJewellary.mdb"
        'mvarConnection.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsnERP"
        'mvarConnection.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsnJewelery"
        mvarConnection.ConnectionString = "DSN=dsnFA"
        mvarConnection.Open
        SetConnection = True
        Exit Function
ErrOpen:
        SetConnection = False
    End Function
    
    Public Function SetEnvironment() As Boolean
        
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objAcc As New clsAccounts
        Dim LCID As Long       'get the locale for the user
        Dim GetLongDateFormat As String
        Dim GetShortDateFormat As String
        
        If gbLocalBodyID = 222 Then
            gbPDEMode = True
        Else
            gbPDEMode = False
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        'NOTE: CHANGE BANK PERMITED UPTO A DATE FOR
        '       RECEIPTS FROM OTHER LSGIs
                gbBankChangePermitDate = "01-Apr-2014"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        gbSaankhyaINI = App.Path & "\Saankhya.INI"
        mSql = "SELECT *, GetDate() as SysDate from faFinancialYear WHERE tinCurrentFinancialYearFlag =1"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
'            If mPCDate = Rec!SysDate Then
'            End If
            gbTransactionDate = Rec!SysDate
            
             gbServerdate = Rec!SysDate  ' Server date
            
            gbTransactionDate = DateSerial(Year(gbTransactionDate), Month(gbTransactionDate), Day(gbTransactionDate))
            gbStartingDate = Rec!dtStartingDate
            gbEndingDate = Rec!dtEndingDate
            If Not (gbTransactionDate <= gbEndingDate And gbTransactionDate >= gbStartingDate) Then
                MsgBox "Error: Server Transaction Date!", vbInformation
                GoTo ErrorCheck:
            End If
            gbFinancialYearID = Rec!intFinancialYearID
            gbDate = Date
            
            gbCurrentPeriodID = Month(gbTransactionDate)
            If gbCurrentPeriodID > 3 And gbCurrentPeriodID < 10 Then
                gbCurrentPeriodID = 1
            Else
                gbCurrentPeriodID = 2
            End If
        Else
            MsgBox "Error: FinancialYear Not Set!", vbInformation
            End
        End If
        Rec.Close
        '''''' To check system long date format and short date format
        LCID = GetSystemDefaultLCID()
        If LCID <> 0 Then
          'return the long date format
           GetLongDateFormat = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
           'return the short date format
'           GetShortDateFormat = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)

           Select Case GetLongDateFormat
            Case "dd MM yyyy"
            Case "dd MMM yyyy"
            Case "dd /MM /yyyy"
            Case "dd/MM/yyyy"
            Case "dd/MMM/yyyy"
            Case "dd /MMM/yyyy"
            Case "dd-MM-yyyy"
            Case "dd-MMM-yyyy"
            Case "dd MMMM yyyy"

            Case Else
                MsgBox "Please set System date format as dd MMM yyyy OR dd MM yyyy "
                End
           End Select
'           If GetLongDateFormat = "dd MM yyyy" Or GetLongDateFormat = "dd /MM /yyyy" Or GetLongDateFormat = "dd -MM -yyyy" Then
'           Else
'                MsgBox "Please set System date format as dd MM yyyy"
'                End
'           End If
'           If GetShortDateFormat = "dd MM yyyy" Or GetShortDateFormat = "dd MM yyyy" Or GetShortDateFormat = "dd/MM/yyyy" Or GetShortDateFormat = "dd -MM -yyyy" Then
'           Else
'                MsgBox "Please set System date format as dd MM yyyy"
'                End
'           End If
         End If
        ''''''''''''''''''
        
        mSql = "Select * From faLBSettings"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            gbLocalBodyID = Rec!intLBID
            gbLBID = Rec!intLBID
            gbDistID = Rec!tnyDistrictID
            gbLBName = Rec!chvLBNameEnglish
            gbLBType = Rec!tnyLBTypeID
            gbLBTitle = IIf(IsNull(Rec!chvTitle), gbLBName, Rec!chvTitle)
            gbLocationID = Rec!intLocationID
            gbLocation = IIf(IsNull(Rec!chvLocation), "", Rec!chvLocation)
            gbTitle1 = IIf(IsNull(Rec!chvTitle), "", Rec!chvTitle)
            gbTitle2 = IIf(IsNull(Rec!chvAddressEnglish), "", Rec!chvAddressEnglish)
            gbnumZonalID = gbLocationID
            gbLBCODE = IIf(IsNull(Rec!chvLBCode), "", Rec!chvLBCode)
            
        End If
        Rec.Close
        
        mSql = "Select * From faDistrict Where intDistrictID=" & gbDistID
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            gbDistrict = Rec!vchDistrict
        End If
        Rec.Close
        '---------------------GST IN---------------------------------
        mSql = "Select * From faGStNo Where tnyActive=1 and tnyStatus=1"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            gbGSTIN = Rec!vchGSTNo
        Else
            gbGSTIN = ""
        End If
        Rec.Close
'        '-------------------------------------------------------
        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
            gbLBPanchayat = 1
        Else
            gbLBPanchayat = 0
        End If
        Set Rec = GetRecordSet("SELECT faFunds.intFundID,vchFundCode,vchFund FROM faSeats Inner Join faFunds On faSeats.intFundID = faFunds.intFundID WHERE numSeatID = " & gbSeatID)
        If Not (Rec.BOF And Rec.EOF) Then
            gbFundID = Rec!intFundID
            gbFundCode = Rec!vchFundCode + " " + Rec!vchFund
            gbSeatWiseFundID = Rec!intFundID
        End If
        Rec.Close
        '----------------gbLastPostingDate-------------------------------


        mSql = "SELECT MAX(dtPostingDate) dtPostingDate FROM faPostingIndex WHERE tnyStage=2"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            gbLastPostingDate = Format(Rec!dtPostingDate, "dd-mmm-yyyy")
        End If
        Rec.Close
        '----------------------------------------------------------------
        
        
        gbSeatGroupChairPerson = 1
        gbSeatGroupDyChairPersion = 2
        gbSeatGroupSecretary = 3
        gbSeatGroupAdditionalSecretary = 4
        gbSeatGroupAccountsOfficer = 5
        gbSeatGroupAccountsSuperintended = 6
        gbSeatGroupAccountsClerk = 7
        gbSeatGroupCashSuperintended = 8
        gbSeatGroupChiefCashier = 9
        gbSeatGroupCashier = 10
        gbSeatGroupDemandSectionSuperintended = 11
        gbSeatGroupDemandSectionClerk = 12
        gbSeatGroupAuditorsGroup = 13
        gbSeatGroupAccountSectionClerk = 15
        gbSeatGroupHeadClerk = 16
        gbSeatGroupAssistantSecretary = 17
        
        mSql = "Select * From faConfig"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            Call objAcc.SetAccountID(IIf(IsNull(Rec!intRegularTreasuryPensionAccountHeadID), 0, Rec!intRegularTreasuryPensionAccountHeadID))
            gbAcHeadCodeRegularTrasuryPension = objAcc.AccountCode
            Call objAcc.SetAccountID(IIf(IsNull(Rec!intContingentTreasuryPensionAccountHeadID), 0, Rec!intContingentTreasuryPensionAccountHeadID))
            gbAcHeadCodeContingentTreasuryPension = objAcc.AccountCode
            '--- For Manual Recipt Flag in Config       Sinoj
            gbManualReceiptNewBool = False
            On Error Resume Next
            If IsError(Rec!tnyManualReceipt) = False Then
                gbManualReceiptNewBool = IIf(Rec!tnyNewManualReceiptFormat = 1, True, False)
            End If
            On Error GoTo 0
            '------------------------------------------------'
        End If
        Rec.Close
        
        '-------------------------------------------------------------------------
        Dim objCounter As New clsCounter
        
        '------------------------------------------------------------------------------'
        ' Local Body Type Municipality And Corporation                                 '
        '------------------------------------------------------------------------------'
        If gbLBType = 3 Or gbLBType = 4 Then
        
            '-------------------------------------------------------------------------------------'
            '                   Nature of Funds   - Added By Poornima (used in frmBanks)
            '-------------------------------------------------------------------------------------'
                gbOwnFund = 4502
                gbSpecialFund = 4504
                gbGrantFund = 4506
                '----------------------------------------------------------------------------------------'
                '       Opening Balance -Capital Fund                -Added by Poornima On 28/12/2010    '
                '----------------------------------------------------------------------------------------'
                    gbAcHeadCodeForCapitalFund = "310100100"
                
                '----------------------------------------------------------------------------------------'
                
                '----------------------------------------------------------------------------------------'
                '       AccountHead used in KMBR                -Added by Anisha On 3/1/2011   '
                '----------------------------------------------------------------------------------------'
                    gbAcHeadIDOtherFee = 128
                    gbAcHeadCodeOtherFee = "140409900"
                
                '----------------------------------------------------------------------------------------'
                
                
                '----------------------------------------------------------------------------------------'
            '       AccountHead used in Allotment Letter                -Added by Vinod On 04/01/2011   '
                '----------------------------------------------------------------------------------------'
                    gbAcHeadCodeDevelopmentFundGeneralCapital = "320200101"
                    gbAcHeadCodeDevelopmentFundSCPCapital = "320200102"
                    gbAcHeadCodeDevelopmentFundTSPCapital = "320200103"
                    
                                
                    objAcc.SetAccountCode gbAcHeadCodeDevelopmentFundGeneralCapital
                    gbAcHeadIDDevelopmentFundGeneralCapital = objAcc.AccountHeadID
                                
                    objAcc.SetAccountCode gbAcHeadCodeDevelopmentFundSCPCapital
                    gbAcHeadIDDevelopmentFundSCPCapital = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeDevelopmentFundTSPCapital
                    gbAcHeadIDDevelopmentFundTSPCapital = objAcc.AccountHeadID
                                                                
                    gbAcHeadCodeTreasuryAccount1 = "450250100"
                    
                    
                    
                    gbAcHeadCodeTreasuryAccount2 = "450650100"
                    gbAcHeadCodeTreasuryAccount3 = "450650200"
                    gbAcHeadCodeTreasuryAccount4 = "450650300"
                    gbAcHeadCodeTreasuryAccount5 = "450650400"
                    
                    gbAcHeadCodeTreasuryAccount6 = "450650101"    ''Added by Minu on 08.12.2012 for new Treasury Accounts
                    gbAcHeadCodeTreasuryAccount7 = "450650102"
                    
                    gbAcHeadCodeTreasuryAccountTSB = "450250101" ''CHANGED FOR TREASURY_TSB
                    gbAcHeadCodeTreasuryAccountSpecialTSB = "450650103" ''CHANGED FOR Joint Venture
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount1
                    gbAcHeadIDTreasuryAccount1 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount2
                    gbAcHeadIDTreasuryAccount2 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount3
                    gbAcHeadIDTreasuryAccount3 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount4
                    gbAcHeadIDTreasuryAccount4 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount5
                    gbAcHeadIDTreasuryAccount5 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount6 ''Added by Minu on 08.12.2012 for new Treasury Accounts
                    gbAcHeadIDTreasuryAccount6 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount7
                    gbAcHeadIDTreasuryAccount7 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccountTSB ''CHANGED FOR TREASURY_TSB
                    gbAcHeadIDTreasuryAccountTSB = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccountSpecialTSB ''CHANGED FOR Joint Venture on 04 Mar 2017 By Anisha C
                    gbAcHeadIDTreasuryAccountSpecialTSB = objAcc.AccountHeadID
                    
                    
      
                    gbAcHeadCodeMaintenanceFundRoadAssets = "320200108" '"160100401"
                    gbAcHeadCodeMaintenanceFundNonRoadAssets = "320200109" '"160100402"
                    gbAcHeadCodeCentralFinanceCommission = "320200104"
                    gbAcHeadCodeKLGSDP = "320200105"
                    gbAcHeadCodeSpecialGrant = "320200106"
                    gbAcHeadCodeRoadRenovationGrant = "320200107"
                    
                    
                    
                    objAcc.SetAccountCode gbAcHeadCodeMaintenanceFundRoadAssets
                    gbAcHeadIDMaintenanceFundRoadAssets = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeMaintenanceFundNonRoadAssets
                    gbAcHeadIDMaintenanceFundNonRoadAssets = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeCentralFinanceCommission
                    gbAcHeadIDCentralFinanceCommission = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeKLGSDP
                    gbAcHeadIDKLGSDP = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeSpecialGrant
                    gbAcHeadIDSpecialGrant = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeRoadRenovationGrant
                    gbAcHeadIDRoadRenovationGrant = objAcc.AccountHeadID
                    
                    gbAcHeadCodeGeneralPurposeFund = "160100501"
                    
                    objAcc.SetAccountCode gbAcHeadCodeGeneralPurposeFund
                    gbAcHeadIDGeneralPurposeFund = objAcc.AccountHeadID
                    
                    
                    objAcc.SetAccountCode gbAcHeadCodeMaintenanceFundRoadAssets
                    gbAcHeadIDMaintenanceFundRoadAssets = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeMaintenanceFundNonRoadAssets
                    gbAcHeadIDMaintenanceFundNonRoadAssets = objAcc.AccountHeadID
                    
                    gbAcHeadCodeGeneralPurposeFund = "160100500"
                    
                    objAcc.SetAccountCode gbAcHeadCodeGeneralPurposeFund
                    gbAcHeadIDGeneralPurposeFund = objAcc.AccountHeadID
                    
                '----------------------------------------------------------------------------------------'
                
                '----------------------------------------------------------------------------------------'
                '       AccountHead used in Subsidiary Cash Book            -Added by Vinod On 10/01/2011'
                '----------------------------------------------------------------------------------------'
                    gbAcHeadCodeUnemploymentWages = "250600300"
                    objAcc.SetAccountCode gbAcHeadCodeUnemploymentWages
                    gbAcHeadIDUnemploymentWages = objAcc.AccountHeadID
                    
                    gbAcHeadCodeUnpaidSalaries = "350110300"
                    objAcc.SetAccountCode gbAcHeadCodeUnpaidSalaries
                    gbAcHeadIDUnpaidSalaries = objAcc.AccountHeadID
                    
                    gbAcHeadCodeVehicleHireCharges = "230400100"
                    objAcc.SetAccountCode gbAcHeadCodeVehicleHireCharges
                    gbAcHeadIDVehicleHireCharges = objAcc.AccountHeadID
                    
                    gbAcHeadCodeMiscAdministrationExpenses = "220809900"
                    objAcc.SetAccountCode gbAcHeadCodeMiscAdministrationExpenses
                    gbAcHeadIDMiscAdministrationExpenses = objAcc.AccountHeadID
                    
                    gbAcHeadCodeEquipmentHireCharges = "230400200"
                    objAcc.SetAccountCode gbAcHeadCodeEquipmentHireCharges
                    gbAcHeadIDEquipmentHireCharges = objAcc.AccountHeadID
                    
                    gbAcHeadCodeExpensesForBuryingUnclaimedDeadBodies = "230800300"
                    objAcc.SetAccountCode gbAcHeadCodeExpensesForBuryingUnclaimedDeadBodies
                    gbAcHeadIDExpensesForBuryingUnclaimedDeadBodies = objAcc.AccountHeadID
                    
                    gbAcHeadCodeRepairsAndMaintenanceDrainage = "230500400"
                    objAcc.SetAccountCode gbAcHeadCodeRepairsAndMaintenanceDrainage
                    gbAcHeadIDRepairsAndMaintenanceDrainage = objAcc.AccountHeadID
                    
                    gbAcHeadCodeDevFundProgrammesPublicHealthAndSanitation = "250401200"
                    objAcc.SetAccountCode gbAcHeadCodeDevFundProgrammesPublicHealthAndSanitation
                    gbAcHeadIDDevFundProgrammesPublicHealthAndSanitation = objAcc.AccountHeadID
                    
                    gbAcHeadCodeMiscAdvance = "460100700"
                    objAcc.SetAccountCode gbAcHeadCodeMiscAdvance
                    gbAcHeadIDMiscAdvance = objAcc.AccountHeadID
                '----------------------------------------------------------------------------------------'
                    gbAcHeadCodeCentralPensionFundPayable = "350110600"
                    gbAcHeadCodePensionAndGratuityPayable = "350110500"
                    gbAcHeadCodeOtherReceivablesCur = "431409901"
                    gbAcHeadCodePensionFundForContingentStaff = "311700100"
                    gbAcHeadCodeContributionToPensionFundForContingentStaff = "210300202"
        
                    gbFunctionaryAccountsDepartmentID = 4
                    gbFunctionaryAccountsDepartmentCode = "040000"
                    gbFunctionAccountsID = 6
                    gbFunctionAccountsCode = "00030100"
            
                    gbGeneralTransactionIDReceipts = 100 ' Should link to Master TransactionType Database Later
                    gbGeneralTransactionIDPayments = 200
                    gbGeneralTransactionIDContraE = 300
                    gbGeneralTransactionIDJournal = 400
                    
                    gbAcHeadCodePropertyTaxCurrent = "431100100"
                    gbAcHeadCodePropertyTaxArrear = "431100200"
                    gbAcHeadCodeLibraryCess = "350300100"
                    gbAcHeadCodePoorHomeCess = "350300200"
                    gbAcHeadCodeNoticeFee = "140400200"
            
            '''''-----------------------------------------------------
            '''' Added By Anisha On 11.10.12 For Property tax calculator
            '''''-----------------------------------------------------
                    gbAcHeadIDServicceCessCurrent = 1804
                    gbAcHeadIDServicceCessArrear = 1805
                    
                    gbAcHeadIDSurPTCurrent = 1806
                    gbAcHeadIDSurPTArrear = 1807
                    gbAcHeadIDSurCentralGovtBuildCurrent = 1808
                    gbAcHeadIDSurCentralGovtBuildArrear = 1809
                    gbAcHeadIDSplServicesCurrent = 1810
                    gbAcHeadIDSplServicesArrear = 1811
        
                    
                    gbAcHeadCodeServicceCessCurrent = 431800110
                    gbAcHeadCodeServicceCessArrear = 431800120
                    gbAcHeadCodeSurPTCurrent = 431800130
                    gbAcHeadCodeSurPTArrear = 431800140
                    gbAcHeadCodeSurCentralGovtBuildCurrent = 431800150
                    gbAcHeadCodeSurCentralGovtBuildArrear = 431800160
                    gbAcHeadCodeSplServicesCurrent = 431800170
                    gbAcHeadCodeSplServicesArrear = 431800180
                    
                '''''-----------------------------------------------------
                    gbAcHeadCodePenalInterest = "140200200"
                    gbAcHeadCodeRoundOff = "180809900"
                    gbAcHeadCodeAdvancePTax = "350410101"
                    
                    'Rent On Land
                    gbAcHeadCodeRLBArrear = "431400102"
                    gbAcHeadCodeRLBCurrent = "431400101"
                    gbAcHeadCodeAdvanceRLB = "350410404"
                    
                    gbAcHeadCodeRentLandArrear = "431400108"
                    gbAcHeadCodeRentLandCurrent = "431400107"
                    gbAcHeadCodeServiceTax = "350300500"
                    gbAcHeadCodeAdvanceLand = "350410404"
                    gbAcHeadCodeCGST = "350300700"   'Addded on 27 Sep 2017 for GST
                    gbAcHeadCodeSGST = "350300800"
                    gbAcHeadCodeFloodCess = "350300820" 'Addded on 27 Sep 2017 for FloodCess
                    
                    'Prof.Tax Institutions
                    gbAcHeadCodeProfTaxEmployees = "110100200"
                    gbAcHeadCodeProfTaxTraders = "110100100"
                    gbAcHeadCodeProfTaxTradersCurrent = "431190101" 'Added by poornim on 13/02/2012
                    gbAcHeadCodeProfTaxTradersArrears = "431190102" 'Added by poornim on 13/02/2012
                    
                    gbAcHeadCodeAdvanceDandO = "350410301"
                    
                    ''Addes By Vipin On 31.8.2011 For Exclude Advance FrommDeductionList
                    gbAcHeadDeductionStart = 350200200
                    gbAcHeadDeductionEnd = 350200218
                    gbAcHeadDeductionExcludeProfTax = 350200213
                
                    'Rent On Building
                    gbAcHeadCodeCivicAmenitiesArrear = "431400102"
                    gbAcHeadCodeCivicAmenitiesCurrent = "431400101"
                    gbAcHeadCodeAdvanceBuilding = "350410401"
                    
                    gbAcHeadCodeCash = "450100100"
                    gbAcHeadIDCash = 1504
                    
                    objAcc.SetAccountCode gbAcHeadCodePropertyTaxArrear
                    gbAcHeadIDPropertyTaxArrear = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodePropertyTaxCurrent
                    gbAcHeadIDPropertyTaxCurrent = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeLibraryCess
                    gbAcHeadIDLibraryCess = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodePoorHomeCess
                    gbAcHeadIDPoorHomeCess = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodePenalInterest
                    gbAcHeadIDPenalInterest = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeRoundOff
                    gbAcHeadIDRoundOff = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeAdvancePTax
                    gbAcHeadIDAdvancePTax = objAcc.AccountHeadID
                    
                    'Rent On Land/Building Variable For Keeping Id
                    objAcc.SetAccountCode gbAcHeadCodeCivicAmenitiesArrear
                    gbAcHeadIDCivicAmenitiesArrear = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeCivicAmenitiesCurrent
                    gbAcHeadIDCivicAmenitiesCurrent = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeAdvanceBuilding
                    gbAcHeadIDAdvanceBuilding = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeRentLandArrear
                    gbAcHeadIDRentLandArrear = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeRentLandCurrent
                    gbAcHeadIDRentLandCurrent = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeServiceTax
                    gbAcHeadIDServiceTax = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeAdvanceLand
                    gbAcHeadIDAdvanceLand = objAcc.AccountHeadID
                    
                    '''Added On 12 feb 2010 to Print Summary for Rent
                    objAcc.SetAccountCode gbAcHeadCodeCGST
                    gbAcHeadIDCGST = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeSGST
                    gbAcHeadIDSGST = objAcc.AccountHeadID
                    
                    ''Added on 1 Aug 2017
                    objAcc.SetAccountCode gbAcHeadCodeNoticeFee
                    gbAcHeadIDNoticeFee = objAcc.AccountHeadID
                    ' ------------------------------------------------------------------------------- '
                    ' P A Y M E N T S
                    ' ------------------------------------------------------------------------------- '
                    gbTransactionTypePayBills = 1001
                    gbTransactionTypeUnUtilizedAmount = 1401
                    gbTransactionTypeProjectExpGO = 1411
                    
                    gbAcHeadCodeGrossSalaryPayable = 350110100
                    objAcc.SetAccountCode gbAcHeadCodeGrossSalaryPayable
                    gbAcHeadIDGrossSalaryPayable = objAcc.AccountHeadID
                    
                    gbAcHeadCodeNetSalaryPayable = 350110200
                    objAcc.SetAccountCode gbAcHeadCodeNetSalaryPayable
                    gbAcHeadIDNetSalaryPayable = objAcc.AccountHeadID
                    

                    ' -------------------------------------------------- '
                    ' NOTE:- Rent On Land And Buildings                  '
                    ' Rent Receivable From Civic Amenities (Current)     '
                    ' -------------------------------------------------- '
                    objAcc.SetAccountCode gbAcHeadCodeRLBCurrent
                    gbAcHeadIDRLBCurrent = objAcc.AccountHeadID
                    
                    ' -------------------------------------------------- '
                    ' Rent Receivable From Civic Amenities (Arrear)      '
                    ' -------------------------------------------------- '
                    objAcc.SetAccountCode gbAcHeadCodeRLBArrear
                    gbAcHeadIDRLBArrear = objAcc.AccountHeadID
                    
                    gbTransactionTypePTax = 1
                    gbTransactionTypeRentOnBuilding = 4
                    gbTransactionTypeRentOnLand = 5
                    gbTransactionTypeProfTaxTrade = 2
                    gbTransactionTypeProfTaxEmp = 3
                    gbTransactionTypeDandO = 6
                    gbTransactionTypePFA = 7
                    
                    gbTransactionTypeProfTaxTradeAccrual = 202
                    gbTransactionTypeBrith = 12
                    gbTransactionTypeDeath = 12
                    gbTransactionTypeMarriage = 11
                    gbTransactionTypeCmnMarriage = 152
                    
                    gbTransactionTypeOutDoor = 9998
                    gbTransactionTypeZonalCollection = 9997
                    gbTransactionTypeFriendsCollection = 9996
                    
                    gbTransactionTypeETax = 17
                    gbTransactionTypeSTax = 18
                    gbTransactionTypeKCR = 9
                    gbTransactionTypePPR = 8
                    gbTransactionTypeHall = 19
            
                    gbTransactionTypeBFundSSSFund = 112
                    gbTransactionTypeMoneyOrderReturns = 74
                    
                    gbTransactionTypeMOReturnsSocialSecurityPension = 75
                    
                    gbTransactionTypeApplicationForPermitKMBR = 107
                    gbTransactionTypePermitFeeFromKMBR = 70
                    
                    gbTransactionTypeSaleOfTenderForm = 30
                    
                    gbTransactionTypeContraRegularPension = 4003
                    gbTransactionTypeContraContingentPension = 4004
                    
                    gbTransactiontypeDailyCollection = 4001
                    
                    gbTransactionTypeRefundOfPayment = 156
                    
                    gbTransactionTypeTransferCredit = 4006
                    
                    gbTransactionTypePTaxGp = 175 '' added on 27 Feb 2018
                    gbInstrumentCheque = 5
                    gbInstrumentCash = 1
                    gbInstrumentCard = 11
                    gbInstrumentLetterOfAuthority = 6
                    gbFundID = 1
                            
                    gbSeatByDeveloper = 100
                    gbSeatCashGroupID = 1
                    
                    objCounter.SetCounterByIP (GetIPAddress)
                    gbCounterNo = objCounter.CounterNo
                    gbCounterID = objCounter.CounterID
                    gbCounterIP = objCounter.CounterIP
                    gbCounterName = objCounter.CounterDescription
                    objCounter.CounterLogin objCounter.CounterNo, True
                    Set objCounter = Nothing
                    gbShiftID = 1
                    
                    gbSanchayaDbName = "SanchayaObjects" ' Uses to call Stored Procedures from Sanchaya DB
                    mSql = "Select * From faLBSettings"
                    mSql = "Select * From faConfig"
                    If Rec.State Then Rec.Close
                    Set Rec = GetRecordSet(mSql)
                    If Not (Rec.BOF And Rec.EOF) Then
                        gbPrinterMode = Rec!tnyPrinterMode
                        gbLinkWithPropertyTax = Rec!tnyLinkWithPropertyTax
                        gbLinkWithProfTaxEmp = Rec!tnyLinkProfessionTaxEmployee
                        gbLinkWithRentOnLand = Rec!tnyRLB
                        gbLinkWithFinanceHO = Rec!tnyLinkWithFinanceHO
                        gbFineCalculationMode = 1 '2 ' Changed on 28-Apr-2015 by Aiby:: Request by Mayoosh: Cross checked with Sanchaya: CINI
                        gbLinkWithSevana = Rec!tnyLinkWithSevana
                        gbLinkWithSugama = Rec!tnyLinkWithSugama
                        gbLinkWithDandOPFA = Rec!tnyLinkWithDandOPFA
                        gbFetchDemandFromHO = 0 'IIf(IsNull(Rec!tnyFetchDemandFromHO), 0, Rec!tnyFetchDemandFromHO)
                        gbLinkWithMOReturn = IIf(IsNull(Rec!tnyLinkWithMOReturn), 0, Rec!tnyLinkWithMOReturn)
                        gbLinkWithSoochika = IIf(IsNull(Rec!tnySoochikaUniCode), 1, Rec!tnySoochikaUniCode)
                 
                        '----Added By Anisha For Replacing INI File Details in dataBase (faConfig) On 25 Mar 2010
                        
                        gbDefaultBankID = IIf(IsNull(Rec!intDefaultBankID), -1, Rec!intDefaultBankID)
                        gbDefaultUrl = IIf(IsNull(Rec!vchDefaultUrl), "", Rec!vchDefaultUrl)
                        gbDefaultUrlForRequisition = IIf(IsNull(Rec!vchDefaultUrlForRequisition), "", Rec!vchDefaultUrlForRequisition)
                        gbRemittingBank = IIf(IsNull(Rec!vchRemittingBank), "", Rec!vchRemittingBank)
                        gbRemittingPlaceOfBank = IIf(IsNull(Rec!vchRemittingPlaceOfBank), "", Rec!vchRemittingPlaceOfBank)
                        gbDefaultTransactionTypeID = IIf(IsNull(Rec!intDefaultTransactionTypeID), -1, Rec!intDefaultTransactionTypeID)
                        gbFetchDemandFromWeb = IIf(IsNull(Rec!tnyWebDemandFlag), -1, Rec!tnyWebDemandFlag)  'Added On 18 Aug 2015
                        gbSaankhyaWeb = IIf(IsNull(Rec!tnySaankhyaWebFlag), 0, Rec!tnySaankhyaWebFlag)
                        
                        gbLinkWithDandOWeb = IIf(IsNull(Rec!tnyLinkWithDandOWeb), 0, Rec!tnyLinkWithDandOWeb)
                        
                        gbLinkWithProfTradeWeb = IIf(IsNull(Rec!tnyLinkWithProfTradeWeb), 0, Rec!tnyLinkWithProfTradeWeb)
                        'gbLinkWithProfEmpWeb = IIf(IsNull(Rec!tnyLinkWithProfTaxEmpWeb), 0, Rec!tnyLinkWithProfTaxEmpWeb)
                        gbLinkWithProfEmpWeb = 0
                        '----------------------------------------------------------------------------------------
                        ' If gbLBPanchayat = 1 Or gbLBType = 4 Then
                        
                        If IsNull(Rec!dtRPOpeningDate) = False Then
                            gbRPOnlinedate = IIf(IsNull(Rec!dtRPOpeningDate), "", Rec!dtRPOpeningDate)
                        Else
                            gbRPOnlinedate = Null
                        End If
                                                
                        If IsNull(Rec!dtOnlinedate) = False Then
                            gbOnlinedate = IIf(IsNull(Rec!dtOnlinedate), "", Rec!dtOnlinedate)
                        Else
                            gbOnlinedate = gbTransactionDate
                        End If
                        
                    End If
                    Rec.Close
                    Set Rec = GetRecordSet("SELECT  faSeats. intFundID,vchfundCode,vchFund  FROM   faSeats Inner Join faFunds On faSeats.intFundID = faFunds.intFundID WHERE  numSeatID = " & gbSeatID)
                    If Not (Rec.BOF And Rec.EOF) Then
                        gbFundID = Rec!intFundID
                        gbFundCode = Rec!vchFundCode + " " + Rec!vchFund
                    End If
                    Rec.Close
                    
                    If gbFetchDemandFromWeb = 1 Then
                        gbDefaultUrlSanchayaPost = gbDefaultUrl
                    End If
                    gbBold = Chr$(27) + Chr$(69)
                    gbBoldOff = Chr$(27) + Chr$(70)
                    
                    gbContense = Chr$(27) + Chr$(33) + Chr$(1) + Chr$(27) + Chr$(15)
                    gbContenseOff = Chr$(27) + Chr$(18)
                    gbContenseOff = Chr$(27) + Chr$(33) + Chr$(0)
                    
                    gbDoubleWidth = Chr$(27) + Chr$(87) + Chr$(1)
                    gbDoubleWidthOff = Chr$(27) + Chr$(87) + Chr$(0)
                    
        Else
        
            '----------------------
            '----------------------------------------------------------------------------------------'
            ' Local Body type: Panchayat                                                             '
            '----------------------------------------------------------------------------------------'
            'Changed On 23/12/2010
            'Coded By Poornima And Minu
            'set Environment for Panchayat
            
            '--Version-------------------
'            gbVerID = "2"
'            gbVerSubID = "2.1"
'            gbDBVerID = "1"
'            gbDBSubVerID = "0.3"
            '----------------------------
            
            gbOwnFund = 4502
            gbSpecialFund = 4504
            gbGrantFund = 4506
            
            '----------------------------------------------------------------------------------------'
            '       AccountHead used in KMBR                -Added by Anisha On 3/1/2011   '
            '----------------------------------------------------------------------------------------'
            
            '----------------------------PanjythHead for PTax-15-11-12----------------------------------------'
            
                    gbAcHeadCodeServicceCessCurrent = 431100105
                    gbAcHeadCodeServicceCessArrear = 431100106
                    gbAcHeadCodeServicceCessCurrentNonR = 431100107
                    gbAcHeadCodeServicceCessArrearNonR = 431100108
                    
                    gbAcHeadCodeSurPTCurrent = 431100109
                    gbAcHeadCodeSurPTArrear = 431100110
                    gbAcHeadCodeSurPTCurrentNonR = 431100111
                    gbAcHeadCodeSurPTArrearNonR = 431100112
                    
                    gbAcHeadCodeSurCentralGovtBuildCurrent = 431100113
                    gbAcHeadCodeSurCentralGovtBuildArrear = 431100114
                                        
                    gbAcHeadCodeSplServicesCurrent = 431100115
                    gbAcHeadCodeSplServicesArrear = 431100116
            '-----------------------------------------------------------------------------------------'
                    gbAcHeadIDOtherFee = 82
                    gbAcHeadCodeOtherFee = "140400199"
            '----------------------------------------------------------------------------------------'
        
            '----------------------------------------------------------------------------------------'
            '       AccountHead used in Allotment Letter                -Added by Vinod On 04/01/2011   '
            '----------------------------------------------------------------------------------------'
                    '' Added By Vipin On 31.8.11 For Exclude Advance from Deduction List
                    gbAcHeadDeductionStart = 350200200
                    gbAcHeadDeductionEnd = 350200300

                    
                    gbAcHeadCodeDevelopmentFundGeneralCapital = "320200101"
                    gbAcHeadCodeDevelopmentFundSCPCapital = "320200102"
                    gbAcHeadCodeDevelopmentFundTSPCapital = "320200103"
                                
                    objAcc.SetAccountCode gbAcHeadCodeDevelopmentFundGeneralCapital
                    gbAcHeadIDDevelopmentFundGeneralCapital = objAcc.AccountHeadID
                                
                    objAcc.SetAccountCode gbAcHeadCodeDevelopmentFundSCPCapital
                    gbAcHeadIDDevelopmentFundSCPCapital = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeDevelopmentFundTSPCapital
                    gbAcHeadIDDevelopmentFundTSPCapital = objAcc.AccountHeadID
                    
                    gbAcHeadCodeTreasuryAccount1 = "450250101"
                    gbAcHeadCodeTreasuryAccountTSB = "450250110" ''CHANGED FOR TREASRY_TSB
                    
                    gbAcHeadCodeTreasuryAccountSpecialTSB = "450650109"  ''For Joint Venture Project On 06 Mar 2017
                    
                    gbAcHeadCodeTreasuryAccount2 = "450650101"
                    gbAcHeadCodeTreasuryAccount3 = "450650102"
                    gbAcHeadCodeTreasuryAccount4 = "450650103"
                    gbAcHeadCodeTreasuryAccount5 = "450650104"
                    
                    gbAcHeadCodeTreasuryAccount6 = "450650105"    ''Added by Minu on 08.12.2012 for new Treasury Accounts
                    gbAcHeadCodeTreasuryAccount7 = "450650106"
                 
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount1
                    gbAcHeadIDTreasuryAccount1 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount2
                    gbAcHeadIDTreasuryAccount2 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount3
                    gbAcHeadIDTreasuryAccount3 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount4
                    gbAcHeadIDTreasuryAccount4 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount5
                    gbAcHeadIDTreasuryAccount5 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount6 ''Added by Minu on 08.12.2012 for new Treasury Accounts
                    gbAcHeadIDTreasuryAccount6 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccount7
                    gbAcHeadIDTreasuryAccount7 = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccountTSB ''CHANGED FOR TREASRY_TSB
                    gbAcHeadIDTreasuryAccountTSB = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeTreasuryAccountSpecialTSB ''For Joint Venture Project On 06 Mar 2017
                    gbAcHeadIDTreasuryAccountSpecialTSB = objAcc.AccountHeadID
                    
                    gbAcHeadCodeMaintenanceFundRoadAssets = "320200108"     '"160100401"
                    gbAcHeadCodeMaintenanceFundNonRoadAssets = "320200109"     '"160100402"
                    gbAcHeadCodeCentralFinanceCommission = "320200104"
                    gbAcHeadCodeKLGSDP = "320200105"
                    gbAcHeadCodeSpecialGrant = "320200106"
                    gbAcHeadCodeRoadRenovationGrant = "320200107"
                    
                    objAcc.SetAccountCode gbAcHeadCodeMaintenanceFundRoadAssets
                    gbAcHeadIDMaintenanceFundRoadAssets = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeMaintenanceFundNonRoadAssets
                    gbAcHeadIDMaintenanceFundNonRoadAssets = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeCentralFinanceCommission
                    gbAcHeadIDCentralFinanceCommission = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeKLGSDP
                    gbAcHeadIDKLGSDP = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeSpecialGrant
                    gbAcHeadIDSpecialGrant = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeRoadRenovationGrant
                    gbAcHeadIDRoadRenovationGrant = objAcc.AccountHeadID
                    
                    gbAcHeadCodeGeneralPurposeFund = "160100501"
                    
                    objAcc.SetAccountCode gbAcHeadCodeGeneralPurposeFund
                    gbAcHeadIDGeneralPurposeFund = objAcc.AccountHeadID
                    
                                        
                    gbAcHeadCodeIAY = "320100110"   'ADDED BY MINU FOR IAY
                    gbAcHeadCodeIAYSCP = "320100111"
                    gbAcHeadCodeIAYTSP = "320100112"

                    objAcc.SetAccountCode gbAcHeadCodeIAY
                    gbAcHeadIDIAY = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeIAYSCP
                    gbAcHeadIDIAYSCP = objAcc.AccountHeadID
                    
                    objAcc.SetAccountCode gbAcHeadCodeIAYTSP
                    gbAcHeadIDIAYTSP = objAcc.AccountHeadID
                    
             '----------------------------------------------------------------------------------------'
             '
             '----------------------------------------------------------------------------------------'
             '       AccountHead used in Subsidiary Cash Book            -Added by Vinod On 10/01/2011'
             '----------------------------------------------------------------------------------------'
'                    gbAcHeadCodeNetSalaryPayable = "350110102"
'
'                    objAcc.SetAccountCode gbAcHeadCodeNetSalaryPayable
'                    gbAcHeadIDNetSalaryPayable = objAcc.AccountHeadID

'''                    gbAcHeadCodeUnemploymentWages = "250600300"
'''                    objAcc.SetAccountCode gbAcHeadCodeUnemploymentWages
'''                    gbAcHeadIDUnemploymentWages = objAcc.AccountHeadID
                       
'''                    gbAcHeadCodeUnpaidSalaries = "350110300"
'''                    objAcc.SetAccountCode gbAcHeadCodeUnpaidSalaries
'''                    gbAcHeadIDUnpaidSalaries = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeVehicleHireCharges = "230400100"
'''                    objAcc.SetAccountCode gbAcHeadCodeVehicleHireCharges
'''                    gbAcHeadIDVehicleHireCharges = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeMiscAdministrationExpenses = "220809900"
'''                    objAcc.SetAccountCode gbAcHeadCodeMiscAdministrationExpenses
'''                    gbAcHeadIDMiscAdministrationExpenses = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeEquipmentHireCharges = "230400200"
'''                    objAcc.SetAccountCode gbAcHeadCodeEquipmentHireCharges
'''                    gbAcHeadIDEquipmentHireCharges = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeExpensesForBuryingUnclaimedDeadBodies = "230800300"
'''                    objAcc.SetAccountCode gbAcHeadCodeExpensesForBuryingUnclaimedDeadBodies
'''                    gbAcHeadIDExpensesForBuryingUnclaimedDeadBodies = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeRepairsAndMaintenanceDrainage = "230500400"
'''                    objAcc.SetAccountCode gbAcHeadCodeRepairsAndMaintenanceDrainage
'''                    gbAcHeadIDRepairsAndMaintenanceDrainage = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeDevFundProgrammesPublicHealthAndSanitation = "250401200"
'''                    objAcc.SetAccountCode gbAcHeadCodeDevFundProgrammesPublicHealthAndSanitation
'''                    gbAcHeadIDDevFundProgrammesPublicHealthAndSanitation = objAcc.AccountHeadID
'''
'''                    gbAcHeadCodeMiscAdvance = "460100700"
'''                    objAcc.SetAccountCode gbAcHeadCodeMiscAdvance
'''                    gbAcHeadIDMiscAdvance = objAcc.AccountHeadID
                    gbAcHeadCodeMiscAdvance = "460100103"
                    objAcc.SetAccountCode gbAcHeadCodeMiscAdvance
                    gbAcHeadIDMiscAdvance = objAcc.AccountHeadID
            '----------------------------------------------------------------------------------------'
        
            '----------------------------------------------------------------------------------------'
            ' Capital Fund - Added by Poornima On 28/12/2010                                         '
            ' Note:- this will set by login seat later                                               '
            '----------------------------------------------------------------------------------------'
            If gbFundID = 1 Or gbFundID = 2 Then
                gbAcHeadCodeForCapitalFund = "310100101"
            Else
                gbAcHeadCodeForCapitalFund = "310100102"
            End If
            gbAcHeadCodeCentralPensionFundPayable = "350110600"
            gbAcHeadCodePensionAndGratuityPayable = "350110500"
            gbAcHeadCodeOtherReceivablesCur = "431409901"
            gbAcHeadCodePensionFundForContingentStaff = "311700100"
            gbAcHeadCodeContributionToPensionFundForContingentStaff = "210300202"
            
            gbFunctionaryAccountsDepartmentID = 1
            gbFunctionaryAccountsDepartmentCode = "101"
            
            gbFunctionAccountsID = 2
            gbFunctionAccountsCode = "0002"
    
            
            gbGeneralTransactionIDReceipts = 100 ' Should link to Master TransactionType Database Later
            gbGeneralTransactionIDPayments = 200
            gbGeneralTransactionIDContraE = 300
            gbGeneralTransactionIDJournal = 400
            
            gbAcHeadCodePropertyTaxCurrent = "431100101"
            gbAcHeadCodePropertyTaxArrear = "431100102"
            gbAcHeadCodePropertyTax_NonResidential_Current = "431100103"
            gbAcHeadCodePropertyTax_NonResidential_Arrear = "431100104"
           
            
            
            gbAcHeadCodeLibraryCess = "350300101"
            gbAcHeadCodePoorHomeCess = "350300102"
            gbAcHeadCodeNoticeFee = "140400101"  ''Added On 31/Aug/2016
            
            
            gbAcHeadCodePenalInterest = "140200101"
            gbAcHeadCodeRoundOff = "180800199"
            gbAcHeadCodeAdvancePTax = "350410101"
            
            'Rent On Land
            gbAcHeadCodeRLBArrear = "431400102"
            gbAcHeadCodeRLBCurrent = "431400101"
            gbAcHeadCodeAdvanceRLB = "350410402"
                
            gbAcHeadCodeRentLandArrear = "431400102"
            gbAcHeadCodeRentLandCurrent = "431400101"
            gbAcHeadCodeServiceTax = "350300104"
            gbAcHeadCodeAdvanceLand = "350410402"
            gbAcHeadCodeCGST = "350300110"   'Addded on 27 Sep 2017 for GST
            gbAcHeadCodeSGST = "350300111"
            gbAcHeadCodeFloodCess = "350300116" 'Addded on 26 Aug 2019 for FloodCess
            
            'Prof.Tax Institutions
                
            gbAcHeadCodeProfTaxEmployees = "110200102"
            gbAcHeadCodeProfTaxTraders = "110200101"
            
            
            'Rent On Building
            gbAcHeadCodeCivicAmenitiesArrear = "431400102"
            gbAcHeadCodeCivicAmenitiesCurrent = "431400101"
            gbAcHeadCodeAdvanceBuilding = "350410401"
            
            'D and O Licence
            gbAcHeadCodeAdvanceDandO = "350410301"
            
            
            gbAcHeadCodeCash = "450100101"
            gbAcHeadIDCash = 1504
            
            objAcc.SetAccountCode gbAcHeadCodePropertyTaxArrear
            gbAcHeadIDPropertyTaxArrear = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodePropertyTaxCurrent
            gbAcHeadIDPropertyTaxCurrent = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodePropertyTax_NonResidential_Arrear
            gbAcHeadIDPropertyTax_NonResidential_Arrear = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodePropertyTax_NonResidential_Current
            gbAcHeadIDPropertyTax_NonResidential_Current = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeLibraryCess
            gbAcHeadIDLibraryCess = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodePoorHomeCess
            gbAcHeadIDPoorHomeCess = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodePenalInterest
            gbAcHeadIDPenalInterest = objAcc.AccountHeadID
            
                    
            objAcc.SetAccountCode gbAcHeadCodeRoundOff
            gbAcHeadIDRoundOff = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeAdvancePTax
            gbAcHeadIDAdvancePTax = objAcc.AccountHeadID
            
            'Rent On Land/Building Variable For Keeping Id
            objAcc.SetAccountCode gbAcHeadCodeCivicAmenitiesArrear
            gbAcHeadIDCivicAmenitiesArrear = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeCivicAmenitiesCurrent
            gbAcHeadIDCivicAmenitiesCurrent = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeAdvanceBuilding
            gbAcHeadIDAdvanceBuilding = objAcc.AccountHeadID
            
            
            objAcc.SetAccountCode gbAcHeadCodeRentLandArrear
            gbAcHeadIDRentLandArrear = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeRentLandCurrent
            gbAcHeadIDRentLandCurrent = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeServiceTax
            gbAcHeadIDServiceTax = objAcc.AccountHeadID
            
            objAcc.SetAccountCode gbAcHeadCodeAdvanceLand
            gbAcHeadIDAdvanceLand = objAcc.AccountHeadID
            
            ''Added on 26 Dec 2016
            objAcc.SetAccountCode gbAcHeadCodeNoticeFee
            gbAcHeadIDNoticeFee = objAcc.AccountHeadID
            
            ' ------------------------------------------------------------------------------- '
            ' P A Y M E N T S
            ' ------------------------------------------------------------------------------- '
                gbTransactionTypePayBills = 1001
                gbTransactionTypeUnUtilizedAmount = 1467
                gbTransactionTypeProjectExpGO = 1468
                gbAcHeadCodeGrossSalaryPayable = 350110101
                objAcc.SetAccountCode gbAcHeadCodeGrossSalaryPayable
                gbAcHeadIDGrossSalaryPayable = objAcc.AccountHeadID
                
                gbAcHeadCodeNetSalaryPayable = 350110102
                objAcc.SetAccountCode gbAcHeadCodeNetSalaryPayable
                gbAcHeadIDNetSalaryPayable = objAcc.AccountHeadID
                

            ' -------------------------------------------------- '
            ' NOTE:- Rent On Land And Buildings                  '
            ' Rent Receivable From Civic Amenities (Current)     '
            ' -------------------------------------------------- '
            objAcc.SetAccountCode gbAcHeadCodeRLBCurrent
            gbAcHeadIDRLBCurrent = objAcc.AccountHeadID
            
            ' -------------------------------------------------- '
            ' Rent Receivable From Civic Amenities (Arrear)      '
            ' -------------------------------------------------- '
            objAcc.SetAccountCode gbAcHeadCodeRLBArrear
            gbAcHeadIDRLBArrear = objAcc.AccountHeadID
            
            gbTransactionTypePTax = 1
            gbTransactionTypeRentOnBuilding = 4
            gbTransactionTypeRentOnLand = 5
            gbTransactionTypeProfTaxTrade = 2
            gbTransactionTypeProfTaxEmp = 3
            gbTransactionTypeDandO = 6
            gbTransactionTypePFA = 7
            
            gbTransactionTypeProfTaxTradeAccrual = 202
            gbTransactionTypeBrith = 12
            gbTransactionTypeDeath = 12
            gbTransactionTypeMarriage = 11
            gbTransactionTypeCmnMarriage = 152
            
            gbTransactionTypeOutDoor = 9998
            gbTransactionTypeZonalCollection = 9997
            
            gbTransactionTypeETax = 17
            gbTransactionTypeSTax = 18
            gbTransactionTypeKCR = 9
            gbTransactionTypePPR = 8
            gbTransactionTypeHall = 19
    
            gbTransactionTypeBFundSSSFund = 112
            gbTransactionTypeMoneyOrderReturns = 74
            
            gbTransactionTypeApplicationForPermitKMBR = 107
            gbTransactionTypePermitFeeFromKMBR = 70
            
            gbTransactionTypeSaleOfTenderForm = 30
            
            gbTransactionTypeContraRegularPension = 4003
            gbTransactionTypeContraContingentPension = 4004
            
            gbTransactiontypeDailyCollection = 4001
            
            gbTransactionTypeTransferCredit = 4010
            
            gbTransactionTypePTaxGp = 175
            
            gbInstrumentCheque = 5
            gbInstrumentCash = 1
            gbFundID = 1
                    
            gbSeatByDeveloper = 100
            gbSeatCashGroupID = 1
            
        
        
            objCounter.SetCounterByIP (GetIPAddress)
            gbCounterNo = objCounter.CounterNo
            gbCounterID = objCounter.CounterID
            gbCounterIP = objCounter.CounterIP
            gbCounterName = objCounter.CounterDescription
            objCounter.CounterLogin objCounter.CounterNo, True
            Set objCounter = Nothing
            
            gbShiftID = 1
                
            
            gbSanchayaDbName = "SanchayaObjects" ' Uses to call Stored Procedures from Sanchaya DB
            mSql = "Select * From faLBSettings"
            mSql = "Select * From faConfig"
            If Rec.State Then Rec.Close
            Set Rec = GetRecordSet(mSql)
            If Not (Rec.BOF And Rec.EOF) Then
                gbLinkWithPropertyTax = Rec!tnyLinkWithPropertyTax
                gbLinkWithProfTaxEmp = Rec!tnyLinkProfessionTaxEmployee
                gbLinkWithRentOnLand = Rec!tnyRLB
                gbLinkWithFinanceHO = Rec!tnyLinkWithFinanceHO
                gbFineCalculationMode = 1
                gbLinkWithSevana = Rec!tnyLinkWithSevana
                gbLinkWithSugama = Rec!tnyLinkWithSugama
                gbLinkWithDandOPFA = Rec!tnyLinkWithDandOPFA
                gbFetchDemandFromHO = 0 'IIf(IsNull(Rec!tnyFetchDemandFromHO), 0, Rec!tnyFetchDemandFromHO)
                gbLinkWithMOReturn = IIf(IsNull(Rec!tnyLinkWithMOReturn), 0, Rec!tnyLinkWithMOReturn)
                gbLinkWithSoochika = IIf(IsNull(Rec!tnySoochikaUniCode), 1, Rec!tnySoochikaUniCode)
                
                '----Added By Anisha For Replacing INI File Details in dataBase (faConfig) On 25 Mar 2010
                gbDefaultBankID = IIf(IsNull(Rec!intDefaultBankID), -1, Rec!intDefaultBankID)
                gbDefaultUrl = IIf(IsNull(Rec!vchDefaultUrl), "", Rec!vchDefaultUrl)
                gbDefaultUrlForRequisition = IIf(IsNull(Rec!vchDefaultUrlForRequisition), "", Rec!vchDefaultUrlForRequisition)
                gbRemittingBank = IIf(IsNull(Rec!vchRemittingBank), "", Rec!vchRemittingBank)
                gbRemittingPlaceOfBank = IIf(IsNull(Rec!vchRemittingPlaceOfBank), "", Rec!vchRemittingPlaceOfBank)
                gbDefaultTransactionTypeID = IIf(IsNull(Rec!intDefaultTransactionTypeID), -1, Rec!intDefaultTransactionTypeID)
                
                gbFetchDemandFromWeb = IIf(IsNull(Rec!tnyWebDemandFlag), -1, Rec!tnyWebDemandFlag)  'Added On 18 Aug 2015
                '----------------------------------------------------------------------------------------
                gbSaankhyaWeb = IIf(IsNull(Rec!tnySaankhyaWebFlag), 0, Rec!tnySaankhyaWebFlag)
                
                 'gbLinkWithDandOWeb = IIf(IsNull(Rec!tnyLinkWithDandOWeb), 0, Rec!tnyLinkWithDandOWeb)
                 
                ' gbLinkWithProfTradeWeb = IIf(IsNull(Rec!tnyLinkWithProfTradeWeb), 0, Rec!tnyLinkWithProfTradeWeb)
                
                 'gbLinkWithProfEmpWeb = IIf(IsNull(Rec!tnyLinkWithProfTaxEmpWeb), 0, Rec!tnyLinkWithProfTaxEmpWeb)
                  gbLinkWithProfEmpWeb = 0
                'If gbLBPanchayat = 1 Or gbLBType = 4 Then
                If IsNull(Rec!dtRPOpeningDate) = False Then
                    gbRPOnlinedate = IIf(IsNull(Rec!dtRPOpeningDate), "", Rec!dtRPOpeningDate)
                Else
                    gbRPOnlinedate = Null
                End If
                
                If IsNull(Rec!dtOnlinedate) = False Then
                    gbOnlinedate = IIf(IsNull(Rec!dtOnlinedate), "", Rec!dtOnlinedate)
                Else
                    gbOnlinedate = gbTransactionDate
                End If
            End If
            Rec.Close
            Set Rec = GetRecordSet("SELECT  faSeats. intFundID,vchfundCode,vchFund  FROM   faSeats Inner Join faFunds On faSeats.intFundID = faFunds.intFundID WHERE  numSeatID = " & gbSeatID)
            If Not (Rec.BOF And Rec.EOF) Then
                gbFundID = Rec!intFundID
                gbFundCode = Rec!vchFundCode + " " + Rec!vchFund
            End If
            Rec.Close
            
            If gbFetchDemandFromWeb = 1 Then
                gbDefaultUrlSanchayaPost = gbDefaultUrl
            End If
            gbBold = Chr$(27) + Chr$(69)
            gbBoldOff = Chr$(27) + Chr$(70)
            
            gbContense = Chr$(27) + Chr$(33) + Chr$(1) + Chr$(27) + Chr$(15)
            gbContenseOff = Chr$(27) + Chr$(18)
            gbContenseOff = Chr$(27) + Chr$(33) + Chr$(0)
            
            gbDoubleWidth = Chr$(27) + Chr$(87) + Chr$(1)
            gbDoubleWidthOff = Chr$(27) + Chr$(87) + Chr$(0)
        End If
        SetEnvironment = True
        Exit Function

ErrorCheck:
        SetEnvironment = False
    End Function
      Public Function FormatIntoProperCase(mString As String) As String
        '========================================
        ' Written by : Vinod     Date : 15-10-08
        '=======================================
        
        'Dim mString     As String
        Dim mLength     As Integer
        Dim mStart      As Integer
        Dim mCount      As Integer
        Dim mArray      As Variant
        Dim mStr        As String
        Dim mAscii      As Integer
        
        mStart = 1
        mString = Trim(mString)
        mLength = Len(mString)
        For mCount = 0 To mLength
            If mArray = " " Or mArray = "" Or mArray = "." Or mAscii = 13 Or mAscii = 32 Or mAscii = 10 Then
                mArray = UCase(mID(mString, mStart, 1))
                mStart = mStart + 1
                If mArray <> "" Then
                    mAscii = Asc(mArray)
                End If
            Else
                mArray = mID(mString, mStart, 1)
                mStart = mStart + 1
                If mArray <> "" Then
                    mAscii = Asc(mArray)
                End If
            End If
            mStr = mStr + mArray
        Next mCount
        FormatIntoProperCase = mStr

    End Function
    Public Sub PressTabKey()
        '-------------------------------------------------------------------'
        ' Aiby : 15-May,2003                                                '
        ' To trace Enter key                                                '
        '             Function EnterKey() calls this in Module level        '
        '             keybd_even is API (keybd_event)                       '
        '             ib_Tab is constant holds Ascii value of tab           '
        ' Objective : to solve the Number Lock ON/OFF while SendKey         '
        '-------------------------------------------------------------------'
        keybd_event ib_Tab, 0, 0, 0  ' press Tab
                                        keybd_event ib_Tab, 0, KEYEVENTF_KEYUP, 0  ' release Tab
    
    End Sub
    
    Public Sub PrinterInit()
        Dim lReturn As Long
        Dim sPrinter As String
        Dim lhPrinter As Long
        Dim gbPrintStatus As Boolean
        
        'Writen by Aiby for Initialize the required control field
        gbPageWidth = 80
        gbTopMargin = 5
        gbBottomMargin = 5
        gbLeftMargin = 0
        gbRightMargin = 0
        gbNoOfLinesPerPage = 70
        gbNoOfPrintableLines = gbNoOfLinesPerPage - (gbTopMargin + gbBottomMargin)
        gbTextArea = gbPageWidth - (gbLeftMargin + gbRightMargin)
        
        'sPrinter = Printer.DeviceName
        sPrinter = "LPT1"
        'lReturn = OpenPrinter(sPrinter, lhPrinter, 0)
        gbPrintStatus = True
            'If lReturn = 0 Then
            '    MsgBox "Printer is not connected or either turned off...!"
            '    gbPrintStatus = False
            '    Exit Sub
            'Else
            '    gbPrintStatus = True
            'End If
    
        gbPrinterPort = Printer.Port
        gbFileNO = FreeFile
        gbTempFileName = App.Path + "\Temp.txt"
        'gbFileName = App.Path + "\" + Printer.Port
        gbFileName = Printer.Port
        On Error GoTo ErrorSkip:
        Open gbFileName For Output As #gbFileNO
        Exit Sub
ErrorSkip:
        MsgBox "Printer Not Connected!", vbInformation
        
    End Sub
    
    
    Public Sub FileInitialize(Optional mFileName = "Report.txt")
            'Writen by Aiby for Initialize the required control field
            gbPageWidth = 80
            gbTopMargin = 5
            gbBottomMargin = 5
            gbLeftMargin = 2
            gbRightMargin = 2
            gbNoOfLinesPerPage = 70
            gbNoOfPrintableLines = gbNoOfLinesPerPage - (gbTopMargin + gbBottomMargin)
            
            gbTextArea = gbPageWidth - (gbLeftMargin + gbRightMargin)
            
            'gbPrinterPort = Printer.Port
            Reset
            gbFileNO = FreeFile
            gbFileName = App.Path + "\" + mFileName
            Open gbFileName For Output As #gbFileNO
    End Sub
    
    Public Sub ShellPad()
        On Error Resume Next
        Shell "NotePad " & gbFileName, vbMaximizedFocus
        On Error GoTo 0
    End Sub

    Public Function FetchFieldValue(mTable As String, mFieldName As String, mWHERE As String)
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            mSql = "Select " & mFieldName & " From " & mTable & " " & mWHERE
            Rec.CursorLocation = adUseClient
            'On Error GoTo ErrExit
            Rec.Open mSql, gbCnn, adOpenForwardOnly, adLockReadOnly
            If Rec.RecordCount > 0 Then
                FetchFieldValue = Rec.Fields(0).Value
            End If
            
            Exit Function
ErrExit:
    
    End Function
    
    Public Function CheckDateInMMM(Dt As String) As String
        '========================================
        ' Written by : Aiby     Date : 01-08-98
        '=======================================
        
        'Very Very Dangerous Function :-warrning by Shanker
                Dim fDt As String
                Dim X As Integer
                Dim Dd As String
                Dim Mm As String
                Dim mS As String
                Dim Yy As String
                Dim T As String
                Dim mmm(1 To 12) As String
                
                mmm(1) = "Jan"
                mmm(2) = "Feb"
                mmm(3) = "Mar"
                mmm(4) = "Apr"
                mmm(5) = "May"
                mmm(6) = "Jun"
                mmm(7) = "Jul"
                mmm(8) = "Aug"
                mmm(9) = "Sep"
                mmm(10) = "Oct"
                mmm(11) = "Nov"
                mmm(12) = "Dec"
        
                For X = 1 To Len(Dt)
                    Select Case X
                            Case 1, 2
                                    If IsNumeric(mID(Dt, X, 1)) Then
                                        Dd = Dd + mID(Dt, X, 1)
                                    ElseIf mID(Dt, X, 1) <> " " Or mID(Dt, X, 1) <> "\" Or mID(Dt, X, 1) <> "-" Then
                                        mS = mS + mID(Dt, X, 1)
                                    End If
                            Case Else
                                    If Len(Mm) < 2 Then
                                            If IsNumeric(mID(Dt, X, 1)) Then
                                                Mm = Mm + mID(Dt, X, 1)
                                            'ElseIf (Mid(Dt, X, 1) <> " " Or Mid(Dt, X, 1) <> "/" Or Mid(Dt, X, 1) <> "-") And X <> 3 Then
                                            '    T = MM
                                            '    MM = "0" + T
                                            Else
                                                    If Len(Mm) > 0 Then
                                                            T = Mm
                                                            Mm = "0" + T
                                                    End If
                                                    mS = mS + mID(Dt, X, 1)
                                            End If
                                    ElseIf IsNumeric(mID(Dt, X, 1)) And Len(Yy) < 4 Then
                                            Yy = Yy + mID(Dt, X, 1)
                                    ElseIf mID(Dt, X, 1) <> " " Or mID(Dt, X, 1) <> "/" Or mID(Dt, X, 1) <> "-" Then
                                            mS = mS + mID(Dt, X, 1)
                                    End If
                        End Select
                Next X
                    
                    For X = 1 To 12
                            If InStr(1, mS, mmm(X), vbTextCompare) Then Exit For
                    Next X
                
                    If X < 13 Then
                            If Dd = "" Then
                                Dd = Mm
                            Else
                                T = Yy
                                Yy = Mm + T
                            End If
                            Mm = Format(X, "00")
                    Else
                            If val(Mm) > 12 And val(Dd) < 13 Then
                                T = Dd
                                Dd = Mm
                                Mm = T
                                If val(Dd) > 31 Then Dd = Dd Mod 31
                                If val(Mm) > 12 Then Mm = Mm Mod 12
                            Else
                                If val(Dd) > 31 Then Dd = Dd Mod 31
                                If val(Mm) > 12 Then Mm = Mm Mod 12
                            End If
                    End If
                
                    If val(Yy) < 100 And val(Yy) > 50 Then
                                T = Trim(str(val(Yy)))
                                Yy = "19" + T
                     ElseIf val(Yy) <= 50 And val(Yy) > 0 Then
                                T = Format(Trim(str(val(Yy))), "00")
                                Yy = "20" + T
                    ElseIf val(Yy) > 100 And val(Yy) < 1000 Then
                                T = Trim(str(val(Yy)))
                                Yy = "2" + T
                    ElseIf val(Yy) = 0 Then
                                Yy = Year(gbDate)
                    End If
                      
                If Yy > Year(gbDate) + 100 Then Yy = Year(gbDate)
                If val(Dd) = 0 Then Dd = Format(Day(gbDate), "00")
                If val(Mm) = 0 Then Mm = Format(Month(gbDate), "00")
                If Yy = "" Then Yy = Format(Year(gbDate), "0000")
                
                Dd = Format(Dd, "00")
                Mm = Format(Mm, "00")
                Yy = Format(Yy, "0000")
            
                fDt = DateSerial(Yy, Mm, Dd)
                
                If Not IsDate(fDt) Then
                    MsgBox (Dt + "  can't convert into date.")
                    CheckDateInMMM = Format(Day(Now), "00") + "-" + Format(Month(Now), "00") + "-" + Format(Year(Now), "0000")
                Else
                    If val(Mm) = 2 Then
                       If Day(DateSerial(Yy, Mm, Dd)) <> val(Dd) Then Dd = 28
                    End If
                    CheckDateInMMM = Dd + "-" + mmm(Mm) + "-" + Yy
                End If
    
    End Function

    Function DdMmYy(Dt As Date) As String
        '========================================
        ' Written by : Aiby     Date : 10-05-98
        '=======================================
        DdMmYy = Format(Day(Dt), "00") + "/" + Format(Month(Dt), "00") + "/" + Format(Year(Dt), "0000")
    End Function
    
    Function DdMmmYy(Dt As Date) As String
        '========================================
        ' Written by : Aiby     Date : 10-05-98
        '=======================================
        DdMmmYy = Format(Day(Dt), "00") + "-" + Format(Dt, "mmm") + "-" + Format(Year(Dt), "0000")
    End Function
    
    Function MmDdYy(Dt As Date) As String
        '========================================
        ' Written by : Aiby     Date : 10-05-98
        '=======================================
        MmDdYy = Format(Month(Dt), "00") + "/" + Format(Day(Dt), "00") + "/" + Format(Year(Dt), "0000")
    End Function

    Public Function FormatDate(Dt As String) As Date
        '========================================
        ' Written by : Aiby     Date : 01-07-99
        '=======================================
        ' Input in dd/mm/yyyy and output in system format
        Dim mSDate As String
        Dim mD As Integer
        Dim Mm As Integer
        Dim mY As Integer
        
        mSDate = Dt
        If mSDate = "" Then Exit Function
        mD = val(Left(mSDate, 2))
        Mm = val(mID(mSDate, 4, 2))
        mY = val(Right(mSDate, 4))
        FormatDate = DateSerial(mY, Mm, mD)
    End Function
    
    Public Function FormatAmountToValue(mStrAmount As String) As Single
        Dim mTmpAmt As String
        Dim mLen As Integer
        Dim mDigit As String
        mLen = Len(mStrAmount)
        For mLen = 1 To Len(mStrAmount)
            mDigit = mID$(mStrAmount, mLen, 1)
            If (Asc(mDigit) > 47 And Asc(mDigit) < 58) Or Asc(mDigit) = Asc(".") Then
                mTmpAmt = mTmpAmt + mDigit
            End If
        Next
        FormatAmountToValue = val(mTmpAmt)
    End Function

    Public Function FindMaster(mTableName As String, mOutPutFieldName As String, mInputFieldName As String, ByVal mID As Long, Optional ConnStr As enuSourceString = -1) As String
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mCn As New ADODB.Connection
        Dim objdb As clsDB
        
        mSql = "SELECT " & mOutPutFieldName & " FROM " & mTableName & " WHERE " & mInputFieldName & " = " & mID
        Rec.CursorLocation = adUseClient
        
        If ConnStr = -1 Then
            Set mCn = gbCnn
            If mCn.State = 0 Then
                Set objdb = New clsDB
                objdb.SetConnection mCn
            End If
        Else
            Set objdb = New clsDB
            objdb.CreateNewConnection mCn, ConnStr
            If mCn.State = 0 Then
                FindMaster = ""
                Exit Function
            End If
        End If
        
        Rec.Open mSql, mCn, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Rec.RecordCount > 0 Then
            On Error GoTo SkipNull:
            FindMaster = Rec.Fields(0).Value
        Else
SkipNull:
            FindMaster = ""
        End If
        Rec.Close
        Set mCn = Nothing
        Set objdb = Nothing
    End Function
    
    Public Function FindMasterID(mTableName As String, mOutPutFieldName As String, mInputFieldName As String, ByVal mInputFieldValue As String, Optional ConnStr As enuSourceString = -1) As Long
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mCn As New ADODB.Connection
        Dim objdb As clsDB
        
        mSql = "SELECT " & mOutPutFieldName & " FROM " & mTableName & " WHERE " & mInputFieldName & " = '" & mInputFieldValue & "'"
        Rec.CursorLocation = adUseClient
        If ConnStr = -1 Then
            Set mCn = gbCnn
            If mCn.State = 0 Then
                Set objdb = New clsDB
                objdb.SetConnection mCn
            End If
        Else
            Set objdb = New clsDB
            objdb.CreateNewConnection mCn, ConnStr
            If mCn.State = 0 Then
                FindMasterID = ""
                Exit Function
            End If
        End If
        
        Rec.Open mSql, mCn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Rec.RecordCount > 0 Then
                FindMasterID = Rec.Fields(0).Value
            Else
                FindMasterID = -1
            End If
        Rec.Close
    End Function
    
    Public Function FindMasterIDbyID(mTableName As String, mOutPutFieldName As String, mInputFieldName As String, ByVal mInputFieldValue As Long, Optional ConnStr As enuSourceString = -1) As Long
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mCn As New ADODB.Connection
        Dim objdb As clsDB
        
        mSql = "SELECT " & mOutPutFieldName & " FROM " & mTableName & " WHERE " & mInputFieldName & " = " & mInputFieldValue
        Rec.CursorLocation = adUseClient
        
        If ConnStr = -1 Then
            Set mCn = gbCnn
            If mCn.State = 0 Then
                Set objdb = New clsDB
                objdb.SetConnection mCn
            End If
        Else
            Set objdb = New clsDB
            objdb.CreateNewConnection mCn, ConnStr
            If mCn.State = 0 Then
                FindMasterIDbyID = -1
                Exit Function
            End If
        End If
        
        Rec.Open mSql, gbCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Rec.RecordCount > 0 Then
                FindMasterIDbyID = Rec.Fields(0).Value
            Else
                FindMasterIDbyID = -1
            End If
        Rec.Close
    End Function

    Public Function FindTableFieldValue(mTabelName As String, mFieldName As String, Optional mCondition As String) As Variant
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        
        mSql = "SELECT " & mFieldName & " FROM " & mTabelName
        If Not Trim(mCondition) = "" Then
            mSql = mSql + " WHERE " & mCondition
        End If
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, gbCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Rec.RecordCount > 0 Then
                FindTableFieldValue = Rec.Fields(0).Value
            Else
                FindTableFieldValue = Null
            End If
        Rec.Close
    End Function
    
    Public Function MonthInWords(mMonth As Integer) As String
        If mMonth > 12 Then
            mMonth = mMonth Mod 12
        End If
        Select Case mMonth
            Case Is = 1: MonthInWords = "January"
            Case Is = 2: MonthInWords = "February"
            Case Is = 3: MonthInWords = "March"
            Case Is = 4: MonthInWords = "April"
            Case Is = 5: MonthInWords = "May"
            Case Is = 6: MonthInWords = "June"
            Case Is = 7: MonthInWords = "July"
            Case Is = 8: MonthInWords = "August"
            Case Is = 9: MonthInWords = "September"
            Case Is = 10: MonthInWords = "October"
            Case Is = 11: MonthInWords = "November"
            Case Is = 12: MonthInWords = "December"
        End Select
        
    End Function


    Function Heading(s As String, Optional Align As String, Optional b As Boolean, Optional d As Boolean)
        '=============================================================================
        ' Written by : Aiby     Date : 01-07-99
        '                   Last Modified Date : 01/01/2000
        '=============================================================================
        
        Dim Width As Long
        Dim mDwidth As Integer
        
        mDwidth = gbTextArea / 2
        
        s = Trim(s)
        If s = "" Then Exit Function
        
        Width = Len(s)
        If Not IsMissing(d) And d = True Then
                If Width > mDwidth Then s = Left(s, mDwidth)
                Width = Len(s)
        Else
                If Width > gbTextArea Then s = Left(s, gbTextArea)
        End If
        
        If Not IsMissing(Align) Then
                If Not IsMissing(d) Then
                        If d = False Then
                                Select Case Align
                                Case "L"
                                        'If Width < 70 Then S = Space(4) + S
                                        s = Space(gbLeftMargin) + s
                                Case "R"
                                        If Width <= gbTextArea Then s = Space(gbLeftMargin) + Space(gbTextArea - Width) + s
                                Case Else
                                        If Width <= gbTextArea Then s = Space(gbLeftMargin) + Space(Int((gbTextArea - Width) / 2)) + s + Space(Int((gbTextArea - Width) / 2))
                                End Select
                        Else
                                Select Case Align
                                Case "L"
                                        If Width <= mDwidth Then s = Space(gbLeftMargin / 2) + s
                                Case "R"
                                        If Width <= mDwidth Then s = Space(gbLeftMargin / 2) + Space(mDwidth - Width) + s
                                Case Else
                                        If Width <= mDwidth Then s = Space(gbLeftMargin / 2) + Space((mDwidth - Width) / 2) + s + Space((mDwidth - Width) / 2)
                                End Select
                        End If
                Else
                        Select Case Align
                        Case "L"
                                If Width <= gbTextArea Then s = Space(gbLeftMargin) + s
                        Case "R"
                                If Width <= gbTextArea Then s = Space(gbLeftMargin) + Space(gbTextArea - Width) + s
                        Case Else
                                If Width <= gbTextArea Then s = Space(gbLeftMargin) + Space(Int((gbTextArea - Width) / 2)) + s + Space(Int((gbTextArea - Width) / 2))
                        End Select
                End If
        End If
        
        If Not IsMissing(b) Then
        If b Then s = Chr$(27) + Chr$(69) + s + Chr$(27) + Chr$(70)
        End If
        
        If Not IsMissing(d) Then
        If d Then s = Chr$(27) + Chr$(14) + s + Chr$(27) + Chr$(Asc("DC4")) + Chr(13)
        End If
        Heading = s
    End Function
    
    Function Style(s As String, Optional b As Boolean, Optional d As Boolean) As String
        '========================================
        ' Written by : Aiby       Date : 16-08-99
        '========================================
        'S = Trim(S)
        
        If Not IsMissing(b) Then
        If b Then s = Chr$(27) + Chr$(69) + s + Chr$(27) + Chr$(70)
        End If
        
        If Not IsMissing(d) Then
        If d Then s = Chr$(27) + Chr$(14) + s + Chr$(27) + Chr$(Asc("DC4")) '+ Chr(13)
        End If
        Style = s
    End Function

    
    Public Function PadC(s As String, Width As Long) As String
        Dim Ln As Integer
        Dim diff As Integer
        
        Ln = Len(s)
        If Ln > Width Then
            s = Left$(s, Width)
        Else
            diff = Width - Ln
            s = Trim(s)
            s = Space(diff / 2) + s + Space(diff / 2)
        End If
        PadC = s
    End Function
    
    Public Function PadL(s As String, Width As Long) As String
        Dim Ln As Integer
        Dim diff As Integer
        
        Ln = Len(s)
        If Ln > Width Then
            s = Left$(s, Width)
        Else
            diff = Width - Ln
            s = Trim(s)
            s = Space(diff) + s
        End If
        PadL = s
    End Function
    
    Public Function PadR(s As String, Width As Long) As String
        Dim Ln As Integer
        Dim diff As Integer
      
        Ln = Len(s)
        If Ln > Width Then
            s = Left$(s, Width)
        Else
            diff = Width - Ln
            s = Trim(s)
            s = s + Space(diff)
        End If
        PadR = s
        
    End Function

    Public Function Rupees(mRs As Long) As String
            Dim mR As Long
            Dim mP As Long
            mRs = Format(mRs, "0.00")
            mR = Int(mRs)
            mP = (mRs - mR) * 100
            If mP > 0 Then
                Rupees = Words(mR) + " Rupees & " + Words(mP) + " Paisa"
            Else
                Rupees = Words(mR) + " Rupees Only"
            End If
    End Function
        
        
        
    Public Function Words(mNum As Long) As String
        
                Dim ONES(1 To 9) As String
                Dim TEENS(11 To 19) As String
                Dim TENS(1 To 9) As String
                Dim mTemp As String
                Static mAndFlag As Boolean
                
                ONES(1) = "One": ONES(2) = "Two": ONES(3) = "Three": ONES(4) = "Four": ONES(5) = "Five": ONES(6) = "Six": ONES(7) = "Seven": ONES(8) = "Eight": ONES(9) = "Nine"
                TEENS(11) = "Eleven": TEENS(12) = "Twelve": TEENS(13) = "Thirteen": TEENS(14) = "Fourteen": TEENS(15) = "Fifteen": TEENS(16) = "Sixteen": TEENS(17) = "Seventeen": TEENS(18) = "Eighteen": TEENS(19) = "Ninteen"
                TENS(1) = "Ten": TENS(2) = "Twenty": TENS(3) = "Thirty": TENS(4) = "Forty": TENS(5) = "Fifty": TENS(6) = "Sixty": TENS(7) = "Seventy": TENS(8) = "Eighty": TENS(9) = "Ninty"
                
                
                If mNum < 1000000000 And mNum > 99999999 Then
                    mAndFlag = False
                    mTemp = Words(Int(mNum / 10000000)) + " Crore"
                    mNum = mNum Mod 10000000
                    If mNum > 0 Then
                        mAndFlag = True
                        mTemp = mTemp + Words(mNum)
                    End If
                ElseIf mNum < 100000000 And mNum > 9999999 Then
                    mAndFlag = False
                    mTemp = Words(Int(mNum / 10000000)) + " Crore"
                    mNum = mNum Mod 10000000
                    If mNum > 0 Then
                        mAndFlag = True
                        mTemp = mTemp + Words(mNum)
                    End If
                '--------------------------
'''                ElseIf mNum < 10000000 And mNum > 999999 Then
'''                    mAndFlag = False
'''                    mTemp = Words(Int(mNum / 1000000)) + " Ten Lakhs"
'''                    mNum = mNum Mod 1000000
'''                    If mNum > 0 Then
'''                        mAndFlag = True
'''                        mTemp = mTemp + Words(mNum)
'''                    End If
                '-------------------------------------
''''                ElseIf mNum < 1000000 And mNum > 99999 Then
''''                    mAndFlag = False
''''                    mTemp = Words(Int(mNum / 100000)) + " Lakhs"
''''                    mNum = mNum Mod 100000
''''                    If mNum > 0 Then
''''                        mAndFlag = True
''''
''''                        mTemp = mTemp + Words(mNum)
''''                    End If
''''
                '---------------------------------------
                
                '-------------------------------------
                ElseIf mNum < 10000000 And mNum > 99999 Then
                    mAndFlag = False
                    mTemp = Words(Int(mNum / 100000)) + " Lakhs"
                    mNum = mNum Mod 100000
                    If mNum > 0 Then
                        mAndFlag = True

                        mTemp = mTemp + Words(mNum)
                    End If

                '---------------------------------------
           
''''                 ElseIf mNum < 100000 And mNum > 9999 Then
''''                    mAndFlag = False
''''                    mTemp = mTemp + Words(Int(mNum / 10000)) + " Ten Thousand"
''''                    mNum = mNum Mod 10000
''''                    If mNum > 0 Then
''''                        mAndFlag = True
''''                        mTemp = mTemp + " " + Words(mNum)
''''                    End If
''''
''''                '---------------------------------------
''''                ElseIf mNum < 10000 And mNum > 999 Then
''''                    mAndFlag = False
''''                    mTemp = mTemp + Words(Int(mNum / 1000)) + " Thousand"
''''                    mNum = mNum Mod 1000
''''                    If mNum > 0 Then
''''                        mAndFlag = True
''''                        mTemp = mTemp + " " + Words(mNum)
''''                    End If
                '-------------------------------------------
                '---------------------------------------
                ElseIf mNum < 100000 And mNum > 999 Then
                    mAndFlag = False
                    mTemp = mTemp + Words(Int(mNum / 1000)) + " Thousand"
                    mNum = mNum Mod 1000
                    If mNum > 0 Then
                        mAndFlag = True
                        mTemp = mTemp + " " + Words(mNum)
                    End If
                '-------------------------------------------
                
                ElseIf mNum < 1000 And mNum > 99 Then
                    
                    mAndFlag = False
                    mTemp = mTemp + Words(Int(mNum / 100)) + " Hundred"
                    mNum = mNum Mod 100
                    If mNum > 0 Then
                        mTemp = mTemp + " " + "and" + Words(mNum)
                    End If
                    
                ElseIf (mNum < 100 And mNum > 19) Or mNum = 10 Then
                    
                    If mAndFlag Then
                        mTemp = mTemp + " And " + TENS(Int(mNum / 10))
                        mAndFlag = False
                    Else
                        mTemp = mTemp + " " + TENS(Int(mNum / 10))
                    End If
                    
                    mNum = mNum Mod 10
                    If mNum > 0 Then
                        mTemp = mTemp + " " + Words(mNum)
                    End If
                    
                ElseIf mNum < 20 And mNum > 10 Then
                    If mAndFlag Then
                        mTemp = mTemp + " And " + TEENS(mNum)
                        mAndFlag = False
                    Else
                        mTemp = mTemp + " " + TEENS(mNum)
                    End If
                    mNum = mNum Mod 10
                ElseIf mNum < 10 And mNum > 0 Then
                    If mAndFlag Then
                        mTemp = mTemp + " And " + ONES(mNum)
                        mAndFlag = False
                    Else
                        mTemp = mTemp + " " + ONES(mNum)
                    End If
                End If
                Words = mTemp
    End Function
        
    Function Token$(tmp$, Search$)
            Dim X As Long
            X = InStr(1, tmp$, Search$)
            If X Then
               Token$ = mID$(tmp$, 1, X - 1)
               tmp$ = mID$(tmp$, X + 1)
            Else
               Token$ = tmp$
               tmp$ = ""
            End If
    End Function
        
    Function TokenCrop$(tmp$, Search$)
            Dim X As Long
            X = InStr(1, tmp$, Search$)
            If X Then
               TokenCrop$ = mID$(tmp$, 1, X - 1)
               tmp$ = Trim(mID$(tmp$, Len(Search$) + X))
            Else
               TokenCrop$ = tmp$
               tmp$ = ""
            End If
    End Function
        
    Public Sub GetUserInfo()
            Dim mRetVal As Long
            Dim mUser As String * 255
            Dim mComp As String * 255
            mRetVal = GetUserName(mUser, 255)
            mRetVal = GetComputerName(mComp, 255)
            
            
            gbUserName = Trim(mUser)
            gbComputerName = Trim(mComp)
            
            mRetVal = InStr(gbUserName, Chr(0))
            gbUserName = Left(gbUserName, mRetVal - 1)
            
            mRetVal = InStr(gbComputerName, Chr(0))
            gbComputerName = Left(gbComputerName, mRetVal - 1)
            
            If Len(gbUserName) < 1 Then gbUserName = "UnKnown_User"
            If Len(gbComputerName) < 1 Then gbComputerName = "UnKnown_Comp"
    End Sub
    
    Public Function PopulateList(objLIST As Object, mSql As String, _
        Optional sActiveIndex As String, _
        Optional bBlankLine As Boolean = False, _
        Optional bClearList As Boolean = True, _
        Optional bItemData As Boolean = False, _
        Optional mConnectTo As enuSourceString = 1) As Boolean
        
        '***********************************************************************************'
        ' Writen : Aiby                                                                     '
        ' Date   : 15-Apr-2000  Modified : 26-May-2003                                      '
        '                                                                                   '
        ' Description:-                                                                     '
        '   objLIST -> should be Combo or List Box                                          '
        '   mSQL    -> SQL Statement                                                        '
        '   mActiveIndex ->  to set to any particular List item after populating the Obj    '
        '                ->  It can be either a List Item or a List Index                   '
        '   mBlankLine   ->  Inserts a blank line in Index 0                                '
        '   mClearLisgt  ->  If TRUE -> it clears the Object befor start filling it         '
        '   mItemData    ->  If TRUE -> ItemDate property will be filled                    '
        '***********************************************************************************'
            
            Dim Rec As New ADODB.Recordset
            Dim mAddress As Long
            Dim iActiveIndex As Long
            Dim mListIndex As Long
            Dim mTemp As String
            
            '-------------------------------------'
            '
            '-------------------------------------'
                Dim objdb As New clsDB
                Dim mCnn As New ADODB.Connection
                
                
                Select Case mConnectTo
                    Case Is = 1
                        objdb.SetConnection mCnn
                    Case Else
                        If Not objdb.CreateNewConnection(mCnn, mConnectTo) Then
                            Debug.Print " Error: Function call PopulateList"
                            Debug.Print "     : Attempt to connect to external source"
                            Exit Function
                        End If
                End Select
            '-------------------------------------'
            
            'Check which type of item we are dealing with
            Select Case TypeName(objLIST)
                Case "ComboBox"
                    mAddress = CB_ADDSTRING
                Case "ListBox"
                    mAddress = LB_ADDSTRING
                Case Else
                    'Incorrect item has been passed to this routine
                    Exit Function
            End Select
            
            'Make sure mandatory items have been entered
            Rec.CursorLocation = adUseClient
            Rec.Open mSql, mCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Rec.RecordCount < 0 Then Exit Function
            
            'Check if we have been passed an Active Index
            If IsNumeric(sActiveIndex) = True Then
                'A number was passed so set this to the Active Index
                iActiveIndex = CInt(sActiveIndex)
                sActiveIndex = ""
            Else
                'No number was passed so for now don't set an Active Index
                iActiveIndex = -1
            End If
            
            
            'When set to True, clear the list
            If bClearList = True Then objLIST.Clear
            
            'Check if a blank line is required in the list
            If bBlankLine = True And mListIndex = 0 Then
                If TypeName(objLIST) = "ComboBox" And objLIST.Style = 2 Then
                    SendMyMessage objLIST.hwnd, mAddress, 0, " "
                    mListIndex = 1
                Else
                    SendMyMessage objLIST.hwnd, mAddress, 0, ""
                    mListIndex = 1
                End If
                
            End If
            
            
            While Not Rec.EOF
                    
                    mTemp = Rec.Fields(0).Value
                    SendMyMessage objLIST.hwnd, mAddress, 0, mTemp
                    
                    If bItemData Then
                        If IsNumeric(Rec.Fields(1).Value) Then
                            objLIST.ItemData(mListIndex) = Rec.Fields(1).Value
                        Else
                            objLIST.ItemData(mListIndex) = -1
                        End If
                    End If
                    
                    'Check if the current Item matches the Active Index parameter
                    If sActiveIndex <> "" And sActiveIndex = Rec.Fields(0).Value Then
                        'Set this item to be the Active Index
                        If bBlankLine Then
                            iActiveIndex = mListIndex
                        Else
                            iActiveIndex = mListIndex - 1
                        End If
                    End If
                    mListIndex = mListIndex + 1
                    Rec.MoveNext
            Wend
            
            'Set the Active Index for the item
            On Error Resume Next
            objLIST.ListIndex = iActiveIndex
            
            'Return the List Count
            PopulateList = True
    
    End Function
    
    Function FillFlexGridCombo(fg As VSFlexGrid, ByVal cboCol As Integer, ByVal SQL As String, _
                    ByVal ADOCmd As ADODB.CommandTypeEnum, ByVal vDsn As Integer)
        '----------------------------'
        ' Fill vsFlexGridCombo       '
        '----------------------------'
        Dim AdoCon As New ADODB.Connection
        Dim arOt, Cnt, lCnt As Integer
        Dim items As String
        Dim objdb As New clsDB
        
        
        If (gbSoochikaVer <> 5) Then
            objdb.CreateNewConnection AdoCon, enuSourceString.SOOCHIKA
            
            ExecuteSP SQL, rselect, ADOCmd, , arOt, AdoCon
        
            items = "#0;..."
            If IsArray(arOt) Then
                For Cnt = 0 To UBound(arOt, 2)
                    items = items & "|"
                    items = items & "#" & arOt(0, Cnt) & ";" & arOt(1, Cnt)
                Next Cnt
            End If
            
            fg.ColComboList(cboCol) = items
            AdoCon.Close
            
        Else
            objdb.CreateNewConnection AdoCon, enuSourceString.SoochikaUnicode
            ExecuteSP SQL, rselect, ADOCmd, , arOt, AdoCon
        
            items = "#0;..."
            If IsArray(arOt) Then
                For Cnt = 0 To UBound(arOt, 2)
                    items = items & "|"
                    items = items & "#" & arOt(1, Cnt) & ";" & arOt(0, Cnt)
                Next Cnt
            End If
            
            fg.ColComboList(cboCol) = items
            AdoCon.Close
        End If
    End Function
    
    Sub gFillVSGrid(vsG As VSFlexGrid, mCol As Integer, mSql As String, Optional cnnStr As enuSourceString)
        Dim objdb As New clsDB
        Dim mCn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mItem As String
        If IsMissing(cnnStr) Then
            objdb.SetConnection mCn
        Else
            objdb.CreateNewConnection mCn, cnnStr
        End If
        mItem = "#0;..."
        Rec.Open mSql, mCn, adOpenForwardOnly, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                mItem = mItem & "|#" & Rec.Fields(0).Value & ";" & Rec.Fields(2).Value
                Rec.MoveNext
            Wend
        End If
        Rec.Close
        mCn.Close
        Set mCn = Nothing
        
        vsG.ColComboList(mCol) = mItem
        
    End Sub
    
    Public Function GetRecordSet(mSource As String, _
                                Optional mCursorType As CursorTypeEnum = adOpenStatic, _
                                Optional mLockType As LockTypeEnum = adLockReadOnly, _
                                Optional adoCn As ADODB.Connection _
                                ) As ADODB.Recordset
        
    '-------------------------------------------------------------------'
    ' Aiby : 1-Jun,2003                                                 '
    ' Desc :                                                            '
    '        This Functions was intend to fetch Recordset from Table by '
    '        Stored procedures or Table Name or Dynamic Queries with    '
    '        Parameters.                                                '
    '        By the this will also permit Transactiona Queries..        '
    ' Ref  :
    '        Decleare ADODB Recordset
    '      1).  Set Rec = GetRecordSet("Select * From  Product Where Category = 1")
    '      2).  Set Rec = GetRecordset("TableName")
    '      3).  set Rec = GetRecordset("StoredProcedureName")
    '      4).  set Rec = GetRecordset("StoredProcedureName 'StringPara1', IntPara, '1-Jan-2002';")
    '      And much more.. sorry i am not in a mood to document much right now!
    '      Best of Luck!  Warning : No Gurenty! Test it out and find it..!
    '-------------------------------------------------------------------'
        Dim adoCmdType          As Integer
        Dim adoLocalCnn         As ADODB.Connection
        Dim Recs                As New ADODB.Recordset
        Dim objdb As New clsDB
        
        If adoCn Is Nothing Then
            objdb.SetConnection adoLocalCnn
        Else
            Set adoLocalCnn = adoCn
        End If
        
        'On Error GoTo Errhandler
        Recs.Open mSource, adoLocalCnn, mCursorType, mLockType, adoCmdType
        Set GetRecordSet = Recs
        Exit Function
ErrHandler:
        Set GetRecordSet = Nothing
    End Function
    
    Public Function ExecuteSQL(mSource As String, Optional adoCnn As Connection = Nothing) As Boolean
    '-------------------------------------------------------------------'
    ' Aiby      : 1-Jun,2003                                            '
    ' Desc      : Derived from ExecuteSp                                '
    ' Status    : Testing                                               '
    ' Objective : To execute Dynamic SQL statements                     '
    '-------------------------------------------------------------------'
        Dim ADOCmd              As New ADODB.Command
        Dim adoLocalCnn As New ADODB.Connection
        'Set adoLocalCnn = mCnn
        'On Error GoTo Errhandler
        
        Dim mCnn As New ADODB.Connection '
        
        Set ADOCmd.ActiveConnection = mCnn
        ADOCmd.CommandType = adCmdText
        ADOCmd.CommandText = mSource
        ADOCmd.Execute
        
        
        ExecuteSQL = 1
        Set ADOCmd = Nothing
        Set adoLocalCnn = Nothing
        Exit Function
        
ErrHandler:
        ExecuteSQL = 0
    End Function
    
    Public Sub subExecuteUpdate(strStoredProcedure As String, Optional varInPut As Variant, Optional adoCnn As Connection = Nothing)
        '======================================================================='
        ' Aiby  : 15-Jun,2003                                                   '
        '       : Derived from Execute SP                                       '
        '       : This Procedure is writen for pass Input values to any Stored  '
        '         Procedure, especially for Update and Insert Procedures        '
        '======================================================================='
    
        Dim adoLocalConnection As ADODB.Connection
        Dim adoCommand As New ADODB.Command
        Dim adoParameters As ADODB.Parameters
        Dim mNewConnectionsFlag As Boolean
        Dim intCount As Integer
        Dim mCnn As ADODB.Connection
        
        On Error GoTo SKIP2:
        
        If adoCnn Is Nothing Then
            Set adoLocalConnection = mCnn
            On Error GoTo Skip:
            If adoLocalConnection Is Nothing Then
                On Error GoTo SKIP2:
                'Set adoLocalConnection = SetUpConnection
                On Error GoTo Skip:
            End If
            adoLocalConnection.BeginTrans
            On Error GoTo EHandler
        Else
            Set adoLocalConnection = adoCnn
            On Error GoTo Skip:
        End If
        
        With adoCommand
            Set .ActiveConnection = adoLocalConnection
            .CommandType = adCmdStoredProc
            .CommandText = strStoredProcedure
            Set adoParameters = .Parameters
        End With
        
        If Not IsMissing(varInPut) Then
            For intCount = 0 To UBound(varInPut)
                adoParameters(intCount + 1).Value = varInPut(intCount)
                'Debug.Print adoParameters(intCount + 1).Name & vbTab & adoParameters(intCount + 1).Value
            Next
        Else
            GoTo Skip:
        End If
        
        adoCommand.Execute
        adoLocalConnection.CommitTrans
        
        GoTo Skip:
EHandler:
        If mNewConnectionsFlag Then
            adoLocalConnection.RollbackTrans
        End If
Skip:
        Set adoLocalConnection = Nothing
SKIP2:
    End Sub
    
    Public Function ExecuteSP(ByVal strForExecute As String, _
                                ByVal QryType As Integer, _
                                ByVal ADOCmd As ADODB.CommandTypeEnum, _
                                Optional vAryIn, _
                                Optional varyOut, _
                                Optional adoConnection As ADODB.Connection)
                                
        Dim AdoCon As New ADODB.Connection
        Dim adocom As New ADODB.Command
        Dim adoRec As New ADODB.Recordset
        Dim intcnt As Integer
        Dim lpCnt1, lpCnt2 As Integer
        
        If Not IsMissing(adoConnection) Then
            Set adocom.ActiveConnection = adoConnection
        Else: Set adocom.ActiveConnection = gFunSetConnection(Dsn.dsnFA)
        
        End If
        If Not IsMissing(vAryIn) Then
            For intcnt = 0 To UBound(vAryIn)
                If vAryIn(intcnt) = "" Or IsEmpty(vAryIn(intcnt)) Then vAryIn(intcnt) = Null
            Next intcnt
        End If
        
        adocom.CommandType = ADOCmd
        adocom.CommandText = strForExecute
        Select Case QryType
            Case RInsert
                Set adoRec = adocom.Execute(, vAryIn)
                If Not IsMissing(varyOut) Then
                    If (adoRec.BOF = False And adoRec.EOF = False) Then
                        varyOut = adoRec.GetRows()
                    End If
                End If
            Case rselect
                If IsMissing(vAryIn) Then
                    Set adoRec = adocom.Execute
                Else
                    Set adoRec = adocom.Execute(, vAryIn)
                End If
                If (adoRec.BOF = False And adoRec.EOF = False) Then varyOut = adoRec.GetRows()
                If IsArray(varyOut) Then
                    For lpCnt1 = 0 To UBound(varyOut)
                        For lpCnt2 = 0 To UBound(varyOut, 2)
                            If IsNull(varyOut(lpCnt1, lpCnt2)) Then varyOut(lpCnt1, lpCnt2) = ""
                        Next lpCnt2
                    Next lpCnt1
                End If
            Case rUpdate
                If IsMissing(vAryIn) Then Set adoRec = adocom.Execute Else Set adoRec = adocom.Execute(, vAryIn)
            Case RDelete
                If IsMissing(vAryIn) Then adocom.Execute Else adocom.Execute , vAryIn
        End Select
Exitfunction:
        Set adocom = Nothing
    End Function
           
    Public Sub gSubSetComboItem(Cmb As ComboBox, Value As Variant)
        'Setting the selected Combo Item
        Dim i As Integer
            If Not Value = "" Then
               For i = 1 To Cmb.ListCount - 1
                    If Cmb.ItemData(i) = Value Then
                        Cmb.ListIndex = i
                        Exit For
                    End If
               Next i
            Else
               Cmb.ListIndex = 0
            End If
    End Sub
            
    Public Sub AccrualJournalByDemandID(mDemandID As Variant)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim RecIDemand As New ADODB.Recordset
        Dim mSql As String
        Dim RecTran As New ADODB.Recordset
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mTotalAmount As Variant
                
        Dim intTransactionID                As Variant
        Dim intLocalBodyID                  As Variant
        Dim intFinancialYearID              As Variant
        Dim dtTransactionDate               As Variant
        Dim intExternalApplicationID        As Variant
        Dim intExternalApplicationModuleID  As Variant
        Dim intFunctionID                   As Variant
        Dim intFunctionaryID                As Variant
        Dim intFieldID                      As Variant
        Dim intFundID                       As Variant
        Dim intBudgetCentreID               As Variant
        Dim vchNarration                    As Variant
        Dim intTransactionTypeID            As Variant
        Dim intProcessID                    As Variant
        Dim vchGroup                        As Variant
        Dim intGroupID                      As Variant
        Dim intKeyID                        As Variant
        Dim numSubLedgerID                  As Variant
        Dim numUserID                       As Variant
        Dim intVoucherNo                    As Variant
                
        'Dim intTransactionID                As Variant
        Dim intSerialNo                     As Variant
        Dim intAccountHeadID                As Variant
        Dim fltAmount                       As Variant
        Dim tinDebitOrCreditFlag            As Variant
        Dim intByAccountHeadID              As Variant
        'Dim vchNarration                    As Variant
        'Dim intFundID                       As Variant
        
        
        mSql = " Select * From faIDemandTbl Inner Join "
        mSql = mSql + " faIDemandChild ON faIDemandChild.numDemandID = faIDemandTbl.numDemandID Inner Join"
        mSql = mSql + " faTransactionType On faTransactionType.intTransactionTypeID = faIDemandTbl.intTransactionTypeID"
        mSql = mSql + " Where tnyAccrualType = 1 And faIDemandTbl.numDemandID = " & mDemandID
        
        objdb.SetConnection mCnn
        RecIDemand.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If RecIDemand.BOF And RecIDemand.EOF Then
            MsgBox "There is no Demand Found for Proceed", vbInformation
            Exit Sub
        End If
        intTransactionID = -1
        intLocalBodyID = gbLocalBodyID
        intFinancialYearID = gbFinancialYearID
        dtTransactionDate = RecIDemand!dtDueDate
        intExternalApplicationID = 115
        intExternalApplicationModuleID = 0
        intFunctionID = RecIDemand!intFunctionID
        intFunctionaryID = RecIDemand!intFunctionaryID
        intFieldID = Null
        intFundID = gbFundID
        intBudgetCentreID = Null
        vchNarration = RecIDemand!vchRemarks
        intTransactionTypeID = RecIDemand!intTransactionTypeID
        intProcessID = Null
        vchGroup = "JV"
        intGroupID = 40
        intKeyID = Null
        numSubLedgerID = RecIDemand!numDemandID
        numUserID = gbUserID
        intVoucherNo = Null
        
        arrInput = Array( _
        intTransactionID, _
        intLocalBodyID, _
        intFinancialYearID, _
        dtTransactionDate, _
        intExternalApplicationID, _
        intExternalApplicationModuleID, _
        intFunctionID, _
        intFunctionaryID, _
        intFieldID, _
        intFundID, _
        intBudgetCentreID, _
        vchNarration, _
        intTransactionTypeID, _
        intProcessID, _
        vchGroup, _
        intGroupID, _
        intKeyID, _
        numSubLedgerID, _
        numUserID, _
        intVoucherNo)
        
        objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut
        If IsArray(arrOutPut) Then
            intTransactionID = arrOutPut(0, 0)
        End If
        intSerialNo = 1
        
        mSql = " Select * From faTransactionType INNER JOIN "
        mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intTransactionTypeID = faTransactionType.intTransactionTypeID "
        mSql = mSql + " Where faTransactionType.intTransactionTypeID = " & RecIDemand!intTransactionTypeID
        mSql = mSql + " Order By intOrder"
        RecTran.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        If Not (RecTran.BOF And RecTran.EOF) Then
            intByAccountHeadID = RecTran!intAccountHeadID
            While Not RecIDemand.EOF
                RecTran.MoveFirst
                While Not RecTran.EOF
                    Debug.Print RecTran!intOrder
                    If RecTran!intAccountHeadID = RecIDemand!intAccountHeadID Then
                            intSerialNo = intSerialNo + 1
                            intAccountHeadID = RecIDemand!intAccountHeadID
                            fltAmount = RecIDemand!fltAmount
                            tinDebitOrCreditFlag = RecTran!tinDebitOrCredit
                            'intByAccountHeadID
                            vchNarration = Null
                            intFundID = gbFundID
                            mTotalAmount = mTotalAmount + RecIDemand!fltAmount
                            
                            arrInput = Array( _
                            intTransactionID, _
                            intSerialNo, _
                            intAccountHeadID, _
                            fltAmount, _
                            tinDebitOrCreditFlag, _
                            intByAccountHeadID, _
                            vchNarration, _
                            intFundID)
                            
                            objdb.ExecuteSP "spSaveTransactionChild", arrInput
                            GoTo SkipLoop:
                    End If
                    RecTran.MoveNext
                Wend
SkipLoop:
                RecIDemand.MoveNext
            Wend
                
            RecTran.MoveFirst
            intSerialNo = 1
            intAccountHeadID = RecTran!intAccountHeadID
            fltAmount = mTotalAmount
            tinDebitOrCreditFlag = RecTran!tinDebitOrCredit
            intByAccountHeadID = Null
            'vchNarration = RecIDemand!vchNarration
            intFundID = gbFundID
            
            arrInput = Array( _
            intTransactionID, _
            intSerialNo, _
            intAccountHeadID, _
            fltAmount, _
            tinDebitOrCreditFlag, _
            intByAccountHeadID, _
            vchNarration, _
            intFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput
            
        End If 'If Not (RecTran.BOF And RecTran.EOF) Then
        'mSQL = "Update faIDemandTbl Set tnyStatus = 1 Where numDemandID = " & mDemandID
        'mCnn.Execute mSQL
        RecIDemand.Close
        RecTran.Close
    End Sub
    
    Public Sub PrintDemandSlip(mDemandID As Variant, mCnn As ADODB.Connection)
        Dim objdb As New clsDB
        'Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim RecAddress As New ADODB.Recordset
        Dim arrInput As Variant
        Dim objTranType As New clsTransactionType
        Dim mSql As String
        Dim mTotalAmt As Double
        
        arrInput = Array(mDemandID)
        'objDB.SetConnection mCnn
        
        mSql = "        Select faIDemandTbl.*, vchSectionName From faIDemandTbl Inner Join "
        mSql = mSql + " faSection On faSection.intSectionID = faIDemandTbl.intSectionID"
        mSql = mSql + " Where numDemandID = " & mDemandID
        
        Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            
            'Call FileInitialize
            Call PrinterInit
                On Error Resume Next
                Print #gbFileNO,
                Print #gbFileNO, Style(gbTitle1, True, True)
                Print #gbFileNO, Style("  Demand Slip", True, True)
                Print #gbFileNO,
                Print #gbFileNO, "Demand No:"; Rec!vchDemandNo
                Print #gbFileNO, "Demand Date : "; DdMmmYy(Rec!dtDemandDate)
                
                objTranType.SetTransactionType Rec!intTransactionTypeID
                Print #gbFileNO, Rec!vchSectionName
                
                If objTranType.TransactionTypeID > 0 Then
                    Print #gbFileNO, objTranType.TransactionType
                Else
                    Print #gbFileNO, "Transaction Type : Unknown < Please Contact System Administrator"
                End If
                
                '----------------------------------------------------------'
                ' iDemandChild recordset is only required here if One have
                ' to Print the head wise details.
                '----------------------------------------------------------'
                mSql = "Select * From faIDemandChild Where numDemandID = " & mDemandID
                RecChild.Open mSql, mCnn, adOpenStatic, adLockOptimistic
                While Not RecChild.EOF
                    mTotalAmt = mTotalAmt + RecChild!fltAmount
                    RecChild.MoveNext
                Wend
                
                Print #gbFileNO, "Amount : " & mTotalAmt
                mSql = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
                RecAddress.Open mSql, mCnn, adOpenStatic, adLockOptimistic
                If Not (RecAddress.BOF And RecAddress.EOF) Then
                    Print #gbFileNO, "Ward    : " & RecAddress!intWardNo; "       ";
                    Print #gbFileNO, "Door No : " & RecAddress!intDoorNo & IIf(Len(RecAddress!vchDoorNo2), "/" & RecAddress!vchDoorNo2, "")
                    Print #gbFileNO, "Name    : " & RecAddress!vchName
                    Print #gbFileNO, "Phone   : " & RecAddress!vchPhone
                End If
                
                Print #gbFileNO,
                Print #gbFileNO, Rec!vchRemarks
                Print #gbFileNO,
                Print #gbFileNO, "Prepared By " & gbUserName & "     Seat Name " & gbSeatName
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                
            Close #gbFileNO
            'Shell "Print " & gbFileName
            'ShellPad
            End If
            Rec.Close
        
'        Set mCnn = Nothing
'        Set objDB = Nothing
        
    End Sub
    
    Public Sub MenuManager()
        '    Dim mCrl As Control
        '    For Each mCrl In Me.Controls
        '    If TypeOf mCrl Is TextBox Then
    
    
        Dim mCrl As Control
        Dim mCrl2 As Control
        FileInitialize
        For Each mCrl In frmMenu.Controls
            If TypeOf mCrl Is Menu Then
                Set mCrl2 = mCrl
                On Error Resume Next
                'If mCrl.Name <> "Administration" Then
                    mCrl2.Visible = False
                'End If
                Print #gbFileNO, mCrl.Name
                GoTo NextLine:
                
            End If
NextLine:
        Next
        Close #gbFileNO
        ShellPad
        
        
'
' frmMenu.Administration
'    frmMenu.AccountHeads
'    frmMenu.SubsidiaryAccount
'    frmMenu.FunctionaryHead
'    frmMenu.Bank
'    frmMenu.ChequeBook
'    frmMenu.Funds
'    frmMenu.Functions
'    frmMenu.Functionaries
'    frmMenu.Fields
'    frmMenu.CollectionRegister
'    frmMenu.StockRegisterReceiptBooks
'    frmMenu.DefineReportSchedules
'    frmMenu.BudgetCentres
'    frmMenu.BudgetAllocation
'    frmMenu.BudgetRevision
'    frmMenu.OpeningBalanceSheet
'    frmMenu.UpdateOpeningBalance
'    frmMenu.User
'    frmMenu.TransactionType
'    frmMenu.CounterList
'    frmMenu.ConfigurationSettings
'    frmMenu.LocalBodySettings
'frmMenu.Transactions
'    frmMenu.Receipts
'    frmMenu.CounterReceipt
'    frmMenu.Payments
'    frmMenu.PaymentOrder
'    frmMenu.JournalEntry
'    frmMenu.ContraEntry
'    frmMenu.BankReconciliationEntry
'    frmMenu.BankReconcile
'    frmMenu.AllotmentLetter
'    frmMenu.SearchPaymentOrder
'frmMenu.Utilities
'    frmMenu.DemandInterface
'    frmMenu.DemandRegister
'    frmMenu.RecieptCancellation
'    frmMenu.AccrualDemand
'    frmMenu.InwardChecksAndDds
'    frmMenu.SearchBuildingTaxRemitance
'    frmMenu.SearchReceipts
'    frmMenu.VoucherUtility
'    frmMenu.ChangePassword
'frmMenu.Reports
'    frmMenu.DailyReports
'    frmMenu.CounterwiseDetails
'    frmMenu.Chitta
'    frmMenu.CancelledReceipts
'    frmMenu.HeadwiseConsolidation
'    frmMenu.DayBookReceipts
'    frmMenu.rptCashBook
'    frmMenu.rptBankBook
'    frmMenu.rptJournalBook
'    frmMenu.rptLedgerBook
'    frmMenu.rptTrialBalance
'    frmMenu.rptBalanceSheet
'    frmMenu.rptIncomeAndExpenditure
'    frmMenu.rptReceiptsAndPayments
'    frmMenu.rptBudgetVariance
'    frmMenu.BankReconciliation
'    frmMenu.ChequeRegister
'    frmMenu.ChequeIssue
'    frmMenu.ChequeReceived
'    frmMenu.rptRegisters
'    frmMenu.rptAppropriationControlRegister
'    frmMenu.rptAssetReplacementRegister
'    frmMenu.rptAuthorisationIssuetoSecretary
'    frmMenu.rptBillofReceiptsRegister
'    frmMenu.rptCollectionregister
'    frmMenu.rptDemandregister
'    frmMenu.rptDepositreceivedregister
'    frmMenu.rptDocumentcontrolRegister
'    frmMenu.rptFormGEN40Register
'    frmMenu.rptFunctionWiseExpenditure
'    frmMenu.rptFunctionwisereceiptsubsidiaryledger
'    frmMenu.rptFundsReceivedRegister
'    frmMenu.rptImmovablePropertyRegister
'    frmMenu.rptImplentingOfficerwiseAllotmentRegister
'    frmMenu.rptIncomeandExpenditureRegister
'    frmMenu.rptLandRegister
'    frmMenu.rptLetterofallotment
'    frmMenu.rptMemorandumofcollectionRegister
'    frmMenu.rptMovablepropertyRegister
'    frmMenu.rptOfficialReceiptRegister
'    frmMenu.rptPaymentOrderRegister
'    frmMenu.rptProjectregister
'    frmMenu.rptRegisterofadvances
'    frmMenu.rptRegisterofbillsforpayment
'    frmMenu.rptRegisterofPermenantadvance
'    frmMenu.rptRegisterofpubliclightingsystem
'    frmMenu.rptRequesitionforReleaseofFundcodes
'    frmMenu.rptStatementofOutstandingLiabilityforexpenses
'    frmMenu.rptStatementonStatusofChequereceived
'    frmMenu.rptSubsidiaryRegister
'    frmMenu.rptSummaryofCollectionRegister
'    frmMenu.rptSummaryStatementodfbills
'    frmMenu.rptSummaryStatementofDeposits
'    frmMenu.rptSummaryStatementofRefundandRemission
'    frmMenu.rptSummaryStatementofWriteoffs
'    frmMenu.TestReport
'    frmMenu.Exit
'    frmMenu.PropertyTax
'    frmMenu.Test
'    frmMenu.rptLedgerView
'    frmMenu.ReportGenerator
'    frmMenu.AccountHeadsNew
'    frmMenu.SearchBuilding
'    frmMenu.TransactionTemp
'    frmMenu.DeleteTransactionEntry
'    frmMenu.ProfessionTax
'    frmMenu.Logoff
'    frmMenu.LogOut

    End Sub

Function Column(Cols As Integer, Optional S1 As String, Optional W1 As Integer, Optional S2 As String, Optional W2 As Integer, Optional S3 As String, Optional W3 As Integer, Optional S4 As String, Optional W4 As Integer, Optional S5 As String, Optional W5 As Integer, Optional S6 As String, Optional W6 As Integer, Optional S7 As String, Optional W7 As Integer) As String
    '=======================================
    ' Written by : Aiby     Date : 01-07-99
    '=======================================
    Dim temp As String
    Dim X As Integer
    Dim Ch As String
    
    gbColLines = 0
    If IsMissing(W1) Then W1 = 0
    If IsMissing(W2) Then W2 = 0
    If IsMissing(W3) Then W3 = 0
    If IsMissing(W4) Then W4 = 0
    If IsMissing(W5) Then W5 = 0
    If IsMissing(W6) Then W6 = 0
    If IsMissing(W6) Then W7 = 0
    'If (W1 + W2 + W3 + W4 + W5 + W6) - (Cols - 1) * 2 > gbTextArea Then
    '    MsgBox "Error in UDF=>Column : Width > TextArea", vbCritical
    '    Exit Function
    'End If
    
    If IsMissing(S1) Then S1 = ""
    If IsMissing(S2) Then S2 = ""
    If IsMissing(S3) Then S3 = ""
    If IsMissing(S4) Then S4 = ""
    If IsMissing(S5) Then S5 = ""
    If IsMissing(S6) Then S6 = ""
    If IsMissing(S6) Then S7 = ""
    
    While (True)
    
        temp = ""
        If W1 > 0 And Len(S1) = 0 Then
            Column = Column + Space(gbLeftMargin) + Space(W1) + Space(2)
        Else
        For X = 1 To Len(S1)
            Ch = mID(S1, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                If Len(temp) < W1 Then
                temp = temp + mID(S1, X, 1)
                Else
                Column = Column + Space(gbLeftMargin) + temp + Space(2)
                S1 = Right(S1, Len(S1) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + Space(gbLeftMargin) + temp + Space(2)
                S1 = Right(S1, Len(S1) - Len(temp) - 2)
                Exit For
            End If
            If temp = S1 Then
                temp = temp + Space(W1 - Len(temp))
                Column = Column + Space(gbLeftMargin) + temp + Space(2)
                S1 = ""
                Exit For
            End If
        Next X
        End If
            
        temp = ""
        If W2 > 0 And Len(S2) = 0 Then
            Column = Column + Space(W2) + Space(2)
        Else
        For X = 1 To Len(S2)
            Ch = mID(S2, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                If Len(temp) < W2 Then
                temp = temp + mID(S2, X, 1)
                Else
                Column = Column + temp + Space(2)
                S2 = Right(S2, Len(S2) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + temp + Space(2)
                S2 = Right(S2, Len(S2) - Len(temp) - 2)
                Exit For
            End If
            If temp = S2 Then
                temp = temp + Space(W2 - Len(temp))
                Column = Column + temp + Space(2)
                S2 = ""
                Exit For
            End If
        Next X
        End If
        
        
        temp = ""
        If W3 > 0 And Len(S3) = 0 Then
            Column = Column + Space(W3) + Space(2)
        Else
        For X = 1 To Len(S3)
            Ch = mID(S3, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                'If X = 19 Then Stop
                If Len(temp) < W3 Then
                temp = temp + mID(S3, X, 1)
                Else
                Column = Column + temp + Space(2)
                S3 = Right(S3, Len(S3) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + temp + Space(2)
                S3 = Right(S3, Len(S3) - Len(temp) - 2)
                Exit For
            End If
            If temp = S3 Then
                temp = temp + Space(W3 - Len(temp))
                Column = Column + temp + Space(2)
                S3 = ""
                Exit For
            End If
        Next X
        End If
        
            
        temp = ""
        If W4 > 0 And Len(S4) = 0 Then
            Column = Column + Space(W4) + Space(2)
        Else
        For X = 1 To Len(S4)
            Ch = mID(S4, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                If Len(temp) < W4 Then
                temp = temp + mID(S4, X, 1)
                Else
                Column = Column + temp + Space(2)
                S4 = Right(S4, Len(S4) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + temp + Space(2)
                S4 = Right(S4, Len(S4) - Len(temp) - 2)
                Exit For
            End If
            If temp = S4 Then
                temp = temp + Space(W4 - Len(temp))
                Column = Column + temp + Space(2)
                S4 = ""
                Exit For
            End If
        Next X
        End If
        
        temp = ""
        If W5 > 0 And Len(S5) = 0 Then
            Column = Column + Space(W5) + Space(2)
        Else
        For X = 1 To Len(S5)
            Ch = mID(S5, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                If Len(temp) < W5 Then
                temp = temp + mID(S5, X, 1)
                Else
                Column = Column + temp + Space(2)
                S5 = Right(S5, Len(S5) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + temp + Space(2)
                S5 = Right(S5, Len(S5) - Len(temp) - 2)
                Exit For
            End If
            If temp = S5 Then
                temp = temp + Space(W5 - Len(temp))
                Column = Column + temp + Space(2)
                S5 = ""
                Exit For
            End If
        Next X
        End If
        
        
        temp = ""
        If W6 > 0 And Len(S6) = 0 Then
            Column = Column + Space(W6) + Space(2)
        Else
        For X = 1 To Len(S6)
            Ch = mID(S6, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                If Len(temp) < W6 Then
                temp = temp + mID(S6, X, 1)
                Else
                Column = Column + temp + Space(2)
                S6 = Right(S6, Len(S6) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + temp + Space(2)
                S6 = Right(S6, Len(S6) - Len(temp) - 2)
                Exit For
            End If
            If temp = S6 Then
                temp = temp + Space(W6 - Len(temp))
                Column = Column + temp + Space(2)
                S6 = ""
                Exit For
            End If
        Next X
        End If
        
       
        temp = ""
        If W7 > 0 And Len(S7) = 0 Then
            Column = Column + Space(W7) + Space(2)
        Else
        For X = 1 To Len(S7)
            Ch = mID(S7, X, 1)
            'If Ch <> Chr(13) Then
            If Ch <> vbCr And Ch <> vbLf Then
                If Len(temp) < W7 Then
                temp = temp + mID(S7, X, 1)
                Else
                Column = Column + temp
                S7 = Right(S7, Len(S7) - Len(temp))
                Exit For
                End If
            Else
                Column = Column + temp
                S7 = Right(S7, Len(S7) - Len(temp) - 2)
                Exit For
            End If
            If temp = S7 Then
                temp = temp + Space(W7 - Len(temp))
                Column = Column + temp
                S7 = ""
                Exit For
            End If
        Next X
        End If
        
        Column = Column + vbCrLf ' Chr$(27) + Chr$(10)
        gbColLines = gbColLines + 1
        
        If Len(S1) < 1 And Len(S2) < 1 And Len(S3) < 1 And Len(S4) < 1 And Len(S5) < 1 And Len(S6) < 1 And Len(S7) < 1 Then GoTo Skip
    Wend
Skip:
    Column = Left(Column, Len(Column) - 1)
    'If gbPaperFlag And (Right(Column, 1) = vbCr Or Right(Column, 1) = vbLf Or Right(Column, 1) = vbNewLine) Then GoTo Skip

End Function


      Public Sub PrintSummaryReceiptPTaxRes(intVoucherID As Double)

        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim strSubstr As String
        Dim strSubR As String
        
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        'FileInitialize
        Open gbFileName For Output As #gbFileNO
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        

        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic
        
        Select Case Rec!intInstrumentTypeID
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(66); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(66); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Is = 5
            Print #gbFileNO, Tab(311); gbDoubleWidth; "CHEQUE"; Tab(66); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Else
            Print #gbFileNO,
        End Select
        Print #gbFileNO,
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(17); gbBold; gbDoubleWidth; Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); Tab(62); Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbBoldOff; gbDoubleWidthOff
            Print #gbFileNO, Tab(32); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
 
            Print #gbFileNO, Tab(15); Style(mName, True); Tab(65); Style(mName, True); Tab(87); gbBold; "GSTIN : "; Tab(96); gbGSTIN;
            
            
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(65); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(65); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(65); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(65); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
            ' Line 15 Next
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(75); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            Print #gbFileNO,
           ' Print #gbFileNO,
            
            '======================================================================================='
            '  B O D Y     P A R T     O F      R E C E I P T                                       '
            '======================================================================================='
            ' Line 18 Next

            Dim mPTaxA As Variant
            Dim mPTaxC As Variant
            Dim mLCA As Variant
            Dim mLCC As Variant
            Dim mPCA As Variant
            Dim mPCC As Variant


            Dim mPenal As Variant
            Dim mRndOff As Variant
            Dim mOthers As Variant
            Dim mNarration As String
            Dim mNarra As String
            Dim mStartingYear As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear As Integer
            Dim mEndingPeriod As Integer
            mStartingYear = 2100
            
             Dim mAmtPTaxCurrent As Double
                Dim mAmtPTaxArrear As Double
                Dim mAmtLC As Double
                Dim mAmtPenal As Double
                Dim mAmtServiceCess As Double
                Dim mAmtSurcharge As Double
                Dim mSplServices As Double
                Dim mSurCentralGovtBuild As Double
                Dim mNotice As Double
                Dim mWarantee As Double
                Dim mAdvance As Double
                While Not Rec.EOF
                    Select Case Rec!vchAccountHeadCode
                        
                        Case gbAcHeadCodePropertyTaxCurrent, gbAcHeadCodePropertyTax_NonResidential_Current
                             mAmtPTaxCurrent = mAmtPTaxCurrent + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodePropertyTaxArrear, gbAcHeadCodePropertyTax_NonResidential_Arrear
                             mAmtPTaxArrear = mAmtPTaxArrear + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodeLibraryCess
                            mAmtLC = mAmtLC + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodePenalInterest
                            mAmtPenal = mAmtPenal + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodeServicceCessCurrent, gbAcHeadCodeServicceCessCurrentNonR, gbAcHeadCodeServicceCessArrear, gbAcHeadCodeServicceCessArrearNonR
                            mAmtServiceCess = mAmtServiceCess + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodeSurPTCurrent, gbAcHeadCodeSurPTCurrentNonR, gbAcHeadCodeSurPTArrear, gbAcHeadCodeSurPTArrearNonR
                            mAmtSurcharge = mAmtSurcharge + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodeSplServicesCurrent, gbAcHeadCodeSplServicesArrear
                            mSplServices = mSplServices + Format(Rec!fltAmount, "0.00")
                        Case gbAcHeadCodeSurCentralGovtBuildCurrent, gbAcHeadCodeSurCentralGovtBuildArrear
                            mSurCentralGovtBuild = mSurCentralGovtBuild + Format(Rec!fltAmount, "0.00")
                        Case 140400200          '''Notice fee
                            mNotice = mNotice + Format(Rec!fltAmount, "0.00")
                        Case 140400300         '''Warrant fee
                            mWarantee = mWarantee + Format(Rec!fltAmount, "0.00")
'                        Case gbAcHeadCode, gbAcHeadCodeSurCentralGovtBuildArrear          '''Advance
'                            mAdvance = mAdvance + Format(Rec!fltAmount, "0.00")
                    
                    End Select
                Rec.MoveNext
                Wend
                If mAmtPTaxCurrent > 0 Then
                    Print #gbFileNO, "Property Tax(Current)"; Tab(26); PadL(Format(mAmtPTaxCurrent, "0.00"), 9); Tab(54); "Receivables for Property Tax(Current)"; Tab(128); PadL(Format(mAmtPTaxCurrent, "0.00"), 9)
                End If
                If mAmtPTaxArrear > 0 Then
                    Print #gbFileNO, "Property Tax(Arrears)"; Tab(26); PadL(Format(mAmtPTaxArrear, "0.00"), 9); Tab(54); "Receivables for Property Tax(Arrears)"; Tab(128); PadL(Format(mAmtPTaxArrear, "0.00"), 9)
                End If
                If mAmtLC > 0 Then
                 Print #gbFileNO, "Library Cess "; Tab(26); PadL(Format(mAmtLC, "0.00"), 9); Tab(54); "Library Cess Payable"; Tab(128); PadL(Format(mAmtLC, "0.00"), 9)
                End If
                If mAmtPenal > 0 Then
                    Print #gbFileNO, "Penal Interest"; Tab(26); PadL(Format(mAmtPenal, "0.00"), 9); Tab(54); "Penal Interest"; Tab(128); PadL(Format(mAmtPenal, "0.00"), 9)
                End If
                If mAmtServiceCess > 0 Then
                    Print #gbFileNO, "Service Cess "; Tab(26); PadL(Format(mAmtServiceCess, "0.00"), 9); Tab(54); "Receivables for Service Cess"; Tab(128); PadL(Format(mAmtServiceCess, "0.00"), 9)
                End If
                If mAmtSurcharge > 0 Then
                    Print #gbFileNO, "Surcharge"; Tab(26); PadL(Format(mAmtSurcharge, "0.00"), 9); Tab(54); "Receivables for Surcharge"; Tab(128); PadL(Format(mAmtSurcharge, "0.00"), 9)
                End If
                If mSplServices > 0 Then
                    Print #gbFileNO, "Special Service"; Tab(26); PadL(Format(mSplServices, "0.00"), 9); Tab(54); "Fees on Buildings for Special Service"; Tab(128); PadL(Format(mSplServices, "0.00"), 9)
                End If
                If mSurCentralGovtBuild > 0 Then
                    Print #gbFileNO, " Service Charge"; Tab(26); PadL(Format(mSurCentralGovtBuild, "0.00"), 9); Tab(54); "Service Charge on Central Govt Buildings"; Tab(128); PadL(Format(mSurCentralGovtBuild, "0.00"), 9)
                End If
                If mNotice > 0 Then
                    Print #gbFileNO, "Notice Fee"; Tab(26); PadL(Format(mNotice, "0.00"), 9); Tab(54); "NoticeFee"; Tab(128); PadL(Format(mNotice, "0.00"), 9)
                End If
                If mWarantee > 0 Then
                    Print #gbFileNO, "Warrant Fee"; Tab(26); PadL(Format(mWarantee, "0.00"), 9); Tab(54); "WarrantFee"; Tab(128); PadL(Format(mWarantee, "0.00"), 9)
                End If
'                Print #gbFileNO, Tab(1); "Advance"; Tab(4); PadL(Format(mAdvance, "0.00"), 9)
                
            End If
             Rec.MoveFirst
            While Not Rec.EOF
                        If gbLBPanchayat Then
                            If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Current Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Arrear Then
                                If mStartingYear > Rec!intYearID Then
                                    mStartingYear = Rec!intYearID
                                    mStartingPeriod = Rec!tnyPeriodID
                                End If
                                If mEndingYear < Rec!intYearID Then
                                    mEndingYear = Rec!intYearID
                                End If
                                mEndingPeriod = Rec!tnyPeriodID
                            End If
                        Else
                            If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Then
                                If mStartingYear > Rec!intYearID Then
                                    mStartingYear = Rec!intYearID
                                    mStartingPeriod = Rec!tnyPeriodID
                                End If
                                If mEndingYear < Rec!intYearID Then
                                    mEndingYear = Rec!intYearID
                                End If
                                mEndingPeriod = Rec!tnyPeriodID
                            End If
                        End If
                        Rec.MoveNext
                Wend
        Rec.MoveFirst
        Print #gbFileNO,
        mLoop = mLoop + 1
        mNarration = "(Being the " & Rec!vchTransactionType
        mNarra = "Collected for the Period"
        Print #gbFileNO, mNarration; Tab(52); mNarration
        Print #gbFileNO, mNarra; Tab(52); mNarra
        mLoop = mLoop + 1
        mNarration = "  of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
        If mStartingPeriod = 1 Then
            mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
        ElseIf mStartingPeriod = 2 Then
            mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
        Else
            mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
        End If

        If mEndingPeriod = 1 Then
            mNarration = mNarration & " Ist Hf )"
        ElseIf mEndingPeriod = 2 Then
            mNarration = mNarration & " IInd Hf )"
        Else
            mNarration = mNarration & ")"
        End If
        mLoop = mLoop + 1
        Print #gbFileNO, mNarration; Tab(52); mNarration

            Rec.MoveFirst
'            For mCount = mLoop + 1 To 9
'                Print #gbFileNO,
'            Next mCount

            'Print #gbFileNO, Tab(37); "Adv.Adj("; Format(Rec!fltAdvAmtAdj, "0.00"); ")"; Tab(126); "Adv.Adj("; Format(Rec!fltAdvAmtAdj, "0.00"); ")"
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(74); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(22); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(74); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
            Dim SCount As Integer
            Dim SecondString As String
            If Len(Rupees(Rec!TotalAmt)) > 32 Then
                strSubstr = Left$(Rupees(Rec!TotalAmt), 32)
                SCount = Len(Rupees(Rec!TotalAmt)) - 32
                SecondString = Trim(Right(Rupees(Rec!TotalAmt), SCount))
                Print #gbFileNO, strSubstr
                Print #gbFileNO, SecondString
            Else
                Print #gbFileNO, Rupees(Rec!TotalAmt);
            End If
            
       '     Print #gbFileNO, Rupees(Rec!TotalAmt);
            Print #gbFileNO, Tab(60); Rupees(Rec!TotalAmt)
            Print #gbFileNO,
            Print #gbFileNO, Tab(5); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 40); Tab(61); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            'Print #gbFileNO,
            Dim Uname As String
            Dim Counter As Integer
            Dim CounterDis As String
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                Counter = objCounter.CounterNo
                CounterDis = objCounter.CounterDescription
               ' Print #gbFileNO, Tab(11); objCounter.CounterNo;
               ' Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
            End If
            
            
            objUser.SetUser (Rec!intUserID)
            If objUser.UserID > -1 Then
                Uname = objUser.UserName
               ' Print #gbFileNO, Tab(20); objUser.UserName;
                'Print #gbFileNO, Tab(61); objUser.UserName
            End If
            Print #gbFileNO, Tab(11); Counter; Uname;
            Print #gbFileNO, Tab(61); Counter & " : " & CounterDis; Uname
       ' End If   by me



        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,

        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,

        'Print #gbFileNO, 'Chr$(27) + Chr$(12)

        Close #gbFileNO
        'ShellPad
        
        
        'Shell "Print " & gbFileName
        'Kill gbFileName
        
        
        
        Dim mFlag As Integer
        Dim X As Integer
        
        mFlag = Shell("Print " & gbFileName)
        Sleep 1000
        
    End Sub
    Public Sub PrintSummaryReceiptPTax(intVoucherID As Double)
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        'FileInitialize
        Open gbFileName For Output As #gbFileNO
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        
        
'        mSQL = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
'        mSQL = mSQL + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
'        mSQL = mSQL + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
'        mSQL = mSQL + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
'        mSQL = mSQL + " Where faVouchers.intVoucherID = " & intVoucherID
        objdb.SetConnection mCnn
'        Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic
        
        Select Case Rec!intInstrumentTypeID
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(13); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Else
            Print #gbFileNO,
        End Select
        
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(31); gbBold; gbDoubleWidth; Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); Tab(65); Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbBoldOff; gbDoubleWidthOff
            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            
            Print #gbFileNO, Tab(15); Style(mName, True); Tab(65); Style(mName, True)
            
            
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(65); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(65); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(65); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(65); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
            ' Line 15 Next
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(55); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            Print #gbFileNO,
            Print #gbFileNO,
            
            '======================================================================================='
            '  B O D Y     P A R T     O F      R E C E I P T                                       '
            '======================================================================================='
            ' Line 18 Next
            
            Dim mPTaxA As Variant
            Dim mPTaxC As Variant
            Dim mLCA As Variant
            Dim mLCC As Variant
            Dim mPCA As Variant
            Dim mPCC As Variant
            
            
            Dim mPenal As Variant
            Dim mRndOff As Variant
            Dim mOthers As Variant
            Dim mNarration As String
            Dim mNarra As String
            Dim mStartingYear As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear As Integer
            Dim mEndingPeriod As Integer
            mStartingYear = 2100
            Rec.MoveFirst
            While Not Rec.EOF
                Select Case Rec!intAccountHeadID
                    Case gbAcHeadIDPropertyTaxArrear
                        mPTaxA = mPTaxA + Rec!fltAmount
                        If mStartingYear > Rec!intYearID Then
                            mStartingYear = Rec!intYearID
                            mStartingPeriod = Rec!tnyPeriodID
                        End If
                        If mEndingYear < Rec!intYearID Then
                            mEndingYear = Rec!intYearID
                        End If
                        mEndingPeriod = Rec!tnyPeriodID
                    Case gbAcHeadIDPropertyTaxCurrent
                        mPTaxC = mPTaxC + Rec!fltAmount
                        If mStartingYear > Rec!intYearID Then
                            mStartingYear = Rec!intYearID
                            mStartingPeriod = Rec!tnyPeriodID
                        End If
                        If mEndingYear < Rec!intYearID Then
                            mEndingYear = Rec!intYearID
                        End If
                        mEndingPeriod = Rec!tnyPeriodID
                    Case gbAcHeadIDLibraryCess:
                        If Rec!tnyArrearFlag = 1 Then
                            mLCA = mLCA + Rec!fltAmount
                        Else
                            mLCC = mLCC + Rec!fltAmount
                        End If
                    Case gbAcHeadIDPoorHomeCess
                        If Rec!tnyArrearFlag = 1 Then
                            mPCA = mPCA + Rec!fltAmount
                        Else
                            mPCC = mPCC + Rec!fltAmount
                        End If
                    Case gbAcHeadIDPenalInterest: mPenal = mPenal + Rec!fltAmount
                    Case gbAcHeadIDRoundOff: mRndOff = mRndOff + Rec!fltAmount
                    Case Else:  mOthers = mOthers + Rec!fltAmount
                End Select
                
                Rec.MoveNext
            Wend
                
            If mPTaxA > 0 Then
                Print #gbFileNO, gbAcHeadCodePropertyTaxArrear; Tab(27); PadL(Format(mPTaxA, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Receivable for Property Tax (Arrear)", 60); Tab(109); PadL(Format(mPTaxA, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mLCA > 0 Then
                Print #gbFileNO, gbAcHeadCodeLibraryCess; Tab(27); PadL(Format(mLCA, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Library Cess Payable ", 60); Tab(109); PadL(Format(mLCA, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            
            If mPCA > 0 Then
                Print #gbFileNO, gbAcHeadCodePoorHomeCess; Tab(27); PadL(Format(mPCA, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Poor Home Cess Payable ", 60); Tab(109); PadL(Format(mPCA, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mPTaxC > 0 Then
                Print #gbFileNO, gbAcHeadCodePropertyTaxCurrent; Tab(27); PadL(Format(mPTaxC, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Receivable for Property Tax (Current)", 60); Tab(126); PadL(Format(mPTaxC, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mLCC > 0 Then
                Print #gbFileNO, gbAcHeadCodeLibraryCess; Tab(27); PadL(Format(mLCC, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Library Cess Payable ", 60); Tab(126); PadL(Format(mLCC, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mPCC > 0 Then
                Print #gbFileNO, gbAcHeadCodePoorHomeCess; Tab(27); PadL(Format(mPCC, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Poor Home Cess Payable ", 60); Tab(126); PadL(Format(mPCC, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
             
            If mPenal > 0 Then
                Print #gbFileNO, gbAcHeadCodePenalInterest; Tab(27); PadL(Format(mPenal, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Penal Interest ", 60); Tab(126); PadL(Format(mPenal, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mOthers > 0 Then
                Print #gbFileNO, " Others"; Tab(27); PadL(Format(mOthers, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Others ", 60); Tab(126); PadL(Format(mOthers, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            Print #gbFileNO,
            mNarration = "(Being the Property Tax Collected for the Period"
            Print #gbFileNO, mNarration; Tab(52); mNarration
            mLoop = mLoop + 1
            
            mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
            If mStartingPeriod = 1 Then
                mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 2 Then
                mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            Else
                mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            End If
            
            If mEndingPeriod = 1 Then
                mNarration = mNarration & " Ist Hf )"
            ElseIf mEndingPeriod = 2 Then
                mNarration = mNarration & " IInd Hf )"
            Else
                mNarration = mNarration & ")"
            End If
            
            Print #gbFileNO, mNarration; Tab(52); mNarration
            mLoop = mLoop + 2
            
            Rec.MoveFirst
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount

            'Print #gbFileNO, Tab(37); "Adv.Adj("; Format(Rec!fltAdvAmtAdj, "0.00"); ")"; Tab(126); "Adv.Adj("; Format(Rec!fltAdvAmtAdj, "0.00"); ")"
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
                            
            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(117); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
            
            Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
            Print #gbFileNO,
            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 40); Tab(61); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            Print #gbFileNO,
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                Print #gbFileNO, Tab(11); objCounter.CounterNo;
                Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
            End If
            objUser.SetUser (Rec!intUserID)
            If objUser.UserID > -1 Then
                Print #gbFileNO, Tab(11); objUser.UserName;
                Print #gbFileNO, Tab(61); objUser.UserName
            End If
        End If
        
        
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        
        Close #gbFileNO
        'ShellPad
        Shell "Print " & gbFileName
        'Kill gbFileName
    End Sub
    Public Sub PrintSummaryReceiptRLB(intVoucherID As Double)
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mTrType     As Integer
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        'FileInitialize
        Open gbFileName For Output As #gbFileNO
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,

        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic
        mTrType = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
        Select Case Rec!intInstrumentTypeID
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Else
            Print #gbFileNO,
        End Select
        
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(120); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            
            Print #gbFileNO, Tab(15); Style(mName, True); Tab(65); Style(mName, True); Tab(87); gbBold; "GSTIN : "; Tab(96); gbGSTIN;
            
            
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(65); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(65); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(65); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(65); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
            ' Line 15 Next
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(55); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            Print #gbFileNO,
            Print #gbFileNO,
            
            '======================================================================================='
            '  B O D Y     P A R T     O F      R E C E I P T                                       '
            '======================================================================================='
            ' Line 18 Next
            
            Dim mRLBA As Variant
            Dim mRLBC As Variant
            Dim mSTaxA As Variant
            Dim mSTaxC As Variant
            Dim mPCA As Variant
            Dim mCGST As Variant
            Dim mSGST As Variant
            
            Dim mPenal As Variant
            Dim mRndOff As Variant
            Dim mOthers As Variant
            Dim mNarration As String
            Dim mNarra As String
            Dim mStartingYear As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear As Integer
            Dim mEndingPeriod As Integer
            mStartingYear = 2100
            Rec.MoveFirst
            While Not Rec.EOF
                Select Case Rec!intAccountHeadID
                    Case gbAcHeadIDRentLandArrear, gbAcHeadIDCivicAmenitiesArrear
                        mRLBA = mRLBA + Rec!fltAmount
                        If mStartingYear > Rec!intYearID Then
                            mStartingYear = Rec!intYearID
                            mStartingPeriod = Rec!tnyPeriodID
                        End If
                        If mEndingYear < Rec!intYearID Then
                            mEndingYear = Rec!intYearID
                        End If
                        mEndingPeriod = Rec!tnyPeriodID
                    Case gbAcHeadIDRentLandCurrent, gbAcHeadIDCivicAmenitiesCurrent
                        mRLBC = mRLBC + Rec!fltAmount
                        If mStartingYear > Rec!intYearID Then
                            mStartingYear = Rec!intYearID
                            mStartingPeriod = Rec!tnyPeriodID
                        End If
                        If mEndingYear < Rec!intYearID Then
                            mEndingYear = Rec!intYearID
                        End If
                        mEndingPeriod = Rec!tnyPeriodID
                    Case gbAcHeadIDServiceTax:
                        If Rec!tnyArrearFlag = 1 Then
                            mSTaxA = mSTaxA + Rec!fltAmount
                        Else
                            mSTaxC = mSTaxC + Rec!fltAmount
                        End If
                    Case gbAcHeadIDCGST
                        If Rec!tnyArrearFlag = 1 Then
                            mCGST = mCGST + Rec!fltAmount
                        Else
                            mCGST = mCGST + Rec!fltAmount
                        End If
                    Case gbAcHeadIDSGST:
                        If Rec!tnyArrearFlag = 1 Then
                            mSGST = mSGST + Rec!fltAmount
                        Else
                            mSGST = mSGST + Rec!fltAmount
                        End If
                    Case gbAcHeadIDPenalInterest: mPenal = mPenal + Rec!fltAmount
                    Case gbAcHeadIDRoundOff: mRndOff = mRndOff + Rec!fltAmount
                    Case Else:  mOthers = mOthers + Rec!fltAmount
                End Select
                Rec.MoveNext
            Wend
            If mTrType = gbTransactionTypeRentOnBuilding Then
                If mRLBA > 0 Then
                    Print #gbFileNO, gbAcHeadCodeRLBArrear; Tab(27); PadL(Format(mRLBA, "0.00"), 9);
                    Print #gbFileNO, Tab(48); PadR("Rent receivable from Civic Amenities (Arrears)", 60); Tab(109); PadL(Format(mRLBA, "0.00"), 9)
                    mLoop = mLoop + 1
                End If
            ElseIf mTrType = gbTransactionTypeRentOnLand Then
                If mRLBA > 0 Then
                    Print #gbFileNO, gbAcHeadCodeRLBArrear; Tab(27); PadL(Format(mRLBA, "0.00"), 9);
                    Print #gbFileNO, Tab(48); PadR("Rent receivable from Lease on Lands (Arrears)", 60); Tab(109); PadL(Format(mRLBA, "0.00"), 9)
                    mLoop = mLoop + 1
                End If
            End If
            If mSTaxA > 0 Then
                Print #gbFileNO, gbAcHeadCodeServiceTax; Tab(27); PadL(Format(mSTaxA, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Service Tax Payable ", 60); Tab(109); PadL(Format(mSTaxA, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            If mTrType = gbTransactionTypeRentOnBuilding Then
                If mRLBC > 0 Then
                    Print #gbFileNO, gbAcHeadCodeCivicAmenitiesCurrent; Tab(27); PadL(Format(mRLBC, "0.00"), 9);
                    Print #gbFileNO, Tab(48); PadR("Rent receivable from Civic Amenities (Current)", 60); Tab(109); PadL(Format(mRLBA, "0.00"), 9)
                    mLoop = mLoop + 1
                End If
            ElseIf mTrType = gbTransactionTypeRentOnLand Then
                If mRLBC > 0 Then
                    Print #gbFileNO, gbAcHeadCodeRentLandCurrent; Tab(27); PadL(Format(mRLBC, "0.00"), 9);
                    Print #gbFileNO, Tab(48); PadR("Rent receivable from Lease on Lands (Current)", 60); Tab(126); PadL(Format(mRLBA, "0.00"), 9)
                    mLoop = mLoop + 1
                End If
            End If
  
            If mSTaxC > 0 Then
                Print #gbFileNO, gbAcHeadCodeServiceTax; Tab(27); PadL(Format(mSTaxC, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Service Tax Payable ", 60); Tab(126); PadL(Format(mSTaxC, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mCGST > 0 Then
                Print #gbFileNO, gbAcHeadCodeCGST; Tab(27); PadL(Format(mPenal, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Goods And Service Tax-CGST", 60); Tab(126); PadL(Format(mCGST, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            If mSGST > 0 Then
                Print #gbFileNO, gbAcHeadCodeSGST; Tab(27); PadL(Format(mPenal, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Goods And Service Tax-SGST", 60); Tab(126); PadL(Format(mSGST, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            If mPenal > 0 Then
                Print #gbFileNO, gbAcHeadCodePenalInterest; Tab(27); PadL(Format(mPenal, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Penal Interest ", 60); Tab(126); PadL(Format(mPenal, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            If mOthers > 0 Then
                Print #gbFileNO, " Others"; Tab(27); PadL(Format(mOthers, "0.00"), 9);
                Print #gbFileNO, Tab(48); PadR("Others ", 60); Tab(126); PadL(Format(mOthers, "0.00"), 9)
                mLoop = mLoop + 1
            End If
            
            Print #gbFileNO,
            mNarration = "(Being the Building/Bunk Rent Collected for the Period"
            Print #gbFileNO, mNarration; Tab(52); mNarration
            mLoop = mLoop + 1
            
            mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
            If mStartingPeriod = 11 Then
                mNarration = mNarration & " Jan to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 12 Then
                mNarration = mNarration & " Feb to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 13 Then
                mNarration = mNarration & " Mar to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 14 Then
                mNarration = mNarration & " Apr to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 15 Then
                mNarration = mNarration & " May to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 16 Then
                mNarration = mNarration & " Jun to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 17 Then
                mNarration = mNarration & " Jul to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 18 Then
                mNarration = mNarration & " Aug to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 19 Then
                mNarration = mNarration & " Sep to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 20 Then
                mNarration = mNarration & " Oct to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 21 Then
                mNarration = mNarration & " Nov to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            ElseIf mStartingPeriod = 22 Then
                mNarration = mNarration & " Dec to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            Else
                mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
            End If
            
            If mEndingPeriod = 11 Then
                mNarration = mNarration & " Jan )"
            ElseIf mEndingPeriod = 12 Then
                mNarration = mNarration & " Feb )"
            ElseIf mEndingPeriod = 13 Then
                mNarration = mNarration & " Mar )"
            ElseIf mEndingPeriod = 14 Then
                mNarration = mNarration & " Apr )"
            ElseIf mEndingPeriod = 15 Then
                mNarration = mNarration & " May )"
            ElseIf mEndingPeriod = 16 Then
                mNarration = mNarration & " Jun )"
            ElseIf mEndingPeriod = 17 Then
                mNarration = mNarration & " Jul )"
            ElseIf mEndingPeriod = 18 Then
                mNarration = mNarration & " Aug )"
            ElseIf mEndingPeriod = 19 Then
                mNarration = mNarration & " Sep )"
            ElseIf mEndingPeriod = 20 Then
                mNarration = mNarration & " Oct )"
            ElseIf mEndingPeriod = 21 Then
                mNarration = mNarration & " Nov )"
            ElseIf mEndingPeriod = 22 Then
                mNarration = mNarration & " Dec )"
            Else
                mNarration = mNarration & ")"
            End If
            
            Print #gbFileNO, mNarration; Tab(52); mNarration
            mLoop = mLoop + 2
            
            Rec.MoveFirst
            For mCount = mLoop + 1 To 6
                Print #gbFileNO,
            Next mCount

            'Print #gbFileNO, Tab(37); "Adv.Adj("; Format(Rec!fltAdvAmtAdj, "0.00"); ")"; Tab(126); "Adv.Adj("; Format(Rec!fltAdvAmtAdj, "0.00"); ")"
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
                            
            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(117); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
            
            Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
            Print #gbFileNO,
            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 40); Tab(61); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            Print #gbFileNO,
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                Print #gbFileNO, Tab(11); objCounter.CounterNo;
                Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
            End If
            objUser.SetUser (Rec!intUserID)
            If objUser.UserID > -1 Then
                Print #gbFileNO, Tab(11); objUser.UserName;
                Print #gbFileNO, Tab(61); objUser.UserName
            End If
        End If

        Close #gbFileNO
        'Shell "Print " & gbFileName
        'Kill gbFileName
        
        
        
        Dim mFlag As Integer
        Dim X As Integer
        
        mFlag = Shell("Print " & gbFileName)
        Sleep 1000
        
    End Sub
    
    Public Function CalculateFineforPTax(mYearID As Integer, mPeriodID As Integer, mPTax As Double) As Double
        '==============================================================================='
        ' Modified By : Aiby                                                            '
        '             : For Kollam  Corporation                                         '
        '                                                                               '
        '==============================================================================='
        
        Dim dtFromDt As Variant
        Dim mNoOfMonths As Long
        Dim mAmount     As Double
        'Debug.Print mYearID, mPeriodID, mPTax
        'mFineAmt = 0
        
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation Mode 1= Act and 2 = Circular                          '
        '-------------------------------------------------------------------------------'
        'If mYearID = 2009 Then Stop
        
        If gbFineCalculationMode = 1 Then
            
            If mPeriodID = 1 Then
                dtFromDt = DateSerial(mYearID, 10, 1)
            Else
                dtFromDt = DateSerial(mYearID, 4, 1)
            End If
            If mYearID = gbFinancialYearID And mPeriodID = 2 Then
                CalculateFineforPTax = 0
                Exit Function
            End If
            
            If mYearID < 2006 Then
                If mYearID = 2005 Then
                    If mPeriodID > 1 Then
                        GoTo Skip:
                    End If
                End If
                mNoOfMonths = (2005 - mYearID) * 24 + 10
                If mPeriodID = 2 Then
                    mNoOfMonths = mNoOfMonths - 12
                End If
                'mNoOfMonths = mNoOfMonths + (gbFinancialYearID - 2005) * 12
                
                mYearID = 2005
                mPeriodID = 2
            End If
Skip:
            mNoOfMonths = mNoOfMonths + (gbFinancialYearID - mYearID) * 12
            
            If mNoOfMonths = 0 Then
                If mYearID = gbFinancialYearID Then
                    If mPeriodID = 1 Then
                        mNoOfMonths = -5
                    Else
                        mNoOfMonths = 0
                    End If
                End If
            End If
            
            If mPeriodID = 2 Then
                mNoOfMonths = mNoOfMonths - 6
            End If
            
            If Month(gbTransactionDate) > 3 Then
                mNoOfMonths = mNoOfMonths + Month(gbTransactionDate) - 3
            Else
                mNoOfMonths = mNoOfMonths + Month(gbTransactionDate) + 9
            End If
            
            mNoOfMonths = mNoOfMonths - 1
            'Debug.Print mNoOfMonths
            CalculateFineforPTax = mPTax * mNoOfMonths / 100
            Exit Function
            
        ElseIf gbFineCalculationMode = 2 Then
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation As Per Circular                                       '
        '-------------------------------------------------------------------------------'
            'mPTax = Format(mPTax * 2, "0.00")
            dtFromDt = DateSerial(mYearID, 11, 1)
            If mYearID = gbFinancialYearID Then
                CalculateFineforPTax = 0
                Exit Function
            End If
            If mYearID < 2005 Then
                mNoOfMonths = DateDiff("m", dtFromDt, DateSerial(2005, 8, 1))
                dtFromDt = DateSerial(2005, 9, 1)
                mNoOfMonths = mNoOfMonths + DateDiff("m", dtFromDt, gbTransactionDate)
            End If
            mNoOfMonths = mNoOfMonths + DateDiff("m", dtFromDt, gbTransactionDate) + 1
            CalculateFineforPTax = mPTax * mNoOfMonths / 100
            Exit Function
        End If
    End Function
    
    Public Function RoundOffAdjustment(mAmt As Double) As Single
        Dim mStrAmt As String
        Dim mTemp As String
        Dim mStrDecimal As String
        
        mTemp = Trim(str(mAmt))
        mStrAmt = Token(mTemp, ".")
        Select Case Len(mTemp)
            Case Is = 0: mTemp = ".00"
            Case Is = 1: mTemp = "." & mTemp & "0"
            Case Is = 2: mTemp = "." & mTemp
            Case Else
                mTemp = Left(mTemp, 2)
                mTemp = "." & mTemp
        End Select
        mStrAmt = mStrAmt + mTemp
        If 1 - mTemp = 1 Then
            RoundOffAdjustment = 0
        Else
            RoundOffAdjustment = 1 - mTemp
        End If
    End Function
    
    Function IsIKMLAB(mAcID As String) As Boolean
        IsIKMLAB = False
        If Date <= DateSerial(2021, 12, 1) Then
            If mAcID = "E0CB4E34C783" Or _
               mAcID = "001EC92B058E" Or _
               mAcID = "001FD0E79EE3" Or _
               mAcID = "001EC92B0484" Or _
               mAcID = "001EC92B01FF" Or _
               mAcID = "001EC92B042A" Or _
               mAcID = "0016EC8D21DC" Or _
               mAcID = "0800272D5CEE" Or _
               mAcID = "0800273D5DED" Or _
               mAcID = "080027A6CF8D" Then
               IsIKMLAB = True
            End If
        End If

        
    End Function

    Function GetRendomKey(mLen As Integer) As String
        Randomize
        Dim AllowableChars As String
        Dim i As Integer
        Dim mStr As String
        Dim mLength As Integer
        'mLength = 9
        AllowableChars = "abcdefghijklmnopqrstuvwxyz1234567890"
        For i = 1 To mLen
           mStr = mStr & mID$(AllowableChars, Int(Rnd() * Len(AllowableChars) + 1), 1)
        Next
        GetRendomKey = mStr
    End Function
    
    Sub LogFile(Message As String)
            Dim LogFile As Integer
            LogFile = FreeFile
            Open App.Path + "\SaankhyaDE.Log" For Append As #LogFile
            Print #LogFile, Message
            Close #LogFile
    End Sub
        
    Sub MDIChild(mParent As Long, mChild As Long)
        SetParent mChild, mParent
    End Sub
    
    Sub ModalForm(mChild As Long)
        SetParent mChild, 0
    End Sub
     Public Sub SetgbLastPostingDate() '***************TO SET THE LAST POSTING DATE***********************'
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mSql    As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "SELECT MAX(dtPostingDate) dtPostingDate FROM faPostingIndex WHERE tnyStage=2"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            gbLastPostingDate = Format(Rec!dtPostingDate, "dd-mmm-yyyy")
        End If
        Rec.Close
        mCnn.Close
    End Sub
    
    Public Sub UpdateVoucherIndex(mVoucherID As Long)   '******TO UPDATE tnyChangeFlag=1 in faVoucherIndex***************
        Dim mCnn   As New ADODB.Connection
        Dim Rec    As New ADODB.Recordset
        Dim objdb  As New clsDB
        Dim mSql   As String
    
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        'mSQL = "UPDATE faVouchers SET tnyChangeFlag=1 WHERE intVoucherID = " & mVoucherID
        'objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        
        mSql = "UPDATE faVoucherIndex SET tnyChangeFlag=1 WHERE intVoucherID = " & mVoucherID
        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        mCnn.Close
        
    End Sub
    Private Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String
    ''' To check System date format
    
     Dim sReturn As String
     Dim r As Long
    'call the function passing the Locale type 'variable to retrieve the required size of
    'the string buffer needed
     r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    'if successful..
     If r Then
       'pad the buffer with r spaces
        sReturn = Space$(r)
       'and call again passing the buffer
        r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
       'if successful (r > 0)
        If r Then
          'r holds the size of the string
          'including the terminating null
           GetUserLocaleInfo = Left$(sReturn, r - 1)
        End If
     End If
    End Function
