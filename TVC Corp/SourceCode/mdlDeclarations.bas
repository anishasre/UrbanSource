Attribute VB_Name = "mdlDeclarations"
Option Explicit
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long

    Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer$, nSize As Long) As Long
    Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    
    Public Const KEYEVENTF_EXTENDEDKEY = &H1
    Public Const KEYEVENTF_KEYUP = &H2
    Public Const ib_Tab = 9
    
    Public Declare Function SendMyMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
      ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
    
    Public Const CB_ADDSTRING As Long = &H143
    Public Const LB_ADDSTRING As Long = &H180
    
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    'Sertching and Tab Stop
    Public Const CB_FINDSTRING = &H14C
    Public Const LB_FINDSTRING = &H18F
    Public Const CB_LIMITTEXT = &H141
    Public Const LB_SETTABSTOPS = &H192
    
    
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    '------------ADDED BY SYALIMA ON 29-01-2014-------------------
    Public lSoochikaFileIDKMBR As Variant
    '-------------------------------------------------------------
    
    
    '------------------------------------'
    ' V E R S I O N   C O N T R O L      '
    '------------------------------------'
    'For Corporation And Muncipality
    Public Const gbVerID = "2"
    Public Const gbVerSubID = "2.23" 'modified On 16.10.15
    Public Const gbDBVerID = "1"
    'Public Const gbDBSubVerID = "0.35" 'modified On 17.9.18
    'Public Const gbDBSubVerID = "0.37" 'Modif 16 Mar 19 'modified On 17.9.18
    Public Const gbDBSubVerID = "0.37" 'Modif 16 Mar 19 'modified On 17.9.18
    'Public Const gbDBSubVerID = "0.38" 'Modif 16 Mar 19 'modified On 31.1.20  for Pofessiontax 6 5 2020
    
    'For Panchayat
    Public Const gbPVerID = "2"
    Public Const gbPVerSubID = "2.20" 'modified On 20.10.15
    Public Const gbPDBVerID = "1"
    'Public Const gbPDBSubVerID = "0.34" 'modified On 17.9.18
    'Public Const gbPDBSubVerID = "0.36" ' modified On 17.9.18
    Public Const gbPDBSubVerID = "0.36" 'Modif 16 Mar 19 modified On 17.9.18
    'Public Const gbPDBSubVerID = "0.37" 'Modif 16 Mar 19 modified On 17.9.20 for Pofessiontax  6 5 2020
    
    ' ------------------------------------'
    
    Public gbSaankhyaINI    As String
    Public gbTempFileName   As String
    Public gbConnectionToFa As New ADODB.Connection
    
    Public gbSearchID           As Double
    Public gbSearchStr          As String
    Public gbSearchCode         As String
    Public gbSearchSubID        As Double
    Public gbProject            As uProject
    
    Public intWrBillSearchID        As Integer
    Public intWrBillCircleID        As Integer
    Public intWrBillDivisionID      As Integer
    Public intWrBillSubDivisionID   As Integer
    Public intWrBillCaretakerID     As Integer

    
    Public gbSeatGroupChairPerson            As Integer
    Public gbSeatGroupDyChairPersion         As Integer
    Public gbSeatGroupSecretary              As Integer
    Public gbSeatGroupAdditionalSecretary    As Integer
    Public gbSeatGroupAccountsOfficer        As Integer
    Public gbSeatGroupAccountsSuperintended  As Integer
    Public gbSeatGroupAccountsClerk          As Integer
    Public gbSeatGroupCashSuperintended      As Integer
    Public gbSeatGroupChiefCashier           As Integer
    Public gbSeatGroupCashier                As Integer
    Public gbSeatGroupDemandSectionSuperintended As Integer
    Public gbSeatGroupDemandSectionClerk     As Integer
    Public gbSeatGroupAuditorsGroup          As Integer
    Public gbSeatGroupAccountSectionClerk    As Integer
    Public gbSeatGroupHeadClerk              As Integer ''added on 20 Feb 2015
    Public gbSeatGroupAssistantSecretary     As Integer ''added on 23 Feb 2015
    
    
    Public gbTransactionDate    As Date
    Public gbStartingDate       As Date
    Public gbEndingDate         As Date
    
    Public gbLastPostingDate    As Date
    
    Public gbLocalBodyID        As Long
    Public gbLBName             As String
    Public gbLBTitle            As String
    Public gbLBType             As Integer
    Public gbLBPanchayat        As Integer  ' set as 1 if lbTypeID =1,2,5
    Public gbLocationID         As Long      ' Location Where the Application is installed
    Public gbLocation           As String
    Public gbFinancialYearID    As Long
    Public gbCurrentPeriodID    As Integer
    Public gbDistrict           As String
    Public gbLBCODE           As String ''Added by Anisha On 11/Jul/18
    
                                     
    Public gbUserID         As Double
    Public gbUserName       As String
    Public gbUserTypeID     As Long
    Public gbUserActiveFlag As Integer
    
    
    Public gbCounterID              As Long
    Public gbCounterNo              As Long
    Public gbCounterIP              As String
    Public gbCounterName            As String
    Public gbCounterMacID           As String
    Public gbCounterSectionID       As Integer
    Public gbCounterSection         As String
    Public gbCounterOperationModeID As Integer
    Public Const gbJSKSectionID = 99 ' Setting Janasevana Kendram Section ID As 99
    Public gbGSTIN          As String ''Added by Anisha On 14/Sep/18
    
    
    Public gbShiftID        As Variant
    Public gbShiftName      As Variant
    Public gbSeatID         As Variant
    Public gbSeatGroupID    As Variant
    Public gbSeatName       As String
    Public gbReceiptDate    As Date
    Public Const gbSeatGroupAccountsSupt = 4 'Accounts Superintend
    
    Public gbSeatByDeveloper As Variant
    Public gbSeatCashGroupID As Variant
    
    Public gbSectionID      As Variant ' User Section ID
    Public gbSectionName    As Variant ' User Section
    
    Public gbUserSessionID      As Variant      ' User Login Sessions from UserMovement Table
    Public gbCounterStatusFlag  As Variant   ' To check Counter Closed or not
    
    Public gbTitle1     As String
    Public gbTitle2     As String
    
    
    
    
    Public gbInstrumentCheque               As Long
    Public gbInstrumentCash                 As Long
    Public gbInstrumentCard                 As Long
    Public gbInstrumentLetterOfAuthority    As Long
    Public gbSanchayaDbName                 As String
    
    Public gbFundID     As Long
    Public gbFundCode   As Variant '--Added Sinoj on 26 Jan 2010
    '-----------------------------------------------------------
    ' -----FaConfig Variables-----------------------------------
    Public gbLinkWithPropertyTax As Integer
    Public gbLinkWithProfTaxEmp  As Integer 'Added By Poornima On 22-09-2010
    Public gbLinkWithRentOnLand  As Integer
    Public gbLinkWithFinanceHO   As Integer
    Public gbLinkWithSevana      As Integer
    Public gbLinkWithSugama      As Integer
    Public gbLinkWithDandOPFA    As Integer
    Public gbFetchDemandFromHO   As Integer
    Public gbLinkWithMOReturn    As Integer
    Public gbLinkWithSoochika    As Integer ' 1=Urban 2=Unicode 'Added by Aiby on 14-Aug-2012
    Public gbFetchDemandFromWeb  As Integer '1= Property Tax Demand From Web 'Added by Anisha on 20-JUl-2015
    Public gbSaankhyaWeb         As Integer ' 1=Saankhay Web is enabled
    Public gbLinkWithDandOWeb    As Integer ' 1=D&O Web Integration is Active
    Public gbLinkWithProfTradeWeb    As Integer ' 1=Professiontax Trades/Institution Web Integration is Active
    Public gbLinkWithProfEmpWeb    As Integer ' 1=Professiontax EmployeesWeb Integration is Active
    
    ' Added on 28/Jun/2011  '' To Enable all Instruments for all transactions in panchayat Before gbOnlinedate
    Public gbOnlinedate          As Variant
    Public gbRPOnlinedate        As Variant
    '----Added By Anisha For Replacing INI File Details in dataBase (faConfig) On 25 Mar 2010
    Public gbDefaultBankID              As Variant
    Public gbDefaultUrl                 As Variant
    Public gbDefaultUrlForRequisition  As Variant
    Public gbRemittingBank              As Variant
    Public gbRemittingPlaceOfBank       As Variant
    Public gbDefaultUrlSanchayaPost     As Variant
    '-------------------------------------------------------------
    Public gbDefaultTransactionTypeID   As Variant
    Public gbBold                       As String
    Public gbBoldOff                    As String
    '-----------------------------------------
    
  
    Public gbContense       As String
    Public gbContenseOff    As String
            
    Public gbDoubleWidth    As String
    Public gbDoubleWidthOff As String
    Public gbColLines       As Long
    
    
    
    Public gbGeneralTransactionIDReceipts As Long  ' Using in Receipts to generate
    Public gbGeneralTransactionIDPayments As Long
    Public gbGeneralTransactionIDContraE  As Long
    Public gbGeneralTransactionIDJournal  As Long
    
    'Transaction Type
    Public gbTransactionTypePTax            As Long
    Public gbTransactionTypeRentOnBuilding  As Long
    Public gbTransactionTypeRentOnLand      As Long
    Public gbTransactionTypeProfTaxTrade    As Long
    Public gbTransactionTypeProfTaxTradeAccrual As Long
    Public gbTransactionTypeProfTaxEmp      As Long
    Public gbTransactiontypeDailyCollection As Long 'added By poornima on 14/10/2010
    Public gbTransactionTypeRefundOfPayment As Long    'added By Sinoj On 01/Mar/2010
    
    Public gbTransactionTypeDandO           As Long
    Public gbTransactionTypePFA             As Long
    
    Public gbTransactionTypeETax            As Long
    Public gbTransactionTypeSTax            As Long
    Public gbTransactionTypeKCR             As Long
    Public gbTransactionTypePPR             As Long
    Public gbTransactionTypeHall            As Long
    Public gbTransactionTypeTransferCredit  As Long ''Added On 29 Sep 2015 by Anisha    Contra Transaction Type
    Public gbTransactionTypePTaxGp          As Long
    
    
    Public gbAcHeadCodeCash                 As String
    Public gbAcHeadIDCash                   As Long
    
    Public gbAcHeadCodePropertyTaxCurrent   As String
    Public gbAcHeadCodePropertyTaxArrear    As String
    
    ''  Added On 11/oCT/2012
    Public gbAcHeadCodeServicceCessCurrent  As String
    Public gbAcHeadCodeServicceCessArrear   As String
    
    Public gbAcHeadCodeSurPTCurrent         As String
    Public gbAcHeadCodeSurPTArrear          As String
    Public gbAcHeadCodeSurCentralGovtBuildCurrent   As String
    Public gbAcHeadCodeSurCentralGovtBuildArrear    As String
    Public gbAcHeadCodeSplServicesCurrent           As String
    Public gbAcHeadCodeSplServicesArrear            As String

   
    
    Public gbAcHeadCodePropertyTax_NonResidential_Current   As String
    Public gbAcHeadCodePropertyTax_NonResidential_Arrear    As String
    
    Public gbAcHeadCodeLibraryCess      As String
    Public gbAcHeadCodePoorHomeCess     As String
    Public gbAcHeadCodePenalInterest    As String
    Public gbAcHeadCodeRoundOff         As String
    Public gbAcHeadCodeAdvancePTax      As String
    Public gbAcHeadCodeNoticeFee        As String    '' Added On 31/Aug/2016
    
    Public gbAcHeadIDPropertyTaxCurrent As Long
    Public gbAcHeadIDPropertyTaxArrear  As Long
    
     '' Added on 15/Nov/2012
    
     'Public gbAcHeadCodeServicceCessCurrentR As String
    ' Public gbAcHeadCodeServicceCessArrearR As String
     Public gbAcHeadCodeServicceCessCurrentNonR As String
     Public gbAcHeadCodeServicceCessArrearNonR  As String
                    
'     Public gbAcHeadCodeSurPTCurrentR As String
'     Public gbAcHeadCodeSurPTArrearR As String
     Public gbAcHeadCodeSurPTCurrentNonR    As String
     Public gbAcHeadCodeSurPTArrearNonR     As String
                    
'     Public gbAcHeadCodeSurCentralGovtBuildCurrentP As String
'     Public gbAcHeadCodeSurCentralGovtBuildArrearP As String

'     Public gbAcHeadCodeSplServicesCurrentP As String
'     Public gbAcHeadCodeSplServicesArrearP As String
'
    
    ''  Added On 11/oCT/2012
    Public gbAcHeadIDServicceCessCurrent    As String
    Public gbAcHeadIDServicceCessArrear     As String
    Public gbAcHeadIDSurPTCurrent           As String
    Public gbAcHeadIDSurPTArrear            As String
    Public gbAcHeadIDSurCentralGovtBuildCurrent As String
    Public gbAcHeadIDSurCentralGovtBuildArrear  As String
    Public gbAcHeadIDSplServicesCurrent         As String
    Public gbAcHeadIDSplServicesArrear          As String
    
    Public gbAcHeadIDPropertyTax_NonResidential_Current As Long
    Public gbAcHeadIDPropertyTax_NonResidential_Arrear  As Long
    
    Public gbAcHeadIDLibraryCess    As Long
    Public gbAcHeadIDPoorHomeCess   As Long
    Public gbAcHeadIDPenalInterest  As Long
    Public gbAcHeadIDRoundOff       As Long
    Public gbAcHeadIDAdvancePTax    As Long
    Public gbAcHeadIDNoticeFee      As Long         '' Added On 31/Aug/2016
    
    '----------'Note:-Added by Vinod For Allotment Letter On 04/01/2011------------'
    
    Public gbAcHeadIDDevelopmentFundGeneralCapital      As Long
    Public gbAcHeadIDDevelopmentFundSCPCapital          As Long
    Public gbAcHeadIDDevelopmentFundTSPCapital          As Long
    Public gbAcHeadIDGeneralPurposeFund                 As Long
    
    Public gbAcHeadCodeDevelopmentFundGeneralCapital    As String
    Public gbAcHeadCodeDevelopmentFundSCPCapital        As String
    Public gbAcHeadCodeDevelopmentFundTSPCapital        As String
    Public gbAcHeadCodeGeneralPurposeFund               As String
    
    Public gbAcHeadIDTreasuryAccount1                   As Long
    Public gbAcHeadIDTreasuryAccount2                   As Long
    Public gbAcHeadIDTreasuryAccount3                   As Long
    Public gbAcHeadIDTreasuryAccount4                   As Long
    Public gbAcHeadIDTreasuryAccount5                   As Long
    Public gbAcHeadIDTreasuryAccount6                   As Long  'Added by Minu on 08.12.2012
    Public gbAcHeadIDTreasuryAccount7                   As Long
    Public gbAcHeadIDTreasuryAccountTSB                 As Long 'Added by Minu on 02.06.2015
    Public gbAcHeadIDTreasuryAccountSpecialTSB          As Long 'Added by Anisha on 06.Mar.2017 For JointVenture
    
    Public gbAcHeadCodeTreasuryAccount1                 As String
    Public gbAcHeadCodeTreasuryAccount2                 As String
    Public gbAcHeadCodeTreasuryAccount3                 As String
    Public gbAcHeadCodeTreasuryAccount4                 As String
    Public gbAcHeadCodeTreasuryAccount5                 As String
    Public gbAcHeadCodeTreasuryAccount6                 As String  'Added by Minu on 08.12.2012
    Public gbAcHeadCodeTreasuryAccount7                 As String
    Public gbAcHeadCodeTreasuryAccountTSB               As String 'Added by Minu on 02.06.2015
    Public gbAcHeadCodeTreasuryAccountSpecialTSB        As String ''Added by Anisha on 06.Mar.2017 For JointVenture
    
    Public gbAcHeadIDMaintenanceFundRoadAssets          As Long
    Public gbAcHeadIDMaintenanceFundNonRoadAssets       As Long
    Public gbAcHeadIDCentralFinanceCommission           As Long
    Public gbAcHeadIDKLGSDP                             As Long
    Public gbAcHeadIDSpecialGrant                       As Long
    Public gbAcHeadIDRoadRenovationGrant                As Long
    
    Public gbAcHeadCodeMaintenanceFundRoadAssets        As String
    Public gbAcHeadCodeMaintenanceFundNonRoadAssets     As String
    Public gbAcHeadCodeCentralFinanceCommission         As String
    Public gbAcHeadCodeKLGSDP                           As String
    Public gbAcHeadCodeSpecialGrant                     As String
    Public gbAcHeadCodeRoadRenovationGrant              As String
    
    
    Public gbAcHeadIDIAY                                As Long 'ADDED BY MINU FOR IAY
    Public gbAcHeadIDIAYSCP                             As Long
    Public gbAcHeadIDIAYTSP                             As Long
    
    Public gbAcHeadCodeIAY                              As String
    Public gbAcHeadCodeIAYSCP                           As String
    Public gbAcHeadCodeIAYTSP                           As String
    
    '-------------------------------------------------------------------------------'
    
    '----------'Note:-Added by Vinod For Subsidiary Cash Book On 10/01/2011---------'
    Public gbAcHeadIDUnemploymentWages As Long
    'Public gbAcHeadIDNetSalaryPayable                                   As Long
    Public gbAcHeadIDVehicleHireCharges                                 As Long
    Public gbAcHeadIDMiscAdministrationExpenses                         As Long
    Public gbAcHeadIDEquipmentHireCharges                               As Long
    Public gbAcHeadIDExpensesForBuryingUnclaimedDeadBodies              As Long
    Public gbAcHeadIDRepairsAndMaintenanceDrainage                      As Long
    Public gbAcHeadIDDevFundProgrammesPublicHealthAndSanitation         As Long
    Public gbAcHeadIDMiscAdvance                                        As Long
    Public gbAcHeadIDUnpaidSalaries                                     As Long
    
    Public gbAcHeadCodeUnemploymentWages As String
    'Public gbAcHeadCodeNetSalaryPayable                                   As String
    Public gbAcHeadCodeVehicleHireCharges                                 As String
    Public gbAcHeadCodeMiscAdministrationExpenses                         As String
    Public gbAcHeadCodeEquipmentHireCharges                               As String
    Public gbAcHeadCodeExpensesForBuryingUnclaimedDeadBodies              As String
    Public gbAcHeadCodeRepairsAndMaintenanceDrainage                      As String
    Public gbAcHeadCodeDevFundProgrammesPublicHealthAndSanitation         As String
    Public gbAcHeadCodeMiscAdvance                                        As String
    Public gbAcHeadCodeUnpaidSalaries                                     As String
    
    '-------------------------------------------------------------------------------'
    
    
    'Note:-For Rent On Land And Buildings
    Public gbAcHeadCodeRLBArrear As String
    Public gbAcHeadCodeRLBCurrent As String
'    Public gbAcHeadCodeServiceTax As String
    Public gbAcHeadCodeAdvanceRLB As String
   
    
      'Rent On Land
    Public gbAcHeadCodeRentLandArrear   As String
    Public gbAcHeadCodeRentLandCurrent  As String
    Public gbAcHeadCodeServiceTax       As String
    Public gbAcHeadCodeAdvanceLand      As String
    Public gbAcHeadCodeCGST             As String  '' Added on 27 sept 2017 For GST
    Public gbAcHeadCodeSGST             As String
    Public gbAcHeadCodeFloodCess        As String  '' Added On 26 Aug 2019
    
    Public gbAcHeadIDRentLandArrear     As Long
    Public gbAcHeadIDRentLandCurrent    As Long
    Public gbAcHeadIDServiceTax         As Long
    Public gbAcHeadIDAdvanceLand        As Long
    Public gbAcHeadIDCGST               As Long  '' Added on 11 Feb 2020 For GST  Print summary
    Public gbAcHeadIDSGST               As Long
        
        'Rent On Building
    Public gbAcHeadCodeCivicAmenitiesArrear     As String
    Public gbAcHeadCodeCivicAmenitiesCurrent    As String
    Public gbAcHeadCodeAdvanceBuilding          As String
    
    Public gbAcHeadIDCivicAmenitiesArrear As Long
    Public gbAcHeadIDCivicAmenitiesCurrent As Long
    Public gbAcHeadIDAdvanceBuilding As Long
    
    Public gbAcHeadCodeAdvanceDandO As Long      ''Added on 10 feb 2020
    Public gbAcHeadIDRLBArrear As Long
    Public gbAcHeadIDRLBCurrent As Long
    
    
    ' PROFESSION TAX
    Public gbAcHeadIDProfTaxEmployees As Long
    Public gbAcHeadIDProfTaxTraders As Long
    Public gbAcHeadIDProfTaxTradersCurrent As Long  ' Added on 13/02/2012 By poornima
    Public gbAcHeadIDProfTaxTradersArrears As Long  ' Added on 13/02/2012 By poornima
    
    Public gbAcHeadCodeProfTaxEmployees As String
    Public gbAcHeadCodeProfTaxTraders As String
    Public gbAcHeadCodeProfTaxTradersCurrent As String ' Added on 13/02/2012 By poornima
    Public gbAcHeadCodeProfTaxTradersArrears As String ' Added on 13/02/2012 By poornima
    
    
    '--------------------------------------------
    ' Added On 3/1/2011 Modified By Anisha
    ' For Panchayat Modification (To Avoid HeadCode Hard Coding)
    Public gbAcHeadIDOtherFee         As String
    Public gbAcHeadCodeOtherFee         As String
    '--------------------------------------------
    
    Public gbTransactionTypeBrith As Long
    Public gbTransactionTypeDeath As Long
    Public gbTransactionTypeMarriage As Long
    Public gbTransactionTypeCmnMarriage As Long
    
    Public gbTransactionTypeOutDoor As Long
    Public gbTransactionTypeZonalCollection As Long
    Public gbTransactionTypeFriendsCollection As Long
    
    Public gbTransactionTypeBFundSSSFund As Long
    Public gbTransactionTypeMoneyOrderReturns As Long
    
    Public gbTransactionTypeMOReturnsSocialSecurityPension  As Long
    
    Public gbTransactionTypeApplicationForPermitKMBR As Long
    Public gbTransactionTypePermitFeeFromKMBR As Long
    
    Public gbTransactionTypeSaleOfTenderForm As Long
    
    Public gbTransactionTypeContraRegularPension As Long
    Public gbTransactionTypeContraContingentPension As Long
    
    Public gbFunctionaryAccountsDepartmentID As Long
    Public gbFunctionaryAccountsDepartmentCode As String
    
    Public gbFunctionAccountsID As Long
    Public gbFunctionAccountsCode As String
    
    '---------------------------------------------------------------------------------- '
    ' Declarations Related Opening Balance    - Added by Poornima On 28/12/2010         '
    '---------------------------------------------------------------------------------- '
    Public gbAcHeadCodeForCapitalFund   As String
    Public gbSeatWiseFundID             As Integer
    ' ---------------------------------------------------------------------------------- '
    ' Declarations Related Payments - Transaction Types and Account Heads
    ' ---------------------------------------------------------------------------------- '
    
    Public gbTransactionTypePayBills As Long
    ''Added By Anisha C To capture UnUtilized Fund Amount to treasury
    Public gbTransactionTypeUnUtilizedAmount As Long
    Public gbTransactionTypeProjectExpGO As Long
    
    Public gbAcHeadCodeNetSalaryPayable     As String
    Public gbAcHeadIDNetSalaryPayable       As Long
    
    Public gbAcHeadCodeGrossSalaryPayable   As String
    Public gbAcHeadIDGrossSalaryPayable     As Long
    
    Public gbAcHeadDeductionStart           As Long
    Public gbAcHeadDeductionEnd             As Long
    Public gbAcHeadDeductionExcludeProfTax  As Long

    
    '-------------------------------------------------------------------------------------'
    '                       Added Traeasury Heads On 09 Sep 2010    Sinoj                 '
    '-------------------------------------------------------------------------------------'

    Public gbAcHeadCodeCentralPensionFundPayable                               As String
    Public gbAcHeadCodePensionAndGratuityPayable                               As String
    Public gbAcHeadCodeOtherReceivablesCur                                     As String
    Public gbAcHeadCodePensionFundForContingentStaff                           As String
    Public gbAcHeadCodeContributionToPensionFundForContingentStaff             As String
    Public gbAcHeadCodeRegularTrasuryPension                                   As String
    Public gbAcHeadCodeContingentTreasuryPension                               As String
    '-------------------------------------------------------------------------------------'
    
    '-----------------------------------------------------------------------------------'
    '                       Nature of Funds
    '-----------------------------------------------------------------------------------'
    ' Added By Poornima:
    ' Using in frmBanks
        Public gbOwnFund        As String
        Public gbSpecialFund    As String
        Public gbGrantFund      As String
    '-------------------------------------------------------------------------------------'
    
    '-------------------------------------------'
    '           Config Manual Receipt           '
    Public gbManualReceiptNewBool     As Boolean
    '-------------------------------------------'
    
   '--------------------------------------------------------------------------------'
   ' NOTE:
        Public gbBankChangePermitDate As Variant
   ' This variable is used to set date for permite change default banks for
   ' Transaction type receipts from other lsgs to another banks by User.
   ' Planed to block this facility by April,2014
   '--------------------------------------------------------------------------------'
    
    
    Public gbPDEMode As Boolean
    Public gbFineCalculationMode As Integer  ' 1= Act and 2 = Circular
    
    Public gbPrinterMode As Variant
    
    Public Enum AppID
        Saankhya = 115
        Payroll = 200
        Sulekha = 300
        Sanjaya = 107
        Sthapana = 106
        
    End Enum

    
    Public Enum enuSourceString
        Saankhya = 1
        Sanchaya = 2
        SanchayaLite = 3
        SaankhyaMasters = 4
        Sthapana = 5
        SOOCHIKA = 6
        KMBR = 7
        SevanaPension = 8
        Sulekha = 9
        SevanaCommon = 10
        SevanaKiosk = 11
        SevanaRegn = 12
        
        Sugama = 16 ' Added by Biji
        SoochikaUnicode = 17 'Added By Akheel
        
        SaankhyaBackUp = 93
        SaankhyaHO = 94
        SanchayaHO = 95
        SaankhyaOld = 96
        Sahatha = 97
        iSaankhyaMasters = 98
        DBMaster = 99
    End Enum

    
    
    Public Enum MonthID
        Apr = 1
        May = 2
        Jun = 3
        Jul = 4
        Aug = 5
        Sep = 6
        Oct = 7
        Nov = 8
        Dec = 9
        Jan = 10
        Feb = 11
        Mar = 12
    End Enum
    
    Public Enum UserType
        Developer = 0
        Administrator = 1
        Approver = 2
        AccountsOfficer = 3
        Operator = 4
    End Enum
    
    Public Enum InstrumentType
        Cash = 1
        TreasuryChalan = 2
        PostalOrder = 3
        DemandDraft = 4
        Cheque = 5
        LetterOfAuthority = 6
        TreasuryBill = 7
        BankPayinSlip = 8
        DirectlyCredited = 9
        DirectlyDebited = 10
    End Enum
    
     Public Type uAcc
        LBCode        As Variant
        HeadCode          As Variant
        Amount  As Variant
    End Type
    
    Public Type uTr
        intTransactionID        As Variant
        intLocalBodyID          As Variant
        intFinancialYearID      As Variant
        dtTransactionDate       As Variant
        intExternalApplicationID        As Variant
        intExternalApplicationModuleID  As Variant
        intFunctionID           As Variant
        intFunctionaryID        As Variant
        intFieldID              As Variant
        intFundID               As Variant
        intBudgetCentreID       As Variant
        vchNarration            As Variant
        intTransactionTypeID    As Variant
        intProcessID            As Variant
        vchGroup                As Variant
        intGroupID              As Variant
        intKeyID                As Variant
        numSubLedgerID          As Variant
        numUserID               As Variant
        intVoucherID            As Variant
        
        '-----Added by Anju on 26-05-2014---
        
        intVoucherNo            As Variant
        tnyStatus               As Variant
        tnyVoucherGroupID       As Variant
        tnyReversed             As Variant
        dtValueDate             As Variant
        '.intTransactionID        =
        '.intLocalBodyID          =
        '.intFinancialYearID      =
        '.dtTransactionDate       =
        '.intExternalApplicationID        =
        '.intExternalApplicationModuleID  =
        '.intFunctionID           =
        '.intFunctionaryID        =
        '.intFieldID              =
        '.intFundID               =
        '.intBudgetCentreID       =
        '.vchNarration            =
        '.intTransactionTypeID    =
        '.intProcessID            =
        '.vchGroup                =
        '.intGroupID              =
        '.intKeyID                =
        '.numSubLedgerID          =
        '.numUserID               =
        '.intVoucherID            =
        
    End Type
    
    Public Type uTrChild
        intTransactionID        As Variant
        intSerialNo             As Variant
        intAccountHeadID        As Variant
        fltAmount               As Variant
        tinDebitOrCreditFlag    As Variant
        intByAccountHeadID      As Variant
        vchNarration            As Variant
        intFundID               As Variant
        
        ' Added by Anju on 26-05-2014
        
        fltOpeningBalance       As Variant
        numTockenID             As Variant
        dtReconcileDate         As Variant
        '.intTransactionID        =
        '.intSerialNo             =
        '.intAccountHeadID        =
        '.fltAmount               =
        '.tinDebitOrCreditFlag    =
        '.intByAccountHeadID      =
        '.vchNarration            =
        '.intFundID               =
    End Type
    
    Public Type uVoucher
        intVoucherID_1          As Variant
        intLocalBodyID_2        As Variant
        intTransactionID_3      As Variant
        intTransactionTypeID_4  As Variant
        tnyVoucherTypeID_5      As Variant
        intVoucherNo_6          As Variant
        intBookNo_7             As Variant
        dtDate_8                As Variant
        fltAmount_9             As Variant
        intInstrumentTypeID_10  As Variant
        vchInstrumentNo_11      As Variant
        dtInstrumentDate_12     As Variant
        vchDescription_13       As Variant
        numZoneID_14            As Variant
        numWardID_15            As Variant
        intDoorNoP1_16          As Variant
        vchDoorNoP2_17          As Variant
        vchDoorNoP3_18          As Variant
        intUserID_19            As Variant
        intCounterID_20         As Variant
        numSubLedgerID_21       As Variant
        intKeyID1_22            As Variant
        intKeyID2_23            As Variant
        intExternalApplicationID_24   As Variant
        intExternalModuleID_25    As Variant
        intFinancialYearID_26     As Variant
        tnyShiftID_27           As Variant
        tnyPrintFlag_28         As Variant
        tnyCancelFlag_29        As Variant
        vchBank_33              As Variant
        vchBankPlace_34         As Variant
        intFundID_35            As Variant
        numSeatID               As Variant
        intSessionID            As Variant
        vchRefNo                As Variant
        fltRoundOff             As Variant
        fltAdvAmtAdj            As Variant
        numInwardNo             As Variant
        tnyStatus_32            As Variant
        numLocationID           As Variant
        
        
        '-----Added by Anju on 26-05-2014
        dtRealisationDate       As Variant
        vchRemarks              As Variant
        tnyReconciled           As Variant
        numTockenID             As Variant
        tnyVoucherGroupID       As Variant
        numLinkKeyID            As Variant
        dtTimeStamp             As Variant
        dtChequeRealiseDate     As Variant
        vchVersionKey           As Variant
        tnyReversed             As Variant
        dtValueDate             As Variant
        
        '.intVoucherID_1          =
        '.intLocalBodyID_2        =
        '.intTransactionID_3      =
        '.intTransactionTypeID_4  =
        '.tnyVoucherTypeID_5      =
        '.intVoucherNo_6          =
        '.intBookNo_7             =
        '.dtDate_8                =
        '.fltAmount_9             =
        '.intInstrumentTypeID_10  =
        '.vchInstrumentNo_11      =
        '.dtInstrumentDate_12     =
        '.vchDescription_13       =
        '.numZoneID_14            =
        '.numWardID_15            =
        '.intDoorNoP1_16          =
        '.vchDoorNoP2_17          =
        '.vchDoorNoP3_18          =
        '.intUserID_19            =
        '.intCounterID_20         =
        '.numSubLedgerID_21       =
        '.intKeyID1_22            =
        '.intKeyID2_23            =
        '.intExternalApplicationID_24   =
        '.intExternalModuleID_25    =
        '.intFinancialYearID_26     =
        '.tnyShiftID_27           =
        '.tnyPrintFlag_28         =
        '.tnyCancelFlag_29        =
        '.vchBank_33              =
        '.vchBankPlace_34         =
        '.intFundID_35            =
        '.numSeatID               =
        '.intSessionID            =
        '.vchRefNo                =
        '.fltRoundOff             =
        '.fltAdvAmtAdj            =
        '.numInwardNo             =
        '.tnyStatus_32            =
        '.numLocationID           =
    End Type
    
    Public Type uVChild
        intVoucherID_1      As Variant
        intLocalBodyID_2    As Variant
        intSlNo_3           As Variant
        intAccountHeadID_4  As Variant
        tnyDebitOrCredit_5  As Variant
        intYearID_6         As Variant
        tnyPeriodID_7       As Variant
        tnyArrearFlag_8     As Variant
        numDemandID_9       As Variant
        fltAmount_10        As Variant
        
        '.intVoucherID_1     =
        '.intLocalBodyID_2   =
        '.intSlNo_3          =
        '.intAccountHeadID_4 =
        '.tnyDebitOrCredit_5 =
        '.intYearID_6        =
        '.tnyPeriodID_7      =
        '.tnyArrearFlag_8    =
        '.numDemandID_9      =
        '.fltAmount_10       =
    End Type
    
    Public Type uVoucherAddress
        intVoucherID    As Variant
        intLocalBodyID  As Variant
        vchName         As Variant
        vchInit1        As Variant
        vchInit2        As Variant
        vchInit3        As Variant
        vchInit4        As Variant
        vchHouseName    As Variant
        vchStreetName   As Variant
        vchLocalPlace   As Variant
        vchMainPlace    As Variant
        vchPostOffice   As Variant
        vchDistrict     As Variant
        vchPinNumber    As Variant
        vchPhone        As Variant
        intWardNo       As Variant
        intDoorNo       As Variant
        vchDoorNo2      As Variant
        
        
        '.intVoucherID    =
        '.intLocalBodyID  =
        '.vchName         =
        '.vchInit1        =
        '.vchInit2        =
        '.vchInit3        =
        '.vchInit4        =
        '.vchHouseName    =
        '.vchStreetName   =
        '.vchLocalPlace   =
        '.vchMainPlace    =
        '.vchPostOffice   =
        '.vchDistrict     =
        '.vchPinNumber    =
        '.vchPhone        =
        '.intWardNo       =
        '.intDoorNo       =
        '.vchDoorNo2      =
    End Type
    
    Public Type uVoucherSub
        intVoucherID      As Variant
        decProjectID      As Variant
        intSourceOfFundID As Variant
        intCategoryID     As Variant
        intSectorID       As Variant
        intAllotmentID    As Variant
        intAgreementID    As Variant
        intCashBookID     As Variant
        intImplementingOfficerID  As Variant
        intCreditorTypeID As Variant
        intCreditorsID    As Variant
        intTypeID         As Variant
        intLocalBodyID    As Variant
        
        '.intVoucherID      =
        '.decProjectID      =
        '.intSourceOfFundID =
        '.intCategoryID     =
        '.intSectorID       =
        '.intAllotmentID    =
        '.intAgreementID    =
        '.intCashBookID     =
        '.intImplementingOfficerID  =
        '.intCreditorTypeID =
        '.intCreditorsID    =
        '.intTypeID         =
        '.intLocalBodyID    =
        
    End Type
    
    
    Public Type uPaymentOrder
        intPayOrderID           As Variant
        vchPayOrderNo           As Variant
        dtPayOrderDate          As Variant
        dtDueDate               As Variant
        intFunctionaryID        As Variant
        intFunctionID           As Variant
        intTransactionTypeID    As Variant
        vchBillNo               As Variant
        numBillAmount           As Variant
        dtBillDate              As Variant
        intInstrumentTypeID     As Variant
        intCashOrBankHeadID     As Variant
        vchDescription          As Variant
        vchTitle                As Variant
        intSubLedgerTypeID      As Variant
        intPayToSubLedgerID     As Variant
        intSubsidiaryCashBookID As Variant
        intImplementingOfficerID As Variant
        numProjectNo            As Variant
        intStockRegisterID      As Variant
        vchStockRefNo           As Variant
        intAssetTypeID          As Variant
        intAssetID              As Variant
        numFwdSeatID            As Variant
        intLocalBodyID          As Variant
        intZonalID              As Variant
        intFinancialYearID      As Variant
        numUserID               As Variant
        numSeatID               As Variant
        numApprovingOfficerID   As Variant
        numApprovingSeatID      As Variant
        dtApprovingDate         As Variant
        intVoucherID            As Variant
        intVoucherNo            As Variant
        dtVoucherDate           As Variant
        tnyStatus               As Variant
        
        intKeyID                As Variant
        numKeyID                As Variant
        dtKeyDate               As Variant
        
        tnyCancelled            As Variant
        intAppID                As Variant
        intModuleID             As Variant
        
        intSourceOfFundID       As Variant
        intAllotmentID          As Variant
        intAgreementID          As Variant
        tnyCategoryID           As Variant
        tnySectorID             As Variant
        tnyIsFinalBill          As Variant
        
        
        '.intPayOrderID           =
        '.vchPayOrderNo           =
        '.dtPayOrderDate          =
        '.dtDueDate               =
        '.intFunctionaryID        =
        '.intFunctionID           =
        '.intTransactionTypeID    =
        '.vchBillNo               =
        '.numBillAmount           =
        '.dtBillDate              =
        '.intInstrumentTypeID     =
        '.intCashOrBankHeadID     =
        '.vchDescription          =
        '.vchTitle                =
        '.intSubLedgerTypeID      =
        '.intPayToSubLedgerID     =
        '.intSubsidiaryCashBookID =
        '.intImplementingOfficerID =
        '.numProjectNo            =
        '.intStockRegisterID      =
        '.vchStockRefNo           =
        '.intAssetTypeID          =
        '.intAssetID              =
        '.numFwdSeatID            =
        '.intLocalBodyID          =
        '.intZonalID              =
        '.intFinancialYearID      =
        '.numUserID               =
        '.numSeatID               =
        '.numApprovingOfficerID   =
        '.numApprovingSeatID      =
        '.dtApprovingDate         =
        
        '.intVoucherID            =
        '.intVoucherNo            =
        '.dtVoucherDate           =
        '.tnyStatus               =
        
        '.intKeyID                =
        '.numKeyID                =
        '.dtKeyDate               =
        
        '.tnyCancelled            =
        '.intAppID                =
        '.intModuleID             =
        
        '.intSourceOfFundID       =
        '.intAllotmentID          =
        '.intAgreementID          =
        '.tnyCategoryID           =
        '.tnySectorID             =
        '.tnyIsFinalBill          =
        
    End Type
    
    Public Type uPaymentOrderChild
        intPayOrderID           As Variant
        intSlNo                 As Variant
        intAccountHeadID        As Variant
        vchAccountHeadCode      As Variant
        numAmount               As Variant
        tnyCategoryFlag         As Variant
        tnyDebitOrCreditFlag    As Variant
        vchDescription          As Variant
        tnyExcldeFromSourceFlag As Variant
        '.intPayOrderID         =
        '.intSlNo               =
        '.intAccountHeadID      =
        '.vchAccountHeadCode    =
        '.numAmount             =
        '.tnyCategoryFlag       =
        '.tnyDebitOrCreditFlag  =
        '.vchDescription        =
    End Type
    
    Public Type uPaymentOrderAddress
        intPayOrderID       As Variant
        intSubsidiaryAccountHeadID As Variant
        intSubLegerTypeID   As Variant
        vchSubLedgerCode    As Variant
        vchName             As Variant
        vchHouseName        As Variant
        vchStreet           As Variant
        vchLocalPlace       As Variant
        vchMainPlace        As Variant
        vchPost             As Variant
        vchPinCode          As Variant
        vchPhone            As Variant
        
        
        '.intPayOrderID       =
        '.intSubsidiaryAccountHeadID =
        '.intSubLegerTypeID   =
        '.vchSubLedgerCode    =
        '.vchName             =
        '.vchHouseName        =
        '.vchStreet           =
        '.vchLocalPlace       =
        '.vchMainPlace        =
        '.vchPost             =
        '.vchPinCode          =
        '.vchPhone            =
    End Type
    
    Public Type uProject
        decProjectID    As Variant
        intLBID         As Variant
        intYearID       As Variant
        intProjectSlNo  As Variant
        chvProjectSlNo  As Variant
        chvProjectName  As Variant
        chvProjectnameEnglish As Variant
        intProjCatID    As Variant
        chvDPCOrderNo   As Variant
        dtDPCOrderDate  As Variant
        intSectorTypeID As Variant
        intPlanID       As Variant
        intSourceOfFundID As Variant
        fltEstSourceAmt As Variant
        
        '.decProjectID
        '.intLBID
        '.intYearID
        '.intProjectSlNo
        '.chvProjectSlNo
        '.chvProjectName
        '.chvProjectnameEnglish
        '.intProjCatID
        '.chvDPCOrderNo
        '.dtDPCOrderDate
        '.intSectorTypeID
        '.intPlanID
        '.intSourceOfFundID
        '.fltEstSourceAmt
    End Type
    
    Public Type uDemand
        numDemandID As Variant
        intLBID As Variant
        tnyExtAppID As Variant
        tnyExtModuleID As Variant
        tnyDemandType As Variant
        intTransactionTypeID As Variant
        intYearID As Variant
        tnyPeriodID As Variant
        dtDemandDate As Variant
        numSubLedgerID As Variant
        intKeyID As Variant
        intKeyID2 As Variant
        vchRemarks As Variant
        tnyStatus As Variant
        tnyArrearFlag As Variant
        intVoucherID As Variant
        dtVoucherDate As Variant
        dtExpiryDate As Variant
        intFinancialYearID As Variant
        numSeatID As Variant
        intSectionID As Variant
        numUserID As Variant
        numCounterID As Variant
        vchAdminNote As Variant
        vchDemandNo As Variant
        numZoneID As Variant
        intWardNo As Variant
        intDoorNo As Variant
        vchDoorNo2 As Variant
        numForwardedSeatID As Variant
        intInstrumentTypeID As Variant
        vchInstrumentNo As Variant
        dtInstrumentDate As Variant
        vchDrawnFrom As Variant
        vchDrawnPlace As Variant
        dtDueDate As Variant
        tnyAccrualType As Variant
        numLocationID As Variant
        tnySend As Variant
        intFunctionID As Variant
        intFunctionaryID As Variant
        intSourceFundID As Variant
        dtTransactionDate As Variant 'Added On 4th Jul 2011
        intDemandMode   As Variant
        '.numDemandID =
        '.intLBID =
        '.tnyExtAppID =
        '.tnyExtModuleID =
        '.tnyDemandType =
        '.intTransactionTypeID =
        '.intYearID =
        '.tnyPeriodID =
        '.dtDemandDate =
        '.numSubLedgerID =
        '.intKeyID =
        '.intKeyID2 =
        '.vchRemarks =
        '.tnyStatus =
        '.tnyArrearFlag =
        '.intVoucherID =
        '.dtVoucherDate =
        '.dtExpiryDate =
        '.intFinancialYearID =
        '.numSeatID =
        '.intSectionID =
        '.numUserID =
        '.numCounterID =
        '.vchAdminNote =
        '.vchDemandNo =
        '.numZoneID =
        '.intWardNo =
        '.intDoorNo =
        '.vchDoorNo2 =
        '.numForwardedSeatID =
        '.intInstrumentTypeID =
        '.vchInstrumentNo =
        '.dtInstrumentDate =
        '.vchDrawnFrom =
        '.vchDrawnPlace =
        '.dtDueDate =
        '.tnyAccrualType =
        '.numLocationID =
        '.tnySend =
        '.intFunctionID =
        '.intFunctionaryID =
        '.intSourceFundID =
    End Type
    
    Public Type uDemandChild
        numDemandID As Variant
        intLBID As Variant
        tnySlNo As Variant
        intAccountHeadID As Variant
        vchAccountHeadCode As Variant
        fltAmount As Variant
        intYearID As Variant
        tnyPeriodID As Variant
        tnyArrearFlag As Variant
        vchRemarks As Variant
        tnyStatus As Variant
        dtOnDate As Variant
        snyRate As Variant
        intVoucherID As Variant
        dtVoucherDate As Variant
        intTransactionTypeID As Variant
'        .numDemandID =
'        .intLBID =
'        .tnySlNo =
'        .intAccountHeadID =
'        .vchAccountHeadCode =
'        .fltAmount =
'        .intYearID =
'        .tnyPeriodID =
'        .tnyArrearFlag =
'        .vchRemarks =
'        .tnyStatus =
'        .dtOnDate =
'        .snyRate =
'        .intVoucherID =
'        .dtVoucherDate =
'        .intTransactionTypeID
    End Type
    
    Public Type uDemandAddress
        numDemandID As Variant
        numZoneID As Variant
        intWardNo As Variant
        intDoorNo As Variant
        vchDoorNo2 As Variant
        vchName As Variant
        vchInit1 As Variant
        vchInit2 As Variant
        vchInit3 As Variant
        vchInit4 As Variant
        vchHouseName As Variant
        vchStreet As Variant
        vchLocalPlace As Variant
        vchMainPlace As Variant
        vchPost As Variant
        vchPin As Variant
        vchPhone As Variant
        
        'numDemandID =
        'numZoneID =
        'intWardNo =
        'intDoorNo =
        'vchDoorNo2 =
        'vchName =
        'vchInit1 =
        'vchInit2 =
        'vchInit3 =
        'vchInit4 =
        'vchHouseName =
        'vchStreet =
        'vchLocalPlace =
        'vchMainPlace =
        'vchPost =
        'vchPin =
        'vchPhone =
    End Type
    
    Public Type uRequisition
        tnyStage              As Variant
        vchRequisition        As Variant
        dtRequisitionDate         As Variant
        intImplementingOfficersID   As Variant
        vchDesignation        As Variant
        vchNameofIMPO         As Variant
        vchPlace              As Variant
        vchDepartment         As Variant
        vchDDOCode            As Variant
        fltRequestedAmt       As Variant
        tnyPlanOrNonPlan      As Variant
        numProjectID          As Variant
        numProjectNo          As Variant
        fltProjectCost        As Variant
        vchDPCApprovalNo      As Variant
        dtDPCDate             As Variant
        intSourceID           As Variant
        intCategoryID         As Variant
        
        intTreasuryID         As Variant
        vchTreasuryCode       As Variant
        vchTreasuryName       As Variant
        vchGHeadofAccount     As Variant
        vchGBudgetHead        As Variant
        vchGDemandNo          As Variant
        
        intFunctionaryID      As Variant
        intFunctionID         As Variant
        intAccountHeadID      As Variant
        vchAccountHeadCode    As Variant
        
        intLBID               As Variant
        intFinancialYearID    As Variant
        tnyStatus             As Variant
        
        tnyInstallmentNo      As Variant
        intSchemeID           As Variant
        intSubSecID           As Variant
        intMircoSectorID      As Variant
        tnyTypeID             As Variant
        vchNatureOfClaim      As Variant       ' MODIFIED ON 07.Sep.2015
        
        '.tnyStage          = Null
        '.dtRequisition     = Null
        '.intImplementingOfficersID = Null
        '.vchDesignation    = Null
        '.vchNameofIMPO     = Null
        '.vchPlace          = Null
        '.vchDepartment     = Null
        '.vchDDOCode        = Null
        '.fltRequestedAmt   = Null
        '.tnyPlanOrNonPlan  = Null
        '.numProjectID      = Null
        '.numProjectNo      = Null
        '.fltProjectCost    = Null
        '.vchDPCApprovalNo  = Null
        '.dtDPCDate         = Null
        '.intSourceID       = Null
        '.intCategoryID     = Null
        
        '.intTreuryID       = Null
        '.vchTreuryCode     = Null
        '.vchTreuryName     = Null
        '.vchGHeadofAccount = Null
        '.vchGBudgetHead    = Null
        '.vchGDemandNo      = Null
        
        '.intFunctionaryID  = Null
        '.intFunctionID     = Null
        '.intAccountHeadID  = Null
        '.vchAccountHeadCode = Null
        
        '.intLBID           = Null
        '.intFinancialYearID = Null
        '.tnyStatus         = Null
    
    End Type
    
    
    Public Type uFund
        intSourceOfFundID   As Variant
        strSourceCode       As Variant
        strSourceName       As Variant
        intAllocatedYearID  As Variant
        intSlNo             As Variant
        fltSourceWiseAmount As Variant
        fltSourceWiseUtilisedAmount As Variant
    End Type
    
    
    
