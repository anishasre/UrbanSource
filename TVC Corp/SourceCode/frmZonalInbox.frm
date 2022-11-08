VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmZonalInbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DemandInbox"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   LinkTopic       =   "DemandType"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraZInbox 
      Height          =   1140
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8340
      Begin MSComCtl2.DTPicker dtPkrToDate 
         Height          =   375
         Left            =   900
         TabIndex        =   11
         Top             =   585
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   62128131
         CurrentDate     =   41281
      End
      Begin MSComCtl2.DTPicker dtPkrFromDate 
         Height          =   330
         Left            =   900
         TabIndex        =   10
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   62128131
         CurrentDate     =   41281
      End
      Begin VB.ComboBox cmbZonal 
         Height          =   315
         Left            =   4860
         TabIndex        =   4
         Text            =   "Zonal"
         Top             =   675
         Width           =   2310
      End
      Begin VB.ComboBox cmbDemandType 
         Height          =   315
         ItemData        =   "frmZonalInbox.frx":0000
         Left            =   4860
         List            =   "frmZonalInbox.frx":0002
         TabIndex        =   3
         Text            =   "DEMAND"
         Top             =   270
         Width           =   2310
      End
      Begin VB.Label lblToDate 
         Caption         =   "ToDate"
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   675
         Width           =   690
      End
      Begin VB.Label lblFromDate 
         Caption         =   "FromDate"
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   780
      End
   End
   Begin VB.Frame fraChildInbox 
      Caption         =   "Collection details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   45
      TabIndex        =   5
      Top             =   1260
      Width           =   8385
      Begin VB.CommandButton cmdReciept 
         Caption         =   "Generate Reciept"
         Height          =   330
         Left            =   6300
         TabIndex        =   12
         Top             =   2160
         Width           =   1860
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridInboxCash 
         Height          =   1635
         Left            =   180
         TabIndex        =   6
         Top             =   450
         Width           =   8070
         _cx             =   14235
         _cy             =   2884
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalInbox.frx":0004
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridInboxNoncash 
         Height          =   1635
         Left            =   180
         TabIndex        =   13
         Top             =   2565
         Width           =   8025
         _cx             =   14155
         _cy             =   2884
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalInbox.frx":0244
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   9
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label lblNonCash 
         BackStyle       =   0  'Transparent
         Caption         =   "Non Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   2295
         Width           =   1050
      End
      Begin VB.Label lblSelectedDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   8370
         TabIndex        =   7
         Top             =   180
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmZonalInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mDemandType As Integer
    Dim mReceiptNo  As Double
    Dim mVoucherID As Double
    Dim mCashTotal As Double
    Dim mFlag As Integer


Private Sub FillInboxGrid()
    Dim FromDate As Date
    Dim ToDate As Date
    Dim Location As Double
    Dim DemandMode As Variant
    Dim arrInput As Variant
    Dim mCnn As New ADODB.Connection
    Dim mCnnSvr As New ADODB.Connection
    Dim msqlvN As String
    Dim msqlNonCash As String
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim RecAcc As New ADODB.Recordset
    Dim RecNonCash As New ADODB.Recordset
    Dim RecvN As New ADODB.Recordset
    Dim mRowCount As Integer
    Dim mArrayOut As Variant
    Dim objFunction As New clsFunction
    Dim intFunID As Integer
    Dim mSql, valcmb As String
    Dim intID As Integer
    
    FromDate = dtPkrFromDate.value
    ToDate = dtPkrToDate.value
    Location = cmbZonal.ItemData(cmbZonal.ListIndex) '4016704   'cmbZonal.Index
    If Not cmbDemandType.ListIndex <= 0 Then
        DemandMode = cmbDemandType.ListIndex - 2 'cmbDemandType.ListIndex - 2
    Else
        vsGridInboxCash.Clear 1, 1
        Exit Sub
    End If
    If DemandMode = 2 Then  '----Zonal collection
        Set arrInput = Nothing
        arrInput = Array(FromDate, ToDate, Location, DemandMode)
        objdb.CreateNewConnection mCnn, enuSourceString.SaankhyaHO
        objdb.CreateNewConnection mCnnSvr, enuSourceString.Saankhya
        Set Rec = objdb.ExecuteSP("spGetZoneDemand", arrInput, , , mCnn, adCmdStoredProc)
        If Not Rec.EOF Then
            While Not Rec.EOF
                mRowCount = mRowCount + 1
                vsGridInboxCash.TextMatrix(mRowCount, 1) = mRowCount
                intFunID = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                objFunction.SetFunctionByID val(intFunID)
                vsGridInboxCash.TextMatrix(mRowCount, 3) = objFunction.FunctionName
                vsGridInboxCash.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsGridInboxCash.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                mCashTotal = mCashTotal + Rec!fltAmount
                vsGridInboxCash.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
                vsGridInboxCash.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                mSql = "Select * From faIDemandChild  WHERE numDemandID=" & vsGridInboxCash.TextMatrix(mRowCount, 16)
                RecAcc.Open mSql, mCnn
                If Not Rec.EOF Then
                    vsGridInboxCash.TextMatrix(mRowCount, 13) = IIf(IsNull(RecAcc!intAccountHeadID), "", RecAcc!intAccountHeadID)
                    If Rec!intVoucherID <> 0 Then
                        msqlvN = "select * from faVouchers where intVoucherID=" & Rec!intVoucherID
                        RecvN.Open msqlvN, mCnnSvr
                        If Not RecvN.EOF Then
                            vsGridInboxCash.TextMatrix(mRowCount, 7) = IIf(IsNull(RecvN!intVoucherNo), "", RecvN!intVoucherNo)
                            cmdReciept.Enabled = False
                            RecvN.Close
                        Else
                            cmdReciept.Enabled = True
                            
                        End If
                        
                    Else
                        cmdReciept.Enabled = True
                    End If
                    RecAcc.MoveNext
                 End If
                 
                 RecAcc.Close
                 Rec.MoveNext
            Wend
            '-----NonCash---------
            Dim mRowCountNon As Integer
            Dim arrIn  As Variant
            arrIn = Array(FromDate, ToDate)
            Set RecNonCash = objdb.ExecuteSP("spZoneNonCash", arrIn, , , mCnn, adCmdStoredProc)
            If Not RecNonCash.EOF Then
                While Not RecNonCash.EOF
                    mRowCountNon = mRowCountNon + 1
                    vsGridInboxNoncash.TextMatrix(mRowCountNon, 1) = mRowCountNon
                    vsGridInboxNoncash.TextMatrix(mRowCountNon, 2) = IIf(IsNull(RecNonCash!vchTransactionType), "", RecNonCash!vchTransactionType)
                    vsGridInboxNoncash.TextMatrix(mRowCountNon, 3) = IIf(IsNull(RecNonCash!vchInstrumentType), "", RecNonCash!vchInstrumentType)
                    vsGridInboxNoncash.TextMatrix(mRowCountNon, 4) = IIf(IsNull(RecNonCash!intVoucherNo), "", RecNonCash!intVoucherNo)
                    vsGridInboxNoncash.TextMatrix(mRowCountNon, 5) = IIf(IsNull(RecNonCash!fltAmount), "", RecNonCash!fltAmount)
                    RecNonCash.MoveNext
                Wend
            End If
            RecNonCash.Close
        Else
            vsGridInboxCash.Clear 1, 1
            vsGridInboxNoncash.Clear 1, 1
            cmdReciept.Enabled = True
        End If
    Else '------epayement
        MsgBox ("Select ZonalCollection/e-Payment")
        vsGridInboxCash.Clear 1, 1
        Exit Sub
    End If
End Sub
Private Sub cmbZonal_Click()
    If cmbZonal.ListIndex <> -1 Then
        Call FillInboxGrid
    End If
End Sub

Private Sub cmdReciept_Click()
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mCnnSvr As New ADODB.Connection
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    Dim mSqlVChild As String
    Dim RecVChild As New ADODB.Recordset
    Dim mSqlUpadate As String
    Dim RecUpadate As New ADODB.Recordset
    Dim msqlStatus As String
    Dim RecStatus As New ADODB.Recordset
        
        '---------save to faVouchers-------------
        
    Dim intVoucherID As Double
    Dim intVoucherIDN As Integer
    Dim intLocalBodyID As Integer
    Dim intTransactionIDN As Integer
    Dim intTransactionTypeIDN  As Integer
    Dim tnyVoucherTypeID As Integer
    Dim intVoucherNo As Variant
    Dim intVoucherNoN As Variant
    Dim intBookNo   As Integer
    Dim dtDate As Date
    Dim fltAmount As Variant
    Dim intInstrumentTypeID As Integer
    Dim vchInstrumentNo As String
    Dim dtInstrumentDate As Variant
    Dim vchDescription As String
    Dim numZoneID As Double
    Dim numWardId As Variant
    Dim intDoorNoP1 As Variant
    Dim vchDoorNoP2 As Variant
    Dim vchDoorNoP3 As Variant
    Dim intUserID As Variant
    Dim intCounterID As Integer
    Dim numSubLedgerID As Variant
    Dim intKeyID1 As Variant
    Dim intKeyID2 As Variant
    Dim intExternalApplicationID As Integer
    Dim intExternalModuleID As Integer
    Dim intFinancialYearID  As Integer
    Dim tnyShiftID As Integer
    Dim tnyPrintFlag As Integer
    Dim tnyCancelFlag As Integer
    Dim dtRealisationDate As Variant
    Dim vchRemarks As String
    Dim tnyStatus As String
    Dim vchBank As String
    Dim vchBankPlace As String
    Dim intFundID   As Variant
    Dim numSeatID As Double
    Dim intSessionID As Variant
    
    Dim vchRefNo As Variant
    Dim fltRoundOff As Variant
    Dim tnyReconciled As Variant
    Dim numTockenID As Variant
    Dim fltAdvAmtAdj As Variant
    Dim numInwardNo As Variant
    Dim numLocationID As Variant
    Dim tnyVoucherGroupID As Variant
    Dim numLinkKeyID As Variant
    Dim dtTimeStamp As Variant
    Dim dtChequeRealiseDate As Variant
    Dim vchVersionKey As Variant
    Dim tnyReversed As Variant
    Dim dtValueDate As Variant
    Dim count As Integer
    Dim countDemand As Integer
    Dim mDemandID As Variant
    Dim intTransactionID As Variant
    Dim intTransactionTypeID As Integer
    Dim mRowCount As Integer
    
    If mFlag = 1 Then
        cmdReciept.Enabled = False
    Else
        mRowCount = 1
        objdb.CreateNewConnection mCnn, SaankhyaHO
        mSql = "Select * From faIDemandTbl  WHERE tnyExtModuleID=35 AND intDemandMode=2 AND dtDemandDate between '" & dtPkrFromDate & " ' AND '" & dtPkrToDate & "'"
        Rec.Open mSql, mCnn
    '-----Save FaVouchers-------------
        If Not (Rec.BOF And Rec.EOF) Then
            intVoucherIDN = -1
            intTransactionIDN = 0
            tnyVoucherTypeID = 10
            intVoucherNoN = Null
            intBookNo = 0
            fltAmount = mCashTotal
            numWardId = Null
            vchDoorNoP2 = Null
            vchDoorNoP3 = Null
            intUserID = gbUserID
            numSubLedgerID = 0
            tnyShiftID = gbShiftID
            tnyPrintFlag = 1
            tnyCancelFlag = 0
            dtRealisationDate = Null
            tnyStatus = Rec!tnyStatus
            vchBank = 0
            vchBankPlace = "0"
            intFundID = 1
            numSeatID = gbSeatID
            intSessionID = gbSessionID
            vchRefNo = 0
            fltRoundOff = 0
            tnyReconciled = 0
            numTockenID = 0
            fltAdvAmtAdj = 0
            numInwardNo = 0
            tnyVoucherGroupID = 0
            numLinkKeyID = 0
            dtTimeStamp = 0
            dtChequeRealiseDate = 0
            vchVersionKey = 0
            tnyReversed = 0
            dtValueDate = 0
            
            While Not Rec.EOF
                mDemandID = Rec!numDemandID
                intDoorNoP1 = Rec!intDoorNo
                vchDescription = Rec!vchRemarks
                numZoneID = Rec!numLocationID
                dtDate = Rec!dtDemandDate
                intKeyID1 = Rec!intKeyID
                intKeyID2 = Rec!intKeyID2
                intLocalBodyID = Rec!intLBID
                intTransactionTypeIDN = Rec!intTransactionTypeID
                intInstrumentTypeID = Rec!intInstrumentTypeID
                vchInstrumentNo = Rec!vchInstrumentNo
                dtInstrumentDate = Rec!dtInstrumentDate
                intCounterID = Rec!numCounterID
                intExternalApplicationID = Rec!tnyExtAppID
                intExternalModuleID = Rec!tnyExtModuleID
                intFinancialYearID = Rec!intFinancialYearID
                numLocationID = Rec!numLocationID
                vchRemarks = Rec!vchRemarks
                   
                arrInput = Array(intVoucherIDN, intLocalBodyID, _
                    Null, intTransactionTypeIDN, _
                    tnyVoucherTypeID, intVoucherNoN, _
                    intBookNo, dtDate, fltAmount, intInstrumentTypeID, _
                    vchInstrumentNo, dtInstrumentDate, vchDescription, numZoneID, _
                    numWardId, intDoorNoP1, vchDoorNoP2, vchDoorNoP3, _
                    intUserID, intCounterID, numSubLedgerID, intKeyID1, intKeyID2, _
                    intExternalApplicationID, intExternalModuleID, intFinancialYearID, _
                    tnyShiftID, tnyPrintFlag, tnyCancelFlag, _
                    vchBank, vchBankPlace, intFundID, numSeatID, _
                    intSessionID, vchRefNo, fltRoundOff, fltAdvAmtAdj, _
                    numInwardNo, tnyStatus, numLocationID, tnyVoucherGroupID, numLinkKeyID)
                    
                    Set arrOutPut = Nothing
                    objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnnSvr
                
                    If IsNumeric(arrOutPut(0, 0)) Then
                        intVoucherID = arrOutPut(0, 0)
                        mVoucherID = intVoucherID
                        mReceiptNo = arrOutPut(1, 0)
                    End If
                                        
                    vsGridInboxCash.TextMatrix(mRowCount, 7) = mReceiptNo
                    mRowCount = mRowCount + 1
                    
                    '------Update faIDemandTbl---------
                    objdb.CreateNewConnection mCnn, SaankhyaHO
                    mSqlUpadate = "update  faIDemandTBL set intVoucherID=" & intVoucherID & "where numDemandID=" & mDemandID
                    RecUpadate.Open mSqlUpadate, mCnn
                                          
                   '----Save VoucherChild---------
              
                    Dim tnySlNo As Variant
                    Dim intAccountHeadID   As Variant
                    Dim tnyDebitOrCredit As Variant
                    Dim intYearID   As Variant
                    Dim tnyPeriodID As Variant
                    Dim tnyArrearFlag As Variant
                    Dim mNumDemandID As Variant
                    Dim RecN As New ADODB.Recordset
                    Dim Recv As New ADODB.Recordset
                    Dim mSqlV As String
                    Dim mNumDemandIDGet As Variant
    
                    mSqlVChild = "select * from faIDemandChild where numDemandID =" & mDemandID
                    RecVChild.Open mSqlVChild, mCnn
    
                    While Not RecVChild.EOF
                        If Not (RecVChild.BOF And RecVChild.EOF) Then
                            tnySlNo = RecVChild!tnySlNo
                            intAccountHeadID = RecVChild!intAccountHeadID
                            tnyDebitOrCredit = 0
                            tnyPeriodID = 1
                            intYearID = gbFinancialYearID
                            tnyArrearFlag = 0
                            mNumDemandID = RecVChild!numDemandID
                            fltAmount = RecVChild!fltAmount
                        
                            arrInput = Array( _
                            intVoucherID, _
                            gbLocalBodyID, _
                            tnySlNo, _
                            intAccountHeadID, _
                            tnyDebitOrCredit, _
                            intYearID, _
                            tnyPeriodID, _
                            tnyArrearFlag, _
                            mNumDemandID, _
                            fltAmount)
                             
                            objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnnSvr
                            RecVChild.MoveNext
                        End If
                    Wend
                    RecVChild.Close
               
                 '-----save VoucherAddress----
                
                    Dim vchName As Variant
                    Dim vchInit1 As Variant
                    Dim vchInit2 As Variant
                    Dim vchInit3 As Variant
                    Dim vchInit4 As Variant
                    Dim vchHouseName As Variant
                    Dim vchStreetName As Variant
                    Dim vchLocalPlace As Variant
                    Dim vchMainPlace As Variant
                    Dim vchPostOffice As Variant
                    Dim vchDistrict As Variant
                    Dim vchPinNumber As Variant
                    Dim vchPhone As Variant
                    Dim intWardNo As Variant
                    Dim intDoorNo As Variant
                    Dim vchDoorNo2 As Variant
    
                    arrInput = Array(intVoucherID, _
                               gbLocalBodyID, _
                               vchName, _
                               vchInit1, _
                               vchInit2, _
                               vchInit3, _
                               vchInit4, _
                               vchHouseName, _
                                vchStreetName, _
                               vchLocalPlace, _
                               vchMainPlace, _
                               vchPostOffice, _
                               vchDistrict, _
                               vchPinNumber, _
                               vchPhone, _
                               intWardNo, _
                               intDoorNo, _
                               vchDoorNo2)
                        objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnnSvr
 
                '-----save faTransactions------
                    Dim objFunctions As New clsFunction
                    Dim objFunctionaries As New clsFunctionary
                    Dim dtTransactionDate As Variant
                    Dim intExternalApplicationModuleID  As Variant
                    Dim intFunctionID   As Variant
                    Dim intFunctionaryID    As Variant
                    Dim intFieldID  As Variant
                    Dim intBudgetCentreID   As Variant
                    Dim vchNarration As Variant
                    Dim intProcessID    As Variant
                    Dim intGroupID  As Variant
                    Dim vchGroup As String
                    Dim intKeyID    As Variant
                    Dim numUserID As Variant
                    Dim intTransactionIDNew As Variant

                    intTransactionIDNew = -1
                    intLocalBodyID = gbLocalBodyID
                    intFinancialYearID = gbFinancialYearID
                    dtTransactionDate = gbTransactionDate
                    intExternalApplicationID = AppID.Saankhya
                    intExternalApplicationModuleID = 45
                    intFunctionID = objFunctions.FunctionID
                    intFunctionaryID = objFunctionaries.FunctionaryID
                    intFieldID = Null
                    intFundID = Null
                    intBudgetCentreID = Null
                    vchNarration = Null
                    vchNarration = Null
                    intVoucherNo = intVoucherID
                    intProcessID = Null
                    vchGroup = "R"
                    intKeyID = Null
                    intGroupID = 10
                    numSubLedgerID = 1
                    
                    arrInput = Array( _
                     intTransactionIDNew, _
                     intLocalBodyID, _
                     intFinancialYearID, _
                     dtDate, _
                     intExternalApplicationID, _
                     intExternalApplicationModuleID, _
                     intFunctionID, _
                     intFunctionaryID, _
                     intFieldID, _
                     intFundID, _
                     intBudgetCentreID, _
                     vchNarration, _
                     intTransactionTypeIDN, _
                     intProcessID, _
                     vchGroup, _
                     intGroupID, _
                     intKeyID, _
                     numSubLedgerID, _
                     gbUserID, _
                     intVoucherNo, tnyVoucherGroupID)
  
                    Set arrOutPut = Nothing
                    objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnnSvr
                    If IsNumeric(arrOutPut(0, 0)) Then
                        intTransactionID = arrOutPut(0, 0)
                    End If
    
                '------Save TransactionChild----
                    Dim intSerialNo As Variant
                    Dim tinDebitOrCreditFlag As Variant
                    Dim intByAccountHeadID  As Variant
                    Dim fltOpeningBalance As Variant
                    Dim dtReconcileDate As Variant
                    Dim intSlNo As Variant
                    Dim RecTChild As New ADODB.Recordset
                    Dim mSqlTChild As String
                    
                    mSqlTChild = "select * from faIDemandChild where numDemandID =" & mDemandID
                    RecTChild.Open mSqlTChild, mCnn
                    
                    fltAmount = mCashTotal
                    intAccountHeadID = 1504
                    intByAccountHeadID = Null
                    intSlNo = 1
                    tinDebitOrCreditFlag = 1
                    vchNarration = Null
                    intFundID = 1
                    arrInput = Array(intTransactionID, intSlNo, _
                    intAccountHeadID, _
                    fltAmount, _
                    tinDebitOrCreditFlag, _
                    intByAccountHeadID, _
                    vchNarration, _
                    intFundID)
                    
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnnSvr
                    
                    While Not RecTChild.EOF
                        fltAmount = RecTChild!fltAmount
                        intAccountHeadID = RecTChild!intAccountHeadID
                        intByAccountHeadID = 1504
                        intSlNo = intSlNo + 1
                        tinDebitOrCreditFlag = 0
                        vchNarration = Null
                        intFundID = 1
                        
                        arrInput = Array(intTransactionID, intSlNo, _
                        intAccountHeadID, _
                        fltAmount, _
                        tinDebitOrCreditFlag, _
                        intByAccountHeadID, _
                        vchNarration, _
                        intFundID)
                        
                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnnSvr
                        RecTChild.MoveNext
                    Wend
                    RecTChild.Close
                    msqlStatus = "update  faVouchers set tnyStatus=1 where tnyStatus=0"
                    RecStatus.Open msqlStatus, mCnn
                    vsGridInboxCash.TextMatrix(mRowCount, 17) = 1
                    Rec.MoveNext
        Wend
        Rec.Close
    End If
'     MsgBox "Successfully Generated !", vbInformation
End If
    Call UpdtaeRecieptField
    
End Sub
Private Sub UpdtaeRecieptField()
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mCount As Integer
    Dim FromDate As Variant
    Dim ToDate As Variant
    Dim arrIn As Variant
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
     
     For mCount = 1 To vsGridInboxCash.Rows - 1
        If val(vsGridInboxCash.TextMatrix(mCount, 1)) > 0 Then
          If vsGridInboxCash.TextMatrix(mCount, 17) = "1" Then
            FromDate = dtPkrFromDate     ' Format(lblSelectedDate, "DD/MMM/YYYY")
            ToDate = dtPkrToDate
            arrIn = Array(FromDate, ToDate)
            Set Rec = objdb.ExecuteSP("spGetRecieptNo", arrIn, , , mCnn, adCmdStoredProc)
            If Not Rec.EOF Then
                vsGridInboxCash.TextMatrix(mCount, 7) = Rec!intVoucherNo
            End If
          End If
         End If
     Next
     cmdReciept.Enabled = False
End Sub
Private Sub Form_Load()
    dtPkrFromDate.value = "1/" & Month(gbTransactionDate) & "/" & Year(gbTransactionDate)
    dtPkrToDate.value = gbTransactionDate
    Call PopulateList(cmbDemandType, "select vchDemandMode from faDemandMode Order By vchDemandMode", , , True)
    Call PopulateList(cmbZonal, "Select chvZoneNameEnglish, numZoneID From GM_Zone WHERE Right(numZoneID,2)<>1 AND intLBID =" & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster)
End Sub



