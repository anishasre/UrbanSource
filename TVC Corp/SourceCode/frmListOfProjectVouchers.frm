VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfProjectVouchers 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13890
   Icon            =   "frmListOfProjectVouchers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   13875
      TabIndex        =   8
      Top             =   6720
      Width           =   13935
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   825
      Width           =   1095
   End
   Begin VB.ComboBox cmbChangedRecords 
      Height          =   315
      Left            =   10110
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   12300
      Top             =   7680
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   5295
      Left            =   -15
      TabIndex        =   2
      Top             =   1440
      Width           =   13890
      _cx             =   24500
      _cy             =   9340
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfProjectVouchers.frx":1CCA
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
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmListOfProjectVouchers.frx":2038
      Left            =   1485
      List            =   "frmListOfProjectVouchers.frx":203A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   825
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13890
      TabIndex        =   0
      Top             =   0
      Width           =   13890
   End
   Begin VB.Label Label3 
      Caption         =   "Year :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   825
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "List Of Changed Records"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7350
      TabIndex        =   4
      Top             =   855
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Month :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   570
      TabIndex        =   3
      Top             =   825
      Width           =   810
   End
End
Attribute VB_Name = "frmListOfProjectVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Sub cmdAdd_Click()
        frmEditProjectVoucherDetails.Show vbModal
    End Sub
    Private Sub cmbChangedRecords_Click()
        Call FillGrid
    End Sub

    Private Sub cmbMonth_Click()
        If cmbMonth.ListIndex > -1 Then
                Call FillGrid
        End If
    End Sub


    Private Sub cmbYear_Click()
        If cmbYear.ListIndex > -1 Then
                Call FillGrid
        End If
    End Sub

    Private Sub Form_Load()
        vsGrid.Cell(flexcpFontName, 0) = "Verdana"
        XPC.InitSubClassing
'        Call fillgrid
        Call PopulateList(cmbMonth, "SELECT  vchPeriodicity,intPeriodicityID - 10 AS newid From faPeriodicity WHERE (intTypeID = 9)", , True, True, True)
        Call PopulateList(cmbChangedRecords, "SELECT     vchPayOrderNo, intRequestID From faRequestForChangeExpVoucher WHERE (tnyStatus = 1)", , True, True, True)
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then   'gbUserTypeID = 4 Then   'Accounts Officer
           Label2.Visible = True
           cmbChangedRecords.Visible = True
        Else
            Label2.Visible = False
            cmbChangedRecords.Visible = False
        End If
        cmbYear.AddItem "2009"
        cmbYear.ItemData(cmbYear.NewIndex) = 0
        cmbYear.AddItem "2010"
        cmbYear.ItemData(cmbYear.NewIndex) = 1
        cmbYear.AddItem "2011"
        cmbYear.ItemData(cmbYear.NewIndex) = 2
    End Sub
    Private Sub FillGrid()
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mRowCnt As Integer
        Dim mChk    As Integer
        Dim mTrType As Variant
        Dim mLoop   As Integer
        Dim mDate   As String
        Dim mProject As Variant
        Dim mAmt    As Variant

        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT faPayOrder.intPayOrderID,faPayOrder.vchPayOrderNo, faPayOrder.dtPayOrderDate, MONTH(faPayOrder.dtPayOrderDate) AS Month,faVouchers.intVoucherID, faVouchers.intVoucherNo, faPayOrder.dtVoucherDate, "
            mSQL = mSQL + " faTransactionType.vchTransactionType, faVouchers.fltAmount, faVouchers.vchInstrumentNo, faPayOrder.numProjectNo,"
            mSQL = mSQL + " suSourceOfFund.vchSourceFundName ,dtDate, faTransactionCategory.vchTransactionCategory,faAllotments.intID,faAllotments.vchAllotmentNo, faTransactionType.intTransactionTypeID, faPayOrder.tnyCancelled, faPayOrder.tnyStatus,"
            mSQL = mSQL + " suSourceOfFund.intSourceFundID, faTransactionCategory.intCategoryID, faRequestForChangeExpVoucher.tnyStatus As NewStatus, faFunctions.vchFunction, faFunctionaries.vchFunctionary ,faAgreements.vchAgreementNo,faPayOrder.intAgreementID  FROM faPayOrder "
            mSQL = mSQL + " LEFT OUTER Join faFunctions ON faPayOrder.intFunctionID = faFunctions.intFunctionID LEFT OUTER JOIN"
            mSQL = mSQL + " faFunctionaries ON faPayOrder.intFunctionaryID = faFunctionaries.intFunctionaryID LEFT OUTER JOIN"
            mSQL = mSQL + " faRequestForChangeExpVoucher ON faPayOrder.vchPayOrderNo = faRequestForChangeExpVoucher.vchPayOrderNo LEFT OUTER JOIN"
            mSQL = mSQL + " faTransactionType ON faPayOrder.intTransactionTypeID = faTransactionType.intTransactionTypeID LEFT OUTER JOIN"
            mSQL = mSQL + " faVouchers ON faPayOrder.intVoucherID = faVouchers.intVoucherID AND"
            mSQL = mSQL + " faPayOrder.intVoucherNo = faVouchers.intVoucherNo LEFT OUTER JOIN"
            mSQL = mSQL + " suSourceOfFund ON faPayOrder.intSourceOfFundID = suSourceOfFund.intSourceFundID LEFT OUTER JOIN"
            mSQL = mSQL + " faTransactionCategory ON faPayOrder.tnyCategoryID = faTransactionCategory.intCategoryID LEFT OUTER JOIN"
            mSQL = mSQL + " faAllotments ON faPayOrder.intAllotmentID = faAllotments.intID "
            mSQL = mSQL + " LEFT OUTER JOIN faAgreements ON faPayOrder.intAgreementID=faAgreements.intAgreementID "
            'mSQL = mSQL + " Where faPayOrder.tnyStatus = 1"
            If cmbMonth.ListIndex > 0 Then
                mSQL = mSQL + " Where  MONTH(faPayOrder.dtPayOrderDate) = " & cmbMonth.ListIndex & " "
                    If cmbYear.ListIndex > -1 Then
                        mSQL = mSQL + " and  YEAR(faPayOrder.dtPayOrderDate) = " & cmbYear.Text & " "
                    End If
            End If
            
            If cmbChangedRecords.ListIndex > 0 Then
                If cmbMonth.ListIndex > 0 Then
                    mSQL = mSQL + " and faPayOrder.vchPayOrderNo=   " & cmbChangedRecords.Text & "   "       'faPayOrder.tnyStatus=2"
                Else
                    mSQL = mSQL + " where faPayOrder.vchPayOrderNo=   " & cmbChangedRecords.Text & "   " 'faPayOrder.tnyStatus=2
                End If
            End If
            mSQL = mSQL + " and faPayOrder.tnyCancelled <> 1"
            mSQL = mSQL + " order by faPayOrder.intPayOrderID"
            Rec.CursorLocation = adUseClient
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            mRowCnt = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCnt, 0) = mRowCnt
                vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
                vsGrid.TextMatrix(mRowCnt, 2) = DdMmmYy(IIf(IsNull(Rec!dtPayOrderDate), "", Rec!dtPayOrderDate))
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                If Rec!dtDate <> "" Then
                    mDate = DdMmmYy(Rec!dtDate)
                Else
                    mDate = ""
                End If
                vsGrid.TextMatrix(mRowCnt, 4) = mDate 'IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                If Rec!fltAmount = "" Or Rec!fltAmount = 0 Or IsNull(Rec!fltAmount) Then
                    mAmt = ""
                Else
                    mAmt = Rec!fltAmount
                End If
                vsGrid.TextMatrix(mRowCnt, 6) = mAmt 'IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                If Rec!numProjectNo = "" Or Rec!numProjectNo = 0 Then
                    mProject = ""
                Else
                    mProject = Rec!numProjectNo
                End If
                vsGrid.TextMatrix(mRowCnt, 9) = mProject    'IIf(IsNull(Rec!numProjectNo), "", Rec!numProjectNo)
                vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                mTrType = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                If mTrType = 1141 Or mTrType = 1151 Or mTrType = 1161 Or mTrType = 1171 Or mTrType = 1181 Or mTrType = 1191 Or mTrType = 1371 Or mTrType = 1381 Then
                    vsGrid.TextMatrix(mRowCnt, 11) = "Sulekha-Project"
                ElseIf mTrType = 1201 Or mTrType = 1391 Then
                    vsGrid.TextMatrix(mRowCnt, 11) = "Non-Project"
                Else
                    vsGrid.TextMatrix(mRowCnt, 11) = "Non-Plan"
                End If
                If Rec!tnyCancelled = 1 Then
                    vsGrid.TextMatrix(mRowCnt, 12) = "Cancelled"
                End If
                
                vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!vchAgreementNo), "", Rec!vchAgreementNo)
                
                vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                vsGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!intPayOrderID), "", Rec!intPayOrderID)
                If Rec!NewStatus = 1 Or Rec!NewStatus = 2 Then 'Rec!tnyStatus = 2 Or Rec!tnyStatus = 3 Then
                    vsGrid.TextMatrix(mRowCnt, 16) = vbChecked
                Else
                     vsGrid.TextMatrix(mRowCnt, 16) = Unchecked
                End If
                If Rec!NewStatus = 2 Then
                    For mLoop = 0 To vsGrid.Cols - 1
                        vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HD2AE9E
                        vsGrid.Editable = flexEDNone
                    Next
                End If
                vsGrid.TextMatrix(mRowCnt, 17) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                vsGrid.TextMatrix(mRowCnt, 18) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                vsGrid.TextMatrix(mRowCnt, 19) = IIf(IsNull(Rec!intCategoryID), "", Rec!intCategoryID)
                vsGrid.TextMatrix(mRowCnt, 20) = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
                vsGrid.TextMatrix(mRowCnt, 21) = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                vsGrid.TextMatrix(mRowCnt, 22) = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                vsGrid.TextMatrix(mRowCnt, 23) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
                vsGrid.TextMatrix(mRowCnt, 24) = IIf(IsNull(Rec!intAgreementID), "", Rec!intAgreementID)
                vsGrid.TextMatrix(mRowCnt, 25) = IIf(IsNull(Rec!intID), "", Rec!intID)
                Rec.MoveNext
                mRowCnt = mRowCnt + 1
            Wend
            Rec.Close
        End If
    End Sub
    Private Sub VSGrid_DblClick()
          Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSQL As String
        
       
        If vsGrid.Row > 0 Then
             If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then Exit Sub
            If vsGrid.TextMatrix(vsGrid.Row, 12) = "Cancelled" Then
                MsgBox "Payment Order Cancelled!!!Can't Edit", vbInformation, "Saankhya"
                vsGrid.Editable = flexEDNone
                Exit Sub
            End If
            frmProjectVoucherDetails.Frame1.Enabled = False
            frmProjectVoucherDetails.txtPaymentOrder.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            frmProjectVoucherDetails.txtPaymentOrderDate.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            frmProjectVoucherDetails.txtVoucherNo.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            frmProjectVoucherDetails.txtVoucherNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 17)
            
            frmProjectVoucherDetails.txtVoucherDate.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
            frmProjectVoucherDetails.txtTransactionType.Text = vsGrid.TextMatrix(vsGrid.Row, 5)
            frmProjectVoucherDetails.txtTransactionType.Tag = vsGrid.TextMatrix(vsGrid.Row, 18)
            frmProjectVoucherDetails.txtAllotmentNo.Text = vsGrid.TextMatrix(vsGrid.Row, 8)
            frmProjectVoucherDetails.txtAllotmentNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 25)
            frmProjectVoucherDetails.txtProjectNo.Text = vsGrid.TextMatrix(vsGrid.Row, 9)
            frmProjectVoucherDetails.txtCategory.Text = vsGrid.TextMatrix(vsGrid.Row, 14)
            frmProjectVoucherDetails.txtPaymentOrder.Tag = vsGrid.TextMatrix(vsGrid.Row, 15)
            If vsGrid.TextMatrix(vsGrid.Row, 11) = "Sulekha-Project" Then
                frmProjectVoucherDetails.chkProject.value = 1
            ElseIf vsGrid.TextMatrix(vsGrid.Row, 11) = "Non-Project" Then
                frmProjectVoucherDetails.chkNonProject.value = 1
            Else
                frmProjectVoucherDetails.chkNonPlan.value = 1
            End If
            frmProjectVoucherDetails.txtCategory.Tag = vsGrid.TextMatrix(vsGrid.Row, 19)
            frmProjectVoucherDetails.txtSource.Text = vsGrid.TextMatrix(vsGrid.Row, 10)
            frmProjectVoucherDetails.txtSource.Tag = vsGrid.TextMatrix(vsGrid.Row, 19)
            frmProjectVoucherDetails.txtAmount.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
            frmProjectVoucherDetails.txtChequeNo.Text = vsGrid.TextMatrix(vsGrid.Row, 7)
            frmProjectVoucherDetails.txtFunction.Text = vsGrid.TextMatrix(vsGrid.Row, 21)
            frmProjectVoucherDetails.txtFunctionary.Text = vsGrid.TextMatrix(vsGrid.Row, 22)
            frmProjectVoucherDetails.txtAgreement.Text = vsGrid.TextMatrix(vsGrid.Row, 13)
            frmProjectVoucherDetails.txtAgreement.Tag = vsGrid.TextMatrix(vsGrid.Row, 24)
        
            If vsGrid.TextMatrix(vsGrid.Row, 1) > 1 Then
                objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                mSQL = "SELECT faPayOrderChild.tnyCategoryFlag, faPayOrderChild.intAccountHeadID, faAccountHeads.vchAccountHead, faVouchers.vchBank, "
                mSQL = mSQL + " faPayOrderChild.intPayOrderID FROM faAccountHeads INNER JOIN"
                mSQL = mSQL + " faPayOrderChild ON faAccountHeads.intAccountHeadID = faPayOrderChild.intAccountHeadID INNER JOIN"
                mSQL = mSQL + " faPayOrder ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID INNER JOIN"
                mSQL = mSQL + " faVouchers ON faPayOrder.intVoucherID = faVouchers.intVoucherID "
                mSQL = mSQL + " Where faPayOrderChild.intPayOrderID = " & vsGrid.TextMatrix(vsGrid.Row, 15) & " "
                Rec.Open mSQL, mCnn
                While Not (Rec.EOF Or Rec.BOF)
                    If Rec!tnyCategoryFlag = 1 Then
                        frmProjectVoucherDetails.txtGrossExpenditureHead.Text = Rec!vchAccountHead
                        frmProjectVoucherDetails.txtGrossExpenditureHead.Tag = Rec!intAccountHeadID
                    ElseIf Rec!tnyCategoryFlag = 2 Then
                        frmProjectVoucherDetails.txtNetPayableHead.Text = Rec!vchAccountHead
                        frmProjectVoucherDetails.txtNetPayableHead.Tag = Rec!intAccountHeadID
                    End If
                    frmProjectVoucherDetails.txtPaymentCreditedFrom.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                    Rec.MoveNext
                Wend
                Rec.Close
            End If
        '-----------------------------------------------------------------------------------------------
        '---Verified----
        
            If vsGrid.TextMatrix(vsGrid.Row, 1) > 0 Then
                If vsGrid.TextMatrix(vsGrid.Row, 16) = vbChecked Then '----if Approved
                    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                    mSQL = " SELECT faRequestForChangeExpVoucher.intRequestID,faRequestForChangeExpVoucher.vchNewAllotmentLetterNo, faRequestForChangeExpVoucher.numNewProjectID, "
                    mSQL = mSQL + " faTransactionType.vchTransactionType, faTransactionCategory.vchTransactionCategory, suSourceOfFund.vchSourceFundName,"
                    mSQL = mSQL + " faRequestForChangeExpVoucher.vchReason , faVouchers.intVoucherID,faVouchers.intVoucherNo, faRequestForChangeExpVoucher.tnyStatus, faFunctions.vchFunction, faFunctionaries.vchFunctionary,faAccountHeads.intAccountHeadID, faAccountHeads.vchAccountHead, faTransactionType.intTransactionTypeID, "
                    mSQL = mSQL + " faRequestForChangeExpVoucher.intPayOrderID,faRequestForChangeExpVoucher.intNewAgreementID,faAgreements.vchAgreementNo, "
                    mSQL = mSQL + " suSourceOfFund.intSourceFundID , faRequestForChangeExpVoucher.intNewAllotmentID, faRequestForChangeExpVoucher.intCategoryID, faFunctions.intFunctionID, faFunctionaries.intFunctionaryID"
                    mSQL = mSQL + " FROM faRequestForChangeExpVoucher LEFT OUTER JOIN"
                    mSQL = mSQL + " faFunctions ON faRequestForChangeExpVoucher.intFunctionID = faFunctions.intFunctionID LEFT OUTER JOIN"
                    mSQL = mSQL + " faFunctionaries ON faRequestForChangeExpVoucher.intFunctionaryID = faFunctionaries.intFunctionaryID"
                    mSQL = mSQL + " LEFT OUTER JOIN faAccountHeads ON faRequestForChangeExpVoucher.intAccountHeadID = faAccountHeads.intAccountHeadID"
                    mSQL = mSQL + " LEFT OUTER JOIN faTransactionType ON faRequestForChangeExpVoucher.intNewTransactionTypeID = faTransactionType.intTransactionTypeID LEFT OUTER JOIN"
                    mSQL = mSQL + " faTransactionCategory ON faRequestForChangeExpVoucher.intNewCategoryID = faTransactionCategory.intCategoryID LEFT OUTER JOIN"
                    mSQL = mSQL + " suSourceOfFund ON faRequestForChangeExpVoucher.intNewSourceID = suSourceOfFund.intSourceFundID LEFT OUTER JOIN"
                    mSQL = mSQL + " faVouchers ON faRequestForChangeExpVoucher.intVoucherID = faVouchers.intVoucherID"
                    mSQL = mSQL + " LEFT OUTER JOIN faAgreements ON faRequestForChangeExpVoucher.intNewAgreementID=faAgreements.intAgreementID"
                    mSQL = mSQL + " where  faRequestForChangeExpVoucher.vchPayOrderNo= " & vsGrid.TextMatrix(vsGrid.Row, 1) & ""
                    Rec.Open mSQL, mCnn
                    While Not (Rec.EOF Or Rec.BOF)
                      frmProjectVoucherDetails.txtPaymentVoucher.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                      frmProjectVoucherDetails.txtNewAllotmentNumber.Text = IIf(IsNull(Rec!vchNewAllotmentLetterNo), "", Rec!vchNewAllotmentLetterNo)
                      frmProjectVoucherDetails.txtNewTranType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                      frmProjectVoucherDetails.txtNewTranType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                      frmProjectVoucherDetails.txtNewProjectNo.Text = IIf(IsNull(Rec!numNewProjectID), "", Rec!numNewProjectID)
                      frmProjectVoucherDetails.txtNewCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                      frmProjectVoucherDetails.txtNewSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                      frmProjectVoucherDetails.txtReason.Tag = IIf(IsNull(Rec!intRequestID), "", Rec!intRequestID)
                      frmProjectVoucherDetails.txtReason.Text = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
                      frmProjectVoucherDetails.txtNewFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                      frmProjectVoucherDetails.txtNewFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                      frmProjectVoucherDetails.txtNewAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                      frmProjectVoucherDetails.txtNewAccountHead.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                      frmProjectVoucherDetails.txtPaymentVoucher.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                      frmProjectVoucherDetails.txtNewAgreement.Tag = IIf(IsNull(Rec!intNewAgreementID), "", Rec!intNewAgreementID)
                      frmProjectVoucherDetails.txtNewAgreement.Text = IIf(IsNull(Rec!vchAgreementNo), "", Rec!vchAgreementNo)
                      frmProjectVoucherDetails.txtNewSource.Tag = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
                      frmProjectVoucherDetails.txtNewAllotmentNumber.Tag = IIf(IsNull(Rec!intNewAllotmentID), "", Rec!intNewAllotmentID)
                      frmProjectVoucherDetails.txtNewCategory.Tag = IIf(IsNull(Rec!intCategoryID), "", Rec!intCategoryID)
                      frmProjectVoucherDetails.txtNewFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                      frmProjectVoucherDetails.txtNewFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                       
                      frmProjectVoucherDetails.cmdApprove.Enabled = True
                 
                      If Rec!tnyStatus = 2 Then  'tnyStatus from faRequestForChangeExpVoucher
                          MsgBox " Already Approved", vbInformation, "Saankhya"
                          frmProjectVoucherDetails.cmdUpdate.Enabled = False
                          frmProjectVoucherDetails.cmdApprove.Enabled = False
                        
                      End If
                     

                      If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                          frmProjectVoucherDetails.cmdApprove.Enabled = False
                      End If
                      frmProjectVoucherDetails.cmdVerify.Enabled = False
                      frmProjectVoucherDetails.cmdEdit.Enabled = False
                      frmProjectVoucherDetails.cmdUpdate.Enabled = False
                      Rec.MoveNext
                    Wend
                    Rec.Close
'''                    If vsGrid.TextMatrix(vsGrid.row, 23) = 0 Then
'''                        MsgBox " Verified", vbInformation, "Saankhya"
'''                        frmProjectVoucherDetails.cmdUpdate.Enabled = False
'''                        frmProjectVoucherDetails.cmdApprove.Enabled = False
'''
'''                    End If
                    frmProjectVoucherDetails.cmdVerify.Enabled = False
                    frmProjectVoucherDetails.cmdEdit.Enabled = False
                    frmProjectVoucherDetails.Frame2.Enabled = True        'Changed
                    frmProjectVoucherDetails.cmdUpdate.Enabled = True     'Changed
                    
                Else
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        MsgBox " Pending for Verification", vbInformation, "Saankhya"
                        Exit Sub
                    End If
                End If
             End If
             'End If
             
         frmProjectVoucherDetails.Show vbModal
         Call FillGrid
         End If
    End Sub
    
