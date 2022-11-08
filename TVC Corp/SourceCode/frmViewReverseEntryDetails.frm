VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmViewReverseEntryDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Reverse Entry Details"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15720
   Icon            =   "frmViewReverseEntryDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   7897
      TabIndex        =   8
      Top             =   7695
      Width           =   1050
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   330
      Left            =   6637
      TabIndex        =   7
      Top             =   7695
      Width           =   1050
   End
   Begin VB.ComboBox cmbVoucherType 
      Height          =   315
      Left            =   10395
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   495
      Width           =   2925
   End
   Begin VB.TextBox txtDateFrom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   495
      Width           =   1185
   End
   Begin VB.TextBox txtDateTo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3300
      TabIndex        =   2
      Top             =   495
      Width           =   1185
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6570
      Left            =   0
      TabIndex        =   6
      Top             =   945
      Width           =   15660
      _cx             =   27622
      _cy             =   11589
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   15
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmViewReverseEntryDetails.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin MSComCtl2.DTPicker dtpDateTo 
      Height          =   315
      Left            =   4500
      TabIndex        =   3
      Top             =   495
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60620801
      CurrentDate     =   40834
   End
   Begin MSComCtl2.DTPicker dtpDateFrom 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   495
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60620801
      CurrentDate     =   40834
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9090
      TabIndex        =   11
      Top             =   495
      Width           =   1170
   End
   Begin VB.Label lblFromDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   495
      Width           =   840
   End
   Begin VB.Label lblToDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2580
      TabIndex        =   9
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reverse Entry Voucher Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   420
      Left            =   135
      TabIndex        =   5
      Top             =   45
      Width           =   15540
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "frmViewReverseEntryDetails.frx":1FD0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15660
   End
End
Attribute VB_Name = "frmViewReverseEntryDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Sub FillVoucherType()
        cmbVoucherType.AddItem ""
        cmbVoucherType.AddItem "Receipt"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 10
        cmbVoucherType.AddItem "Contra"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 30
        cmbVoucherType.AddItem "Journal"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 40
        cmbVoucherType.AddItem "PayOrder"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 50
    End Sub
    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub
    Private Sub dtpDateFrom_CloseUp()
        txtDateFrom.Text = dtpDateFrom.value
    End Sub
    Private Sub dtpDateTo_CloseUp()
        txtDateTo.Text = dtpDateTo.value
    End Sub
    Private Sub Form_Load()
        txtDateFrom.Text = CheckDateInMMM(DateAdd("m", -1, gbTransactionDate))
        txtDateTo.Text = CheckDateInMMM(DateAdd("m", 0, gbTransactionDate))
        Call FillVoucherType
        Call FillGrid
    End Sub
    Private Sub FillGrid()
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim RecRev      As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim arrIn       As Variant
        Dim mRow        As Integer
        Dim VrType      As Integer
        Dim mVrID       As Double
        Dim mStatus     As Integer
        vsGrid.HighLight = flexHighlightNever
        vsGrid.MergeRow(0) = True
        vsGrid.MergeCol(0) = True
        vsGrid.MergeCol(1) = True
        vsGrid.MergeCol(3) = True
        vsGrid.MergeCol(9) = True
        vsGrid.MergeCol(12) = True
        vsGrid.MergeCol(13) = True
        vsGrid.Cell(flexcpFontBold, 0, 0, , 13) = True
        vsGrid.MergeCells = flexMergeFree
        vsGrid.MergeCol(2) = True
        vsGrid.WordWrap = True
        If txtDateFrom.Text = "" Then
            MsgBox "Please Enter Request Date From", vbInformation
            Exit Sub
        End If
         If txtDateTo.Text = "" Then
            MsgBox "Please Enter Request To Date", vbInformation
            Exit Sub
        End If
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "SELECT faReverseEntry.tnyVoucherTypeID,faReverseEntry.dtRequestDate ReqDate,faReasons.vchreason,faReverseEntry.numDemandNo,"
            mSql = mSql + vbNewLine + "vchAccountHead,faReverseEntry.tnyStatus,"
            mSql = mSql + vbNewLine + "faReverseEntry.intPaymentVoucherID,dbo.fnGetUser(faReverseEntry.numRequestedUserID) [ReqUsers],"
            mSql = mSql + vbNewLine + "faVouchers.intVoucherID,faVouchers.intVoucherNo, faVouchers.dtDate,faVouchers.fltAmount"
            mSql = mSql + vbNewLine + "From faReverseEntry"
            mSql = mSql + vbNewLine + "INNER  JOIN faReasons ON faReasons.intReasonID=faReverseEntry.intReasonID"
            mSql = mSql + vbNewLine + "INNER  JOIN faReverseEntryChild ON faReverseEntryChild.intRequestID=faReverseEntry.intRequestID"
            mSql = mSql + vbNewLine + "INNER JOIN faVouchers ON faVouchers.intVoucherID= faReverseEntryChild.intVoucherID"
            mSql = mSql + vbNewLine + "INNER JOIN faAccountHeads On faAccountHeads.intAccountHeadID=faVouchers.intKeyID1"
            If txtDateFrom.Text <> "" And txtDateTo.Text <> "" Then
                mSql = mSql + vbNewLine + "WHERE faReverseEntry.tnyVoucherTypeID<>50 And dtRequestDate between '" & txtDateFrom.Text & "' And '" & txtDateTo.Text & "'"
            End If
            If cmbVoucherType.ListIndex > 0 Then
                mSql = mSql + vbNewLine + "And faReverseEntry.tnyVoucherTypeID=" & cmbVoucherType.ItemData(cmbVoucherType.ListIndex)
            End If
            'mSQL = mSQL + vbNewLine + "Order By faReverseEntry.tnyVoucherTypeID,faReverseEntry.dtRequestDate "
            '-----
            mSql = mSql + vbNewLine + "Union All"
            mSql = mSql + vbNewLine + "SELECT faReverseEntry.tnyVoucherTypeID,faReverseEntry.dtRequestDate ReqDate,faReasons.vchreason,faReverseEntry.numDemandNo,"
            mSql = mSql + vbNewLine + "vchAccountHead,faReverseEntry.tnyStatus,"
            mSql = mSql + vbNewLine + "faReverseEntry.intPaymentVoucherID,dbo.fnGetUser(faReverseEntry.numRequestedUserID) [ReqUsers],"
            mSql = mSql + vbNewLine + "faVouchers.intVoucherID , faVouchers.intVoucherNo, faVouchers.dtDate,faVouchers.fltAmount"
            mSql = mSql + vbNewLine + "From faReverseEntry"
            mSql = mSql + vbNewLine + "INNER JOIN faReasons ON faReasons.intReasonID=faReverseEntry.intReasonID"
            mSql = mSql + vbNewLine + "Left JOIN faVouchers ON faVouchers.intVoucherNo= faReverseEntry.intPaymentVoucherID"
            mSql = mSql + vbNewLine + "Left JOIN faAccountHeads On faAccountHeads.intAccountHeadID=faVouchers.intKeyID1"
            If txtDateFrom.Text <> "" And txtDateTo.Text <> "" Then
                mSql = mSql + vbNewLine + "WHERE faReverseEntry.tnyVoucherTypeID=50 And dtRequestDate between '" & txtDateFrom.Text & "' And '" & txtDateTo.Text & "'"
            End If
            If cmbVoucherType.ListIndex > 0 Then
                mSql = mSql + vbNewLine + "And faReverseEntry.tnyVoucherTypeID=" & cmbVoucherType.ItemData(cmbVoucherType.ListIndex)
            End If
            mSql = mSql + vbNewLine + "Order By faReverseEntry.tnyVoucherTypeID,faReverseEntry.dtRequestDate Desc "
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            vsGrid.Clear 1
            vsGrid.Rows = 2
            mRow = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF)
                   mVrID = IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)
                   mStatus = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
                   mRow = mRow + 1
                   vsGrid.Rows = vsGrid.Row + mRow
                   VrType = IIf(IsNull(Rec!tnyVoucherTypeID), 0, Rec!tnyVoucherTypeID)
                    If VrType = 10 Then
                        vsGrid.TextMatrix(mRow, 0) = "Receipt"
                    ElseIf VrType = 30 Then
                        vsGrid.TextMatrix(mRow, 0) = "Contra"
                    ElseIf VrType = 40 Then
                        vsGrid.TextMatrix(mRow, 0) = "Journal"
                    ElseIf VrType = 50 Then
                        vsGrid.TextMatrix(mRow, 0) = "PayOrder"
                    End If
                    vsGrid.TextMatrix(mRow, 1) = Format(IIf(IsNull(Rec!ReqDate), 0, Rec!ReqDate), "DD-MMM-YYYY")
                    vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!vchReason), 0, Rec!vchReason)
                    If VrType = 50 Then
                        vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!numDemandNo), "", Rec!numDemandNo)
                        vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!intPaymentVoucherID), "", Rec!intPaymentVoucherID)
                        vsGrid.TextMatrix(mRow, 6) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    Else
                        vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                        vsGrid.TextMatrix(mRow, 5) = Format(IIf(IsNull(Rec!dtDate), "", Rec!dtDate), "DD-MMM-YYYY")
                        vsGrid.TextMatrix(mRow, 6) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                        vsGrid.TextMatrix(mRow, 7) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead) 'accHead
                    End If
                    vsGrid.TextMatrix(mRow, 7) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead) 'accHead
                    vsGrid.TextMatrix(mRow, 12) = IIf(IsNull(Rec!ReqUsers), "", Rec!ReqUsers) 'ReqUser
                    If mStatus = 0 Then
                        vsGrid.TextMatrix(mRow, 13) = "Request"
                    ElseIf mStatus = 1 Then
                        vsGrid.TextMatrix(mRow, 13) = "First Level Approved"
                    ElseIf mStatus = 2 Then
                        vsGrid.TextMatrix(mRow, 13) = "Final"
                    ElseIf mStatus = 4 Then
                        vsGrid.TextMatrix(mRow, 13) = "Cancelled"
                    End If
                    vsGrid.TextMatrix(mRow, 14) = mVrID
                    If mVrID > 0 Then
                        mSql = "Select faVouchers.*,vchAccountHead From faVouchers Inner Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
                        mSql = mSql + " Where numLinkKeyID = " & mVrID
                        Set RecRev = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                        If Not (RecRev.EOF And RecRev.BOF) Then
                            vsGrid.TextMatrix(mRow, 8) = IIf(IsNull(RecRev!intVoucherNo), "", RecRev!intVoucherNo)
                            vsGrid.TextMatrix(mRow, 9) = Format(IIf(IsNull(RecRev!dtDate), "", RecRev!dtDate), "DD-MMM-YYYY")
                            vsGrid.TextMatrix(mRow, 10) = IIf(IsNull(RecRev!fltAmount), "", RecRev!fltAmount)
                            vsGrid.TextMatrix(mRow, 11) = IIf(IsNull(RecRev!vchAccountHead), "", RecRev!vchAccountHead)
                        End If
                    End If
                    Rec.MoveNext
                Wend
            Else
                MsgBox "No Record Exists", vbInformation
                Exit Sub
            End If
            
        End If
    End Sub
    
    Private Sub txtDateFrom_LostFocus()
        If IsDate(txtDateFrom.Text) Then
            txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
            If CDate(txtDateFrom.Text) > CDate(txtDateTo.Text) Then
                MsgBox "From Date Must be Less than To date", vbApplicationModal
                Exit Sub
            End If
            Call FillGrid
        Else
            MsgBox "Please Enter Valid Date", vbInformation
            Exit Sub
        End If
    End Sub
    Private Sub txtDateTo_LostFocus()
        If IsDate(txtDateTo.Text) Then
            txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
            If CDate(txtDateTo.Text) < CDate(txtDateFrom.Text) Then
                MsgBox "To Date Must be Greater than From date", vbApplicationModal
                Exit Sub
            End If
            Call FillGrid
        Else
            MsgBox "Please Enter Valid Date", vbInformation
            Exit Sub
        End If
    End Sub
    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 1 Then
            If vsGrid.Col = 4 Then
                If vsGrid.TextMatrix(vsGrid.Row, 4) <> " " Then
                    If (Left(vsGrid.TextMatrix(vsGrid.Row, 4), 1) = 1) Then
                        frmReceipt.DisplayReceiptDetails (vsGrid.TextMatrix(vsGrid.Row, 4))
                    ElseIf (Left(vsGrid.TextMatrix(vsGrid.Row, 4), 1) = 2) Then
                        Call frmIntegratedPayments.DisplayVoucherDetails(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                        frmIntegratedPayments.cmdNew.Enabled = False
                        frmIntegratedPayments.cmdSave.Enabled = False
                    ElseIf (Left(vsGrid.TextMatrix(vsGrid.Row, 4), 1) = 3) Then
                        frmContraEntry.ListContraDemandOrVoucher (val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                        frmContraEntry.cmdNew.Enabled = False
                        frmContraEntry.txtVoucherNo.Enabled = False
                        frmContraEntry.cmdSave.Enabled = False
                     ElseIf (Left(vsGrid.TextMatrix(vsGrid.Row, 4), 1) = 4) Then
'                        frmJournalEntry.txtVoucherNo.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
'                        frmJournalEntry.txtVoucherNo_LostFocus
                        Call DisplayJournal(vsGrid.TextMatrix(vsGrid.Row, 4))
                        frmJournalEntry.cmdNew.Enabled = False
                        frmJournalEntry.cmdSave.Enabled = False
                    End If
                End If
            ElseIf vsGrid.Col = 8 Then
                If vsGrid.TextMatrix(vsGrid.Row, 8) <> " " Then
                    If (Left(vsGrid.TextMatrix(vsGrid.Row, 8), 1) = 1) Then
                        frmReceipt.DisplayReceiptDetails (vsGrid.TextMatrix(vsGrid.Row, 8))
                    ElseIf (Left(vsGrid.TextMatrix(vsGrid.Row, 8), 1) = 2) Then
                        Call frmIntegratedPayments.DisplayVoucherDetails(val(vsGrid.TextMatrix(vsGrid.Row, 8)))
                        frmIntegratedPayments.cmdNew.Enabled = False
                        frmIntegratedPayments.cmdSave.Enabled = False
                    ElseIf (Left(vsGrid.TextMatrix(vsGrid.Row, 8), 1) = 3) Then
                        frmContraEntry.ListContraDemandOrVoucher (val(vsGrid.TextMatrix(vsGrid.Row, 8)))
                        frmContraEntry.cmdNew.Enabled = False
                        frmContraEntry.txtVoucherNo.Enabled = False
                        frmContraEntry.cmdSave.Enabled = False
                    ElseIf (Left(vsGrid.TextMatrix(vsGrid.Row, 8), 1) = 4) Then
                        Call DisplayJournal(vsGrid.TextMatrix(vsGrid.Row, 8))
'                        frmJournalEntry.txtVoucherNo.Text = vsGrid.TextMatrix(vsGrid.Row, 8)
                        frmJournalEntry.cmdNew.Enabled = False
                        frmJournalEntry.cmdSave.Enabled = False
                    End If
                End If
            ElseIf vsGrid.Col = 3 Then
                If vsGrid.TextMatrix(vsGrid.Row, 3) <> " " Then
                    Call DisplayPayOrder(vsGrid.TextMatrix(vsGrid.Row, 3))
                End If
            End If
        End If
    End Sub
    
    
    Private Sub DisplayJournal(ByVal mVoucherNo)
        Dim mSql            As String
        Dim mSqlAccHeads    As String
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mCnn            As New ADODB.Connection
        Dim mRowCount       As Integer
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select *,faTransactionChild.tinDebitOrCreditFlag  From faVouchers"
        mSql = mSql + " Left Join faTransactions On faTransactions.intVoucherId = faVouchers.intVoucherId"
        mSql = mSql + " Left Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID "
        mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
        mSql = mSql + " Left Join faFunctions On fatransactions.intFunctionId = faFunctions.intFunctionId"
        mSql = mSql + " Left Join faFunctionaries On faTransactions.intFunctionaryId = faFunctionaries.intFunctionaryId"
        mSql = mSql + " Left Join faFunds On faFunds.intFundId = faTransactions.intFundId"
        mSql = mSql + " Left Join faFields On faTransactions.intFieldID = faFields.intFieldID"
        mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID"
        mSql = mSql + " Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
        mSql = mSql + " Left Join faAccountHeads On faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID"
        mSql = mSql + " Left Join faBanks On faVouchers.intKeyID1 = faBanks.intAccountHeadID"
        mSql = mSql + " Where faVouchers.intVoucherNo = " & val(mVoucherNo)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            frmJournalEntry.txtVoucherNo.Tag = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0)) 'intVocherID
            frmJournalEntry.txtVoucherNo.Text = mVoucherNo 'IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            frmJournalEntry.txtReference.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            frmJournalEntry.txtReference.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
            frmJournalEntry.txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            frmJournalEntry.txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
            frmJournalEntry.txtFund.Tag = IIf(IsNull(Rec.Fields(34)), "", Rec.Fields(34)) 'intFundID
            frmJournalEntry.txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            frmJournalEntry.txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
            frmJournalEntry.txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            frmJournalEntry.txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            frmJournalEntry.txtField.Text = IIf(IsNull(Rec!vchField), "", Rec!vchField)
            frmJournalEntry.txtField.Tag = IIf(IsNull(Rec!intFieldID), "", Rec!intFieldID)
            If Not IsNull(Rec!tinDebitOrCreditFlag) Then
                If (Rec!tinDebitOrCreditFlag) = 0 Then
                    frmJournalEntry.optDebit.value = False
                    frmJournalEntry.optCredit.value = True
                Else
                    frmJournalEntry.optDebit.value = True
                    frmJournalEntry.optCredit.value = False
                End If
            End If
            frmJournalEntry.txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
            frmJournalEntry.txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            frmJournalEntry.txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
            frmJournalEntry.txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            mSqlAccHeads = "Select * From faTransactionChild"
            mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
            mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & val(frmJournalEntry.txtReference.Tag)
            mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
            RecAccHeads.Open mSqlAccHeads, mCnn
            mRowCount = 1
            While Not Rec.EOF
                While Not RecAccHeads.EOF
                    frmJournalEntry.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                    frmJournalEntry.vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                    frmJournalEntry.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                    frmJournalEntry.vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                    frmJournalEntry.vsGrid.Rows = frmJournalEntry.vsGrid.Rows + 1
                    mRowCount = mRowCount + 1
                    RecAccHeads.MoveNext
                Wend
                Rec.MoveNext
            Wend
            RecAccHeads.Close
        End If
        Rec.Close
        frmJournalEntry.cmdNew.Enabled = False
        frmJournalEntry.cmdSave.Enabled = False
    End Sub
    
    Private Sub DisplayPayOrder(intPayOrderNo As Variant)
            Dim mCnn        As New ADODB.Connection
            Dim Rec         As New ADODB.Recordset
            Dim mSql        As String
            Dim mPayID      As Double
            Dim objdb       As New clsDB
            mSql = "Select * From faPayOrder Where vchPayOrderNo=" & val(vsGrid.TextMatrix(vsGrid.Row, 3))
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mPayID = IIf(IsNull(Rec!intPayOrderID), 0, Rec!intPayOrderID)
                frmPaymentOrder.FillPayOrder (val(mPayID))
                frmPaymentOrder.cmdApproval.Enabled = False
                frmPaymentOrder.cmdNew.Enabled = False
                frmPaymentOrder.cmdSave.Enabled = False
            End If
    End Sub

