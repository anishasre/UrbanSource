VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearchPaymentOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Order Search"
   ClientHeight    =   6750
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4380
      Left            =   -15
      TabIndex        =   19
      Top             =   540
      Width           =   9960
      _cx             =   17568
      _cy             =   7726
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   15792633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483639
      BackColorAlternate=   15463925
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
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSearchPaymentOrder.frx":0000
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00F3FBFB&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   -30
      TabIndex        =   0
      Top             =   4785
      Width           =   9945
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8100
         TabIndex        =   23
         Top             =   270
         Width           =   1290
      End
      Begin VB.CommandButton cmdViewPOReport 
         Caption         =   "View Payorder Report"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9675
         TabIndex        =   22
         Top             =   1575
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7110
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1305
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1305
         Width           =   1515
      End
      Begin VB.CheckBox chkListToApprove 
         BackColor       =   &H00F3FBFB&
         Caption         =   "List of Approved Pay Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   75
         TabIndex        =   12
         Top             =   1455
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5625
         TabIndex        =   11
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5610
         TabIndex        =   10
         Top             =   555
         Width           =   1185
      End
      Begin VB.CommandButton cmdGeneratedSeat 
         Caption         =   "..."
         Height          =   330
         Left            =   3720
         TabIndex        =   7
         Top             =   555
         Width           =   285
      End
      Begin VB.TextBox txtForwardedSeat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox txtGeneratedSeat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   555
         Width           =   1290
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5610
         TabIndex        =   4
         Top             =   885
         Width           =   2715
      End
      Begin VB.CommandButton cmdForwardedSeat 
         Caption         =   "..."
         Height          =   330
         Left            =   3720
         TabIndex        =   3
         Top             =   225
         Width           =   285
      End
      Begin VB.CommandButton cmdTransactionType 
         Caption         =   "..."
         Height          =   315
         Left            =   8340
         TabIndex        =   2
         Top             =   885
         Width           =   285
      End
      Begin VB.TextBox txtPayOrderNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   885
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   6810
         TabIndex        =   8
         Top             =   555
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67502081
         CurrentDate     =   40197
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   6810
         TabIndex        =   9
         Top             =   225
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67502081
         CurrentDate     =   40197
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7335
         TabIndex        =   24
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblPayOrderGeneratedSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Order Generated Seat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   18
         Top             =   570
         Width           =   2145
      End
      Begin VB.Label lblFromDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4650
         TabIndex        =   17
         Top             =   210
         Width           =   870
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4890
         TabIndex        =   16
         Top             =   555
         Width           =   645
      End
      Begin VB.Label lblForwardedSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Order Forwarded Seat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   165
         TabIndex        =   15
         Top             =   255
         Width           =   2160
      End
      Begin VB.Label lblPaymentOrderNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Order No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   810
         TabIndex        =   14
         Top             =   915
         Width           =   1515
      End
      Begin VB.Label lblTransactionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4065
         TabIndex        =   13
         Top             =   900
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmSearchPaymentOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim intLoadMode         As Integer
    Private tnyStatus       As Integer      'As Property
    Dim mPendingTask        As Integer  '1 Pending Task in Previous Year,2 Cancel PO in Previous Year

    Private Function GetSeatName(mSeatID As Variant)
        Dim mCnnSeatName    As New ADODB.Connection
        Dim RecSeatName     As New ADODB.Recordset
        Dim objSeatName     As New clsDB
        Dim mSQLSeatName    As String
        
        On Error GoTo Err:
        objSeatName.CreateNewConnection mCnnSeatName, enuSourceString.DBMaster
        
        mSQLSeatName = "Select * From GL_Seats"
        mSQLSeatName = mSQLSeatName + " Where numSeatID = " & mSeatID
        RecSeatName.Open mSQLSeatName, mCnnSeatName
        If Not (RecSeatName.EOF And RecSeatName.BOF) Then
            GetSeatName = IIf(IsNull(RecSeatName!chvSeatTitle), "", RecSeatName!chvSeatTitle)
        End If
        RecSeatName.Close
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub FormIntialize()
        Dim mCrl As Control
        
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
    End Sub
    
    Private Sub FetchPaymentOrder()
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mArrIn      As Variant
        Dim mSql        As String
        Dim mFromDate   As Variant
        Dim mToDate     As Variant
        Dim mStatus     As Variant
        Dim mCnt        As Integer
        On Error GoTo Err:
        tnyStatus = val(chkListToApprove.value) ' = 1
        vsGrid.Rows = 1
        objDb.SetConnection mCnn
        'Call FinancialYearSetForPEndingTask
        
        If txtDateFrom.Text = "" Then
            mSql = "Select dtStartingDate From  faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mFromDate = gbStartingDate   'IIf(IsNull(Rec!dtStartingDate), "", CheckDateInMMM(Rec!dtStartingDate))
            End If
            Rec.Close
        Else
            mFromDate = CheckDateInMMM(txtDateFrom.Text)
        End If
        
        
        
        If txtDateTo.Text = "" Then
            mSql = "Select dtEndingDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mToDate = gbEndingDate 'IIf(IsNull(Rec!dtEndingDate), "", CheckDateInMMM(Rec!dtEndingDate))
            End If
            Rec.Close
        Else
            mToDate = CheckDateInMMM(txtDateTo.Text)
        End If
        
'        If CDate(mFromDate) < CDate(DateAdd("M", -1, gbStartingDate)) Or mFromDate > gbEndingDate Then
'            mFromDate = gbStartingDate
'        End If
'        If mToDate < gbStartingDate Or mToDate > gbEndingDate Then
'            mToDate = gbEndingDate
'        End If
        If mPendingTask = 0 Then
            If mFromDate < gbStartingDate Then
                mFromDate = gbStartingDate
                txtDateFrom.Text = DdMmmYy(gbStartingDate)
            End If
        Else
            If Not (mFromDate >= DateAdd("yyyy", -1, gbStartingDate) And mFromDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                mFromDate = DateAdd("yyyy", -1, gbStartingDate)
                txtDateFrom.Text = DdMmmYy(CDate(mFromDate))
            End If
            If Not (mToDate >= DateAdd("yyyy", -1, gbStartingDate) And mToDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                mToDate = DateAdd("yyyy", -1, CDate(gbEndingDate))
                txtDateTo = DdMmmYy(CDate(mToDate)) 'DateAdd("yyyy", -1, gbEndingDate)
            End If
            
        End If
'        If mPendingTask = 2 Then
'            txtDateFrom.Text = Format(CDate(DateAdd("yyyy", -1, gbStartingDate)), "dd/mmm/yyyy")
'            txtDateTo.Text = Format(CDate(DateAdd("yyyy", -1, gbEndingDate)), "dd/mmm/yyyy")
'        End If


        mSql = "Select *,faPayOrder.vchDescription As Descriptions,faPayOrder.tnyStatus As Status,faPayOrder.numSeatID As SeatID  From faPayOrder"
        mSql = mSql + " Inner Join faPayOrderChild ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID "
        'If mPendingTask <> 8 Then
        mSql = mSql + " And faPayOrderChild.tnyCategoryFlag = 3"
        'End If
        mSql = mSql + " Inner Join faTransactionType ON faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID"
        mSql = mSql + " Inner Join faUser On faUser.numUserID = faPayOrder.numUserID"
        mSql = mSql + " Left Join faVouchers On faVouchers.intVoucherID = faPayOrder.intVoucherID"
        mSql = mSql + " Where (tnyCancelled <> 1 Or tnyCancelled Is Null)"
        
        If Trim(txtForwardedSeat.Text) <> "" Then
            mSql = mSql + " And numFwdSeatID Like   '%' + '" & Trim(txtForwardedSeat.Tag) & " ' + '%'"
        End If
        If Trim(txtGeneratedSeat.Text) <> "" Then
            mSql = mSql + " And faPayOrder.numSeatID   Like  '%' + '" & Trim(txtGeneratedSeat.Tag) & " '  + '%'"
        End If
        If Trim(txtPayOrderNo.Text) <> "" Then
            mSql = mSql + " And faPayOrder. vchPayOrderNo Like '" & Trim(txtPayOrderNo.Text) & "' "
        End If
        If Trim(txtTransactionType.Text) <> "" Then
            mSql = mSql + " And faPayOrder. intTransactionTypeID Like " & txtTransactionType.Tag & " "
        End If
        If Trim(txtAmount.Text) <> "" Then
            mSql = mSql + " And numAmount = " & Trim(txtAmount.Text) & " "
        End If
        If txtDateFrom.Text <> "" And txtDateTo.Text <> "" Then
             mSql = mSql + " And dtPayOrderDate Between '" & txtDateFrom.Text & "' And ' " & txtDateTo.Text & " ' "
        End If
        If mPendingTask = 1 Then
            If Trim(txtPayOrderNo.Text) <> "" Then
                'mSql = mSql + " And isNull(faPayOrder.tnyStatus,0) <> 1"
                mSql = mSql + " AND faPayOrder.intPayOrderID not in (Select numDemandID From faPendingTaskRequest Where intTaskID is Null AND intTaskId=7 And tnyStatus=8)"
            Else
                mSql = mSql + " And isNull(faPayOrder.tnyStatus,0) <> 1"
                mSql = mSql + " AND faPayOrder.intPayOrderID not in (Select numDemandID From faPendingTaskRequest Where intTaskID is Null AND intTaskId=7 And tnyStatus=8)"
            End If
        ElseIf mPendingTask = 0 Then
            mSql = mSql + " And isNull(faPayOrder.tnyStatus,0) in(0,1,2)"
        Else
            If chkListToApprove.value = 1 Then
                mSql = mSql + " And faPayOrder.tnyStatus = 1"
            End If
        End If

        mSql = mSql + " Order By vchPayOrderNo Desc "
      

        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objDb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        mCnt = 0
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                mCnt = mCnt + 1
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!vchPayOrderNo
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = DdMmmYy(Rec!dtPayOrderDate)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = Rec!numAmount
                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = GetSeatName(Rec!SeatID)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = Rec!vchUserName
                If Rec!Status = 1 Then
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 5) = vbChecked
                Else
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 5) = vbUnchecked
                End If
                vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!intPayOrderID), "", Rec!intPayOrderID)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = IIf(IsNull(Rec!intModuleID), "", Rec!intModuleID)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                
                If mPendingTask = 0 Or mPendingTask = 8 Or mPendingTask = 99 Then
                    If CDate(vsGrid.TextMatrix(vsGrid.Rows - 1, 1)) < CDate(gbStartingDate) Then
                        'If Rec!intTransactionTypeID <> gbTransactionTypePayBills Then
                            'vsGrid.RowHidden(mCnt) = True
                        'End If
                    End If
                Else
                    If IsNumeric(Rec!intVoucherNo) Then
                        vsGrid.RowHidden(mCnt) = True
                    End If
                End If
                
                Rec.MoveNext
            Wend
        End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub chkListToApprove_Click()
        If chkListToApprove.value = vbChecked Then
            tnyStatus = 1
        Else
            tnyStatus = 0
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdForwardedSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtForwardedSeat.Text = gbSearchStr
            txtForwardedSeat.Tag = gbSearchID
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdGeneratedSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtGeneratedSeat.Text = gbSearchStr
            txtGeneratedSeat.Tag = gbSearchID
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdNew_Click()
        frmPaymentOrder.Visible = True
        frmPaymentOrder.ZOrder (0)
    End Sub

    Private Sub cmdSearch_Click()
        Call FetchPaymentOrder
    End Sub

    Private Sub cmdTransactionType_Click()
        frmSearchTransactionType.Show vbModal
        
        txtTransactionType.Text = Trim(gbSearchStr)
        txtTransactionType.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub

    Private Sub cmdView_Click()
        If vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = vbChecked Then
            Dim aryIn As Variant
            aryIn = Array(vsGrid.TextMatrix(vsGrid.Row, 1))
            frmViewVoucher.ArrayIn = aryIn
            frmViewVoucher.FormName = "frmViewPaymentOrder"
            frmViewVoucher.Show vbModal
        Else
            MsgBox "Approved Payment Orders only have Journals", vbInformation
        End If
    End Sub

    Private Sub dtpDateFrom_CloseUp()
        txtDateFrom.Text = dtpDateFrom.value
    End Sub
    
    Private Sub dtpDateTo_CloseUp()
        txtDateTo.Text = dtpDateTo.value
    End Sub

    Private Sub Form_Load()
        Dim mSql As String
        Dim mDt As Date
        Call FormIntialize
        Call FinancialYearSetForPEndingTask
        mDt = DateSerial(Year(gbTransactionDate), Month(gbTransactionDate), 1)
        mDt = DateAdd("m", -1, mDt)
        If mDt < gbStartingDate Then
            mDt = gbStartingDate
        End If
        txtDateFrom.Text = DdMmmYy(mDt)
        txtDateTo.Text = DdMmmYy(gbTransactionDate)
        If mPendingTask = 2 Then
            txtDateFrom.Text = Format(CDate(DateAdd("yyyy", -1, gbStartingDate)), "dd/mmm/yyyy")
            txtDateTo.Text = Format(CDate(DateAdd("yyyy", -1, gbEndingDate)), "dd/mmm/yyyy")
        End If
        cmdSearch.Enabled = False
        chkListToApprove.value = 1
        Call FetchPaymentOrder
        cmdSearch.Enabled = True
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        intLoadMode = 0
        mPendingTask = 0
    End Sub
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDateFrom_LostFocus()
        If txtDateFrom.Text <> "" Then
            txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
            If CDate(txtDateFrom.Text) < CDate(gbStartingDate) Then
                If CDate(txtDateFrom.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                    txtDateFrom.Text = DateAdd("yyyy", -1, gbStartingDate)
                    txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
                End If
                txtDateTo.Text = DateAdd("yyyy", -1, gbEndingDate)
                txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
            End If
        End If
    End Sub

    Private Sub txtDateTo_LostFocus()
        If txtDateTo.Text <> "" Then
            txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
        End If
    End Sub

    Private Sub txtForwardedSeat_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then 'Delete Key
            txtForwardedSeat.Text = ""
            txtForwardedSeat.Tag = ""
        Else
            txtForwardedSeat.Locked = True
        End If
    End Sub
    
    Private Sub txtGeneratedSeat_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then 'Delete Key
            txtGeneratedSeat.Text = ""
            txtGeneratedSeat.Tag = ""
        Else
            txtGeneratedSeat.Locked = True
        End If
    End Sub
   Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = 46 Then 'Delete Key
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        Else
            txtTransactionType.Locked = True
        End If
    End Sub

    Private Sub vsGrid_DblClick()
        
        If vsGrid.Row > 0 Then
            If intLoadMode = 50 Then
                If vsGrid.TextMatrix(vsGrid.Row, 8) = "" Then
                    gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                    gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 0))
                    Unload Me
                End If
            Else
                gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 0))
                Unload Me
            End If
        End If
    End Sub
    Private Sub FinancialYearSetForPEndingTask()
        Dim mSql    As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim Trndate     As Date
        Dim mTrnYear    As Integer
        Dim Curyear     As Integer
        Dim mStartdate   As Date
        Dim mEnddate   As Date
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select * From faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Set Rec = objDb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                Curyear = Rec!intFinancialYear
            End If
            Rec.Close
            If mPendingTask = 1 Then
                mSql = "Select * From faFinancialYear Where intFinancialYear=" & Curyear - 1
                    Set Rec = objDb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.EOF Or Rec.BOF) Then
                        mStartdate = Rec!dtStartingDate
                        mEnddate = Rec!dtEndingDate
                        'gbStartingDate = mStartdate
                        'gbEndingDate = mEnddate
                        'gbTransactionDate = mEnddate
                        'gbFinancialYearID = Curyear - 1
                    End If
                    Rec.Close
            Else
                mSql = "Select *,GetDate() as TrnDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
                Set Rec = objDb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    mTrnYear = Rec!intFinancialYear
                    Trndate = Rec!Trndate
                    'gbTransactionDate = Trndate
                    'gbFinancialYearID = mTrnYear
                    'gbStartingDate = Rec!dtStartingDate
                    'gbEndingDate = Rec!dtEndingDate
                End If
                Rec.Close
            End If
            
            mCnn.Close
        End If
    End Sub
    Public Property Let Staus(ByVal val As Integer)
        tnyStatus = val
    End Property

    Public Property Let Mode(ByVal val As Integer)
        intLoadMode = val
    End Property
    Public Property Let PendingTask(ByVal val As Integer)
        mPendingTask = val
    End Property
    

