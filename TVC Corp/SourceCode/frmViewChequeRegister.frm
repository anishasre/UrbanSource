VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmViewChequeRegister 
   Appearance      =   0  'Flat
   Caption         =   "Cheque Register"
   ClientHeight    =   9360
   ClientLeft      =   -1950
   ClientTop       =   555
   ClientWidth     =   18285
   Icon            =   "frmViewChequeRegister.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   18285
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Search Options"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8520
      TabIndex        =   13
      Top             =   240
      Width           =   9735
      Begin VB.ComboBox cmbStatus 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ListBox lstBanks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   1395
         ItemData        =   "frmViewChequeRegister.frx":1CCA
         Left            =   4980
         List            =   "frmViewChequeRegister.frx":1CCC
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdSearchBank 
         Caption         =   "..."
         Height          =   285
         Left            =   6360
         TabIndex        =   26
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdAdvancedSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtReceivedBanks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton cmdSearchTransaction 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtAmountTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   18
         Top             =   1080
         Width           =   1650
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   1080
         Width           =   1650
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
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
         Left            =   1680
         TabIndex        =   29
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Bank:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1920
         TabIndex        =   24
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4320
         TabIndex        =   21
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   1005
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   15360
      Top             =   9120
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   6780
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   18120
      _cx             =   31962
      _cy             =   11959
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmViewChequeRegister.frx":1CCE
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   8295
      Begin VB.OptionButton OptReceipts 
         Caption         =   "Receipts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptPayments 
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAccountCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1170
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   500
         TabIndex        =   4
         Top             =   1080
         Width           =   4755
      End
      Begin VB.CommandButton cmdAccoundHeads 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6E0E0&
         Caption         =   "..."
         Height          =   315
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   690
         Width           =   1770
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Top             =   705
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cr.A/c Head:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   9
         Top             =   705
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   8
         Top             =   690
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   18225
      TabIndex        =   0
      Top             =   0
      Width           =   18285
   End
End
Attribute VB_Name = "frmViewChequeRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private objCr As New clsAccounts
    Private objBk As New clsBank
    Private Sub cmdAdvancedSearch_Click()
        Call FillGrid
    End Sub
    Private Sub cmdSearchBank_Click()
        Call PopulateList(lstBanks, " SELECT distinct vchBank, tnyVoucherTypeID From faVouchers WHERE (vchBank IS NOT NULL) AND (tnyVoucherTypeID = 10) ORDER BY vchBank ", , , , True)
        lstBanks.Visible = True
        lstBanks.SetFocus
    End Sub
    Private Sub cmdSearchTransaction_Click()
        If (OptReceipts.value = True) Then
           frmSearchTransactionType.ModeOfTransaction = 1
        End If
        If (OptPayments.value = True) Then
           frmSearchTransactionType.ModeOfTransaction = 2
        End If
        frmSearchTransactionType.Show vbModal
        
        txtTransactionType.Text = Trim(gbSearchStr)
        txtTransactionType.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Sub Form_Load()
        vsGrid.Cell(flexcpFontName, 0) = "Verdana"
        XPC.InitSubClassing
        txtFromDate.Text = DdMmmYy(gbTransactionDate)
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        OptPayments.value = True
        OptReceipts.value = False
        
        Call FillGrid
        Call PopulateList(cmbStatus, "SELECT  vchStatus,intID From faMstStatus  Order By  vchStatus", , True, True, True)
    End Sub
    Private Sub Form_Activate()
''        Me.Top = 500
''        Me.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
    Private Sub FillGrid()
        Dim mSql        As String
        Dim mSqlChild   As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mRowCnt     As Integer
        Dim mRecCnt     As Integer
        Dim mTypeID     As Integer
        Dim mLoop       As Integer
        Dim mStatus     As Integer

        If objdb.SetConnection(mCnn) Then
        
                mSql = "SELECT DISTINCT faVouchers.dtDate, faVouchers.vchBank, faVouchers.vchInstrumentNo, SUM(faVouchers.fltAmount) AS TotalAmount, faAccountHeads.vchAccountHead "
                'mSql = mSql + "  ,faVouchers.dtChequeRealiseDate AS Issued, faReverseEntryChild.intVoucherID AS Dishonoured,"
                'mSql = mSql + "faBankReconciliationEntries.intVoucherID AS Realised"
                mSql = mSql + " FROM faVouchers INNER JOIN"
                mSql = mSql + " faAccountHeads ON faVouchers.intLocalBodyID = faAccountHeads.intLocalBodyID AND"
                mSql = mSql + " faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID "
                mSql = mSql + " INNER JOIN faVoucherChild On faVoucherChild.intVoucherID=faVouchers.intVoucherID"
                mSql = mSql + " LEFT OUTER JOIN faBankReconciliationEntries ON faVouchers.intVoucherID = faBankReconciliationEntries.intVoucherID  LEFT OUTER JOIN"
                mSql = mSql + " faReverseEntryChild ON faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                
                mSql = mSql + " WHERE (faVouchers.intInstrumentTypeID = 5) AND (faVouchers.dtInstrumentDate BETWEEN '" & txtFromDate.Text & " '  AND '" & txtToDate & " ' )"
                If OptPayments.value = True Then
                    mSql = mSql + "And faVouchers.tnyVoucherTypeID = 20 " 'Or (faVouchers.tnyVoucherTypeID = 30 AND faVoucherChild.tnyDebitOrCredit=1)"
                    vsGrid.TextMatrix(0, 4) = "Cr.A/c Head"

                ElseIf OptReceipts.value = True Then
                    mSql = mSql + "And faVouchers.tnyVoucherTypeID = 10" ' Or (faVouchers.tnyVoucherTypeID = 30 AND faVoucherChild.tnyDebitOrCredit=0)"
                    vsGrid.TextMatrix(0, 4) = "Dr.A/c Head"
                End If
                If txtAccountHead.Text <> "" Then
                    mSql = mSql + " and faVouchers.intKeyID1='" & txtAccountHead.Tag & "' "
                End If
                If txtChequeNo.Text <> "" Then
                    mSql = mSql + "and faVouchers.vchInstrumentNo ='" & txtChequeNo.Text & "' "
                End If
                If txtTransactionType.Text <> "" Then
                    mSql = mSql + " and faVouchers.intTransactionTypeID=" & txtTransactionType.Tag & ""
                End If
                If val(txtAmount.Text) <> 0 And val(txtAmountTo.Text) <> 0 Then
                    mSql = mSql + " And faVouchers.fltAmount BETWEEN " & val(txtAmount.Text) & " And " & val(txtAmountTo.Text) & ""
                ElseIf val(txtAmount.Text) > 0 Then
                    mSql = mSql + " And faVouchers.fltAmount = " & val(txtAmount.Text) & ""
                End If
                If txtReceivedBanks.Text <> "" Then
                    mSql = mSql + " And faVouchers.vchBank= '" & txtReceivedBanks.Text & "'"
                End If
                If cmbStatus.ListIndex > 0 Then
                    If cmbStatus.Text = "Dishonoured Cheques" Then
                        mSql = mSql + " And  faReverseEntryChild.intVoucherID IS NOT NULL"
                    ElseIf cmbStatus.Text = "Realised Cheques" Then
                        mSql = mSql + " And  faBankReconciliationEntries.intVoucherNo IS NOT NULL"
                    ElseIf cmbStatus.Text = "Issued Cheques" Then
                        mSql = mSql + " And  faVouchers.dtChequeRealiseDate IS NOT NULL"
                    End If
                End If
                mSql = mSql + " GROUP BY faVouchers.vchBank, faVouchers.vchInstrumentNo, faVouchers.dtDate, faAccountHeads.vchAccountHead "
                'mSql = mSql + " ,faReverseEntryChild.intVoucherID, faBankReconciliationEntries.intVoucherID, faVouchers.dtChequeRealiseDate"
                mSql = mSql + " Order By faVouchers.dtdate,faVouchers.vchInstrumentNo,faVouchers.vchBank"
                Rec.CursorLocation = adUseClient
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                
                vsGrid.Clear 1, 1
                mRowCnt = 1
                vsGrid.Rows = 1
                vsGrid.OutlineBar = flexOutlineBarCompleteLeaf
                If Not (Rec.BOF And Rec.EOF) Then
                    While Not (Rec.EOF)
                        vsGrid.Rows = vsGrid.Rows + 1
                        
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = mRowCnt
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!TotalAmount), "", Rec!TotalAmount)
                        'VSGrid.TextMatrix(VSGrid.Rows - 1, 5) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                        'VSGrid.TextMatrix(VSGrid.Rows - 1, 6) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                        'VSGrid.TextMatrix(VSGrid.Rows - 1, 7) = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                        vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
                        vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 0
                        For mLoop = 0 To vsGrid.Cols - 1
                            vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HD2AE9E
                        Next mLoop
                        mSqlChild = "SELECT DISTINCT faVouchers.dtDate, faVouchers.vchBank, faVouchers.vchInstrumentNo, faVouchers.dtInstrumentDate, faVouchers.fltAmount, faVouchers.intVoucherNo,"
                        mSqlChild = mSqlChild + " faTransactionType.vchTransactionType , faVouchers.intVoucherID "
                        mSqlChild = mSqlChild + "  ,faVouchers.dtChequeRealiseDate AS Issued, faReverseEntryChild.intVoucherID AS Dishonoured,"
                        mSqlChild = mSqlChild + "faBankReconciliationEntries.intVoucherID AS Realised"
                        mSqlChild = mSqlChild + " FROM faVouchers INNER JOIN"
                        mSqlChild = mSqlChild + " faTransactionType ON faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
                        mSqlChild = mSqlChild + " LEFT OUTER JOIN faBankReconciliationEntries ON faVouchers.intVoucherID = faBankReconciliationEntries.intVoucherID  LEFT OUTER JOIN"
                        mSqlChild = mSqlChild + " faReverseEntryChild ON faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                        mSqlChild = mSqlChild + " Where (faVouchers.intInstrumentTypeID = 5)AND faVouchers.vchInstrumentNo='" & vsGrid.TextMatrix(vsGrid.Rows - 1, 2) & "' AND faVouchers.vchBank='" & vsGrid.TextMatrix(vsGrid.Rows - 1, 3) & "'"
                        
                        'msqlchild = msqlchild + "And faVouchers.intKeyID1='" & txtAccountHead.Tag & "' "
                        If OptPayments.value = True Then
                           ' msqlchild = msqlchild + "And faVouchers.intKeyID1='" & txtAccountHead.Tag & "' "
                            mSqlChild = mSqlChild + "And tnyVoucherTypeID = 20"
                        ElseIf OptReceipts.value = True Then
                            'msqlchild = msqlchild + "And faVouchers.intKeyID1='" & txtAccountHead.Tag & "' "
                            mSqlChild = mSqlChild + "And tnyVoucherTypeID = 10"
                        End If
                        If val(txtAmount.Text) <> 0 And val(txtAmountTo.Text) <> 0 Then
                            mSqlChild = mSqlChild + " And faVouchers.fltAmount BETWEEN " & val(txtAmount.Text) & " And " & val(txtAmountTo.Text) & ""
                        ElseIf val(txtAmount.Text) > 0 Then
                            mSqlChild = mSqlChild + " And faVouchers.fltAmount = " & val(txtAmount.Text) & ""
                        End If
'''''                        If cmbStatus.ListIndex > 0 Then
'''''                            If cmbStatus.Text = "Dishonoured Cheques" Then
'''''                                mSql = mSql + " And  faReverseEntryChild.intVoucherID IS NOT NULL"
'''''                            ElseIf cmbStatus.Text = "Realised Cheques" Then
'''''                                mSql = mSql + " And  faBankReconciliationEntries.intVoucherNo IS NOT NULL"
'''''                            ElseIf cmbStatus.Text = "Issued Cheques" Then
'''''                                mSql = mSql + " And  faVouchers.dtChequeRealiseDate IS NOT NULL"
'''''                            End If
'''''                        End If
                        RecChild.Open mSqlChild, mCnn, adOpenStatic, adLockOptimistic
                        If Not (RecChild.BOF And RecChild.EOF) Then
                        'MsgBox Rec!vchInstrumentNo
                            While Not (RecChild.EOF)
                                'VSGrid.IsCollapsed(VSGrid.Rows - 1) = flexOutlineCollapsed
                                vsGrid.Rows = vsGrid.Rows + 1
                                
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount)
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(RecChild!intVoucherNo), "", RecChild!intVoucherNo)
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = IIf(IsNull(RecChild!vchTransactionType), "", RecChild!vchTransactionType)
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = IIf(IsNull(RecChild!dtInstrumentDate), "", RecChild!dtInstrumentDate)
                                'VSGrid.TextMatrix(VSGrid.Rows - 1, 13) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                                If Not IsNull(RecChild!Issued) Then
                                    vsGrid.TextMatrix(vsGrid.Rows - 1, 9) = "Issued Cheques"
                                    vsGrid.TextMatrix(vsGrid.Rows - 1, 10) = 1
                                End If
                                If Not IsNull(RecChild!Dishonoured) Then
                                    vsGrid.TextMatrix(vsGrid.Rows - 1, 9) = "Dishonoured Cheques"
                                    vsGrid.TextMatrix(vsGrid.Rows - 1, 11) = 2
                                End If
                                If Not IsNull(RecChild!Realised) Then
                                    vsGrid.TextMatrix(vsGrid.Rows - 1, 9) = "Realised Cheques"
                                    vsGrid.TextMatrix(vsGrid.Rows - 1, 12) = 3
                                End If
                             
                                vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
                                vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 1
                                RecChild.MoveNext
                            Wend
                        End If
                        RecChild.Close
                        Rec.MoveNext
                        mRowCnt = mRowCnt + 1
                    Wend
                End If
                Rec.Close
        End If
    End Sub
    Private Sub lstBanks_LostFocus()
        lstBanks.Visible = False
    End Sub
    Private Sub OptPayments_Click()
        Call FillGrid
        Label5.Caption = "Cr.A/c Head:"
        txtReceivedBanks.Enabled = False
        cmdSearchBank.Enabled = False
    End Sub
    Private Sub OptReceipts_Click()
        Call FillGrid
        Label5.Caption = "Dr.A/c Head:"
        txtReceivedBanks.Enabled = True
        cmdSearchBank.Enabled = True
    End Sub
    Private Sub txtAccountHead_LostFocus()
        Call FillGrid
    End Sub

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub


    Private Sub txtAmountTo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub

    Private Sub txtFromDate_LostFocus()
        If Not IsDate(txtFromDate.Text) Then
            txtFromDate.Text = DdMmmYy(gbStartingDate)
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub



    Private Sub txtToDate_LostFocus()
        If Not IsDate(txtToDate.Text) Then
            txtToDate.Text = DdMmmYy(gbTransactionDate)
        Else
            txtToDate.Text = CheckDateInMMM(Trim(txtToDate))
        End If
        
        If txtFromDate.Text <> "" Then
            Call FillGrid
        End If
    End Sub
    Private Sub cmdAccoundHeads_Click()
            Call txtAccountCode_KeyDown(vbKeyF4, 0)
    End Sub
    Private Sub txtAccountCode_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtAccountCode.Text = Token(gbSearchStr, " ")
            txtAccountHead.Text = Trim(gbSearchStr)
            txtAccountHead.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
        txtAccountCode.SelStart = 0
        txtAccountCode.SelLength = Len(txtAccountCode)
    End Sub
    Private Sub txtAccountCode_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            Call ShowSearchAccountHead
        End If
    End Sub
    Private Sub txtAccountCode_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    Private Sub txtAccountCode_LostFocus()
        Dim mChequeNo As Variant
        objCr.SetAccountCode Trim(txtAccountCode.Text)
        If objCr.AccountHeadID > 0 Then
            txtAccountHead.Text = objCr.AccountHead
            txtAccountCode.Text = objCr.AccountCode
            objBk.SetBankInfoByAccID objCr.AccountHeadID
            If objBk.BankAccountHeadID > -1 Then
                'txtNameOfBank.Text = objBk.BankName
                'txtBranch.Text = objBk.Branch
                'txtAccountNo.Text = objBk.AccountNumber
                'mChequeNo = objBk.GetNeWChequeNumber
                'txtRef.Text = IIf(IsNull(mChequeNo), "", mChequeNo)
            Else
                'txtNameOfBank.Text = ""
                'txtBranch.Text = ""
                'txtAccountNo.Text = ""
                'txtRef.Text = ""
            End If
        Else
            txtAccountHead.Text = ""
            txtAccountCode.Text = ""
        End If
    End Sub
    Private Sub txtAccountHead_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            Call txtAccountCode_KeyDown(vbKeyF4, 0)
        End If
        If KeyCode = vbKeyDelete Then
            txtAccountHead.Text = ""
            txtAccountCode.Text = ""
            txtAccountCode.Tag = ""
        End If
    End Sub
'''    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
'''        If KeyAscii = 13 Then PressTabKey
'''    End Sub
    Private Sub ShowSearchAccountHead()
        Dim mSql As String
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.tinHiddenFlag = 0 AND faAccountHeads.intGroupID =" & faBank
            frmSearchAccountHeads.VoucherMode = 300
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.Show vbModal
            txtAccountCode.SetFocus
    End Sub
    Private Sub lstBanks_DblClick()
        Call lstBanks_KeyDown(13, 0)
    End Sub
    Private Sub lstBanks_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then
            txtReceivedBanks.Text = lstBanks.Text
            lstBanks.Visible = False
        End If
    End Sub

''SELECT     faVouchers.dtDate, faVouchers.vchBank, faVouchers.vchInstrumentNo, SUM(faVouchers.fltAmount) AS TotalAmount,
''                      faAccountHeads.vchAccountHead
''FROM         faVouchers INNER JOIN
''                      faAccountHeads ON faVouchers.intLocalBodyID = faAccountHeads.intLocalBodyID AND
''                      faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID
''WHERE     (faVouchers.intInstrumentTypeID = 5) AND (faVouchers.dtInstrumentDate BETWEEN '01-Apr-2010 ' AND '28-Jul-2010 ') AND
''                      (faVouchers.tnyVoucherTypeID = 10)
''GROUP BY faVouchers.vchBank, faVouchers.vchInstrumentNo, faVouchers.dtDate, faAccountHeads.vchAccountHead
''ORDER BY faVouchers.vchInstrumentNo, faVouchers.vchBank



    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer) 'To delete value in txtTransactionType
        If KeyCode = vbKeyDelete Then
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        End If
    End Sub
