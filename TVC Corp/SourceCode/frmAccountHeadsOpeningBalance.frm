VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAccountHeadsOpeningBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opening Balance Entry"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14415
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAccountHeadsOpeningBalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApprove 
      Caption         =   "&Approve"
      Height          =   405
      Left            =   10410
      TabIndex        =   13
      Top             =   6810
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtTotalDrAmount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12825
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   6345
      Width           =   1545
   End
   Begin VB.TextBox txtCapitalAccountHeadDr 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   7245
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6345
      Width           =   5550
   End
   Begin VB.TextBox txtTotalCrAmount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   6345
      Width           =   1545
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   6720
      TabIndex        =   8
      Top             =   6825
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   11730
      TabIndex        =   7
      Top             =   6780
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Clos&E"
      Height          =   405
      Left            =   13020
      TabIndex        =   6
      Top             =   6780
      Width           =   1245
   End
   Begin VB.TextBox txtCapitalAccountHeadCr 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6345
      Width           =   5595
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   30
      Top             =   7260
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGridCredit 
      Height          =   4905
      Left            =   120
      TabIndex        =   0
      Top             =   1110
      Width           =   6885
      _cx             =   12144
      _cy             =   8652
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      BackColorBkg    =   16777215
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   17
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAccountHeadsOpeningBalance.frx":1CCA
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
      Editable        =   2
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
   Begin VSFlex8LCtl.VSFlexGrid vsGridDebit 
      Height          =   4905
      Left            =   7095
      TabIndex        =   3
      Top             =   1125
      Width           =   7320
      _cx             =   12912
      _cy             =   8652
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      BackColorBkg    =   16777215
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   17
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAccountHeadsOpeningBalance.frx":1D52
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
      Editable        =   2
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Assets"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6975
      TabIndex        =   15
      Top             =   705
      Width           =   7395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liabilities"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   60
      TabIndex        =   14
      Top             =   690
      Width           =   6900
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Credited to Capital Fund"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7260
      TabIndex        =   11
      Top             =   6105
      Width           =   3165
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Debited to Capital Fund"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   6105
      Width           =   3150
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account Heads"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   375
      Width           =   14430
   End
   Begin VB.Label lblBalanceDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Balances As on31/Mar/ ----"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   14445
   End
End
Attribute VB_Name = "frmAccountHeadsOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mTransactionDate As String
    Dim mFinancialYear As Integer
    Dim mVoucherNo As Variant
    Dim mFundID As Integer
    Dim mFundCode As String
    Dim mMinFinancialYear As Integer
    Dim mCrVoucherNo    As Variant
    Dim mDrVoucherNo    As Variant
    Dim mTransnsactionTypeID As Variant
    Dim mintFinancialYearID As Integer
    Private intMode As Integer   ' 1= save  2= Approve
    Private Sub cmdApprove_Click()
        cmdApprove.Enabled = False
    End Sub
    Private Sub cmdClear_Click()
        Call FormInitialize
    End Sub
    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    Private Sub cmdSave_Click()
        Dim mSql As String
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mCnn.Execute "Update faAccountHeads Set fltOpeningBalance = 0, tinDebitOrCredit = Null Where fltOpeningBalance > 0"
'''            If vsGridCredit.Rows < 1 Or vsGridDebit.Rows < 1 = True Then
'''                MsgBox "Please enter details,vbInformation"
'''                Exit Sub
'''            Else end if
        If vsGridCredit.Row = 0 Then
            MsgBox " No Credit data is entered", vbInformation
            Exit Sub
        ElseIf vsGridDebit.Row = 0 Then
            MsgBox " No Debit data is entered", vbInformation
            Exit Sub
        Else
                'If val(vsGridCredit.TextMatrix(1, 0)) = 0 And val(vsGridCredit.TextMatrix(1, 3)) = 0 And _
                '    val(vsGridDebit.TextMatrix(1, 0)) = 0 And val(vsGridDebit.TextMatrix(1, 3)) = 0 Then
                '    MsgBox "Please fill required data", vbInformation
                '    Exit Sub
                'End If
                
'                If val(vsGridCredit.TextMatrix(1, 1)) <> 0 And Trim(vsGridCredit.TextMatrix(1, 2)) <> "" And val(vsGridCredit.TextMatrix(1, 3)) <> 0 Then
                    Call SaveData(vsGridCredit)
'                Else
'                    'MsgBox "Please Enter the Required Data"
'                End If
'                If val(vsGridDebit.TextMatrix(1, 1)) <> 0 And Trim(vsGridDebit.TextMatrix(1, 2)) <> "" And val(vsGridDebit.TextMatrix(1, 3)) <> 0 Then
                    Call SaveData(vsGridDebit)
'                Else
'                    'MsgBox " Please Enter the Required Data "
'                End If
                MsgBox "Opening Balance Entry Done Successfully", vbInformation
                cmdSave.Enabled = False
         End If
            
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        WindowsXPC1.InitIDESubClassing
        vsGridCredit.ColComboList(1) = "|..."
        vsGridDebit.ColComboList(1) = "|..."
        Call FundWiseHead
        Call SetFinancialYear
        Call Calculate
       
        If gbSeatGroupID = gbSeatGroupAccountsSuperintended Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSave.Enabled = True
            cmdApprove.Visible = False
            'cmdApprove.Enabled = True
        Else
            cmdSave.Enabled = False
            cmdApprove.Visible = False
           
                 'cmdApprove.Enabled = False
        End If
        'cmdSave.Enabled = True ' Changed by Aiby For Immediate Release of Version 2.2.4
        ''772 added on 14 sep 2017 for Presentation
         If gbLBID = 1260 Or gbLBID = 1261 Or gbLBID = 1262 Or gbLBID = 1263 Or gbLBID = 1264 Or gbLBID = 772 _
                 Or gbLBID = 1265 Or gbLBID = 1266 Or gbLBID = 1267 Or gbLBID = 1268 Or gbLBID = 1269 _
                 Or gbLBID = 1270 Or gbLBID = 1271 Or gbLBID = 1272 Or gbLBID = 1273 Or gbLBID = 1274 _
                 Or gbLBID = 1275 Or gbLBID = 1276 Or gbLBID = 1277 Or gbLBID = 1278 Or gbLBID = 1279 _
                 Or gbLBID = 1280 Or gbLBID = 1281 Or gbLBID = 1282 Or gbLBID = 1283 Or gbLBID = 1284 _
                 Or gbLBID = 1285 Or gbLBID = 1286 Or gbLBID = 1287 Or gbLBID = 1259 Or gbLBID = 167 Then
                 
                 cmdSave.Enabled = True
            Else
                 cmdSave.Enabled = False
            End If
        Call FillGrid(vsGridCredit)
        Call FillGrid(vsGridDebit)
        If intMode > 2 Then
            cmdApprove.Visible = False
'        Else
'            cmdApprove.Visible = True
         End If
                 
    End Sub
    Private Sub FundWiseHead()
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mYearID As Variant
        Dim arrIn As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        Rec.Open "Select vchFundCode as Fund,  intFundID from faFunds Where tnyActiveFlag=1", mCnn, adOpenDynamic, adLockReadOnly
        Rec.Open " SELECT faFunds.intFundID,faFunds.vchFundCode,faSeats.numSeatID FROM faFunds INNER JOIN faSeats ON faSeats.intFundID = faFunds.intFundID Where faSeats.numSeatID =" & gbSeatID, mCnn, adOpenDynamic, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            mFundID = Rec!intFundID
            mFundCode = Rec!vchFundCode
        End If
        Rec.Close
        mSql = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads Where vchAccountHeadCode ='" & gbAcHeadCodeForCapitalFund & " '"
'        mSql = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads Where vchAccountHeadCode =" & Left(gbAcHeadCodeForCapitalFund, 6) & Right(mFundCode, 3)
        Rec.Open mSql, mCnn, adOpenDynamic, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            txtCapitalAccountHeadCr.Text = Rec!vchAccountHeadCode
            txtCapitalAccountHeadCr.Tag = Rec!intAccountHeadID
            txtCapitalAccountHeadDr.Text = Rec!vchAccountHeadCode
            txtCapitalAccountHeadDr.Tag = Rec!intAccountHeadID
        End If
    End Sub
    Private Sub GenerateVoucherNo(mCrDrFlag As Integer)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If mCrDrFlag = 0 Then
'            mCrVoucherNo = "4" & Right(mFinancialYear, 2) & Right(mFundCode, 3) & "000"
            mCrVoucherNo = "4" & Right(mintFinancialYearID, 2) & Right(mFundCode, 3) & "000"
        Else
'            mDrVoucherNo = "4" & Right(mFinancialYear, 2) & Right(mFundCode, 3) & "001"
            mDrVoucherNo = "4" & Right(mintFinancialYearID, 2) & Right(mFundCode, 3) & "001"
            
        End If
        
    End Sub
'    Private Sub ApproveData(vsGrid As VSFlexGrid)
'        Dim objDB As New clsDb
'        Dim Rec As New ADODB.Recordset
'        Dim mCnn As New ADODB.Connection
'        Dim mSql As String
'        If vsGrid.Name = vsGridCredit.Name Then
'            mVoucherNo = mCrVoucherNo
'        ElseIf vsGrid.Name = vsGridDebit.Name Then
'            mVoucherNo = mDrVoucherNo
'        End If
'        If objDB.SetConnection(mCnn) Then
'            mSql = "Update faVouchers Set tnyStatus = 1 where intVoucherNo = " & mVoucherNo
'            mCnn.Execute mSql
'            cmdApprove.Enabled = False
'            cmdSave.Enabled = False
'        End If
'    End Sub
    Private Sub SaveData(vsGrid As VSFlexGrid)
        Dim mLoop           As Long
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim arrInput        As Variant
        Dim arrOutPut       As Variant
        Dim arrOut          As Variant
        Dim mAmount         As Double
        Dim mintKeyID       As Long
        Dim mumVoucherNo    As Variant
        Dim mVoucherID      As Double
        Dim mTransactionID  As Double
        Dim mDrCr           As Integer
        Dim mSql            As String
        Dim mCrDrAmt        As Double
        Dim voucher         As uVoucher
'On Error GoTo ErrRollBack:
        If vsGrid.Name = vsGridCredit.Name Then
            mAmount = val(txtTotalCrAmount.Text)
            mintKeyID = val(txtCapitalAccountHeadCr.Tag)
            mDrCr = 0
            mVoucherNo = mCrVoucherNo
        ElseIf vsGrid.Name = vsGridDebit.Name Then
             mAmount = val(txtTotalDrAmount.Text)
             mintKeyID = val(txtCapitalAccountHeadDr.Tag)
             mDrCr = 1
             mVoucherNo = mDrVoucherNo
        End If
        objdb.SetConnection mCnn
        Call GenerateVoucherNo(mDrCr)
        mSql = "Select intVoucherID from faVouchers where intVoucherNo = " & mVoucherNo
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mVoucherID = Rec!intVoucherID
        Else
            mVoucherID = -1
        End If
        Rec.Close
        
                        '------------------------------------------------'
                        '                faVouchers                      '
                        '------------------------------------------------'
        With voucher
            .intVoucherID_1 = mVoucherID
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = 3000
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = mTransactionDate
            .fltAmount_9 = mAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            .vchDescription_13 = "OpeningBalance"
            .numZoneID_14 = gbnumZonalID
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = mintKeyID
            .intKeyID2_23 = Null
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 40
            .intFinancialYearID_26 = mMinFinancialYear
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = mFundID
            .numSeatID = gbSeatID
            .intSessionID = Null
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = Null
            .numLocationID = Null
        
        arrInput = Array(.intVoucherID_1, _
                                .intLocalBodyID_2, _
                                .intTransactionID_3, _
                                .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
                                .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
                                .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
                                .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
                                .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
                                .intKeyID2_23, .intExternalApplicationID_24, _
                                .intExternalModuleID_25, .intFinancialYearID_26, _
                                .tnyShiftID_27, .tnyPrintFlag_28, _
                                .tnyCancelFlag_29, .vchBank_33, _
                                .vchBankPlace_34, .intFundID_35, _
                                .numSeatID, .intSessionID, _
                                .vchRefNo, .fltRoundOff, _
                                .fltAdvAmtAdj, .numInwardNo, _
                                .tnyStatus_32, .numLocationID)
                
        
        objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn, adCmdStoredProc
            If IsNumeric(arrOutPut(0, 0)) Then
                    mVoucherID = arrOutPut(0, 0)
                    '''If Not IsError((arrOutPut(1, 0))) Then
                    '''    mVoucherNo = arrOutPut(1, 0)
                    '''End If
            Else
                    'GoTo ErrRollBack:
            End If
        End With

                        '------------------------------------------------'
                            'faVoucherChild Parameters
                        '------------------------------------------------'
        Dim mSlNo                 As Long
        Dim mtnyDebitOrCredit     As Integer
        Dim mintYearID            As Long
        Dim mtnyPeriodID          As Integer
        Dim mtnyArrearFlag        As Integer
        Dim vChild                As uVChild
        
        'If intMode = 2 Then
        mCnn.Execute "Delete From faVoucherChild Where intVoucherID =" & mVoucherID
        'Else
        With vChild
        If vsGrid.Name = vsGridCredit.Name Then
            mtnyDebitOrCredit = 0
        Else
            mtnyDebitOrCredit = 1
         End If
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 1) <> "" Then
                .intVoucherID_1 = mVoucherID
                .intLocalBodyID_2 = gbLocalBodyID
                .intSlNo_3 = mLoop
                .intAccountHeadID_4 = vsGrid.TextMatrix(mLoop, 0)
                .tnyDebitOrCredit_5 = mtnyDebitOrCredit
                .intYearID_6 = Null
                .tnyPeriodID_7 = Null
                .tnyArrearFlag_8 = Null
                .numDemandID_9 = Null
                .fltAmount_10 = val(vsGrid.TextMatrix(mLoop, 3))
                
            arrInput = Array( _
                            .intVoucherID_1, _
                            .intLocalBodyID_2, _
                            .intSlNo_3, _
                            .intAccountHeadID_4, _
                            .tnyDebitOrCredit_5, _
                            .intYearID_6, _
                            .tnyPeriodID_7, _
                            .tnyArrearFlag_8, _
                            .numDemandID_9, _
                            .fltAmount_10 _
                     )
            objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
            mSlNo = mSlNo + 1
            End If
          
                     '----------------------------------------'
                     '              Update faAccountHeads     '
                     '----------------------------------------'
            If val(vsGrid.TextMatrix(mLoop, 3)) <> 0 Then
            
                If vsGrid.Name = vsGridCredit.Name Then
                    mtnyDebitOrCredit = 0
                Else
                    mtnyDebitOrCredit = 1
                End If
                mSql = "UPDATE    faAccountHeads SET    fltOpeningBalance = " & val(vsGrid.TextMatrix(mLoop, 3)) & ",tinDebitOrCredit=" & mtnyDebitOrCredit & " Where intAccountHeadID=" & vsGrid.TextMatrix(mLoop, 0) & "  "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            End If
           Next mLoop
        End With
                    

                        '-------------------------------------'
                        ' Data for Transaction Table          '
                        '-------------------------------------'
        Dim Trans As uTr
        Dim mSqlt As String
        mSqlt = "Select intTransactionID from faTransactions where intVoucherID =" & mVoucherID
        Rec.Open mSqlt, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mTransactionID = Rec!intTransactionID
        Else
            mTransactionID = -1
        End If
        With Trans
            .intTransactionID = mTransactionID
            .intLocalBodyID = gbLocalBodyID
            .intFinancialYearID = mMinFinancialYear
            .dtTransactionDate = Format((mTransactionDate), "DD/MMM/YYYY")
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = 0
            .intFunctionID = Null
            .intFunctionaryID = Null
            .intFieldID = Null
            .intFundID = mFundID
            .intBudgetCentreID = Null
            .vchNarration = "Opening Balance"
            .intTransactionTypeID = 3000
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mVoucherID
            
             arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
                                                                                          
                                            
          objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
                 If IsNumeric(arrOutPut(0, 0)) Then
                    mTransactionID = arrOutPut(0, 0)
                 Else
                    'GoTo ErrRollBack:
                 End If
    End With

                        '----------------------------------------'
                        ' Data for TransactionChild    '
                        '----------------------------------------'
      Dim transChild As uTrChild
      
      'If intMode = 2 Then
      mCnn.Execute "Delete From faTransactionChild Where intTransactionID =" & mTransactionID
      'Else
      With transChild
              If vsGrid.Name = vsGridCredit.Name Then
                mtnyDebitOrCredit = 1
              Else
                mtnyDebitOrCredit = 0
              End If
              
                .intTransactionID = mTransactionID
                .intSerialNo = 1
                .intAccountHeadID = mintKeyID
                .fltAmount = mAmount
                .tinDebitOrCreditFlag = mtnyDebitOrCredit
                .intByAccountHeadID = Null
                .vchNarration = "Opening Balance"
                .intFundID = mFundID
                
                arrInput = Array(.intTransactionID, _
                                 .intSerialNo, _
                                 .intAccountHeadID, _
                                 .fltAmount, _
                                 .tinDebitOrCreditFlag, _
                                 .intByAccountHeadID, _
                                 .vchNarration, _
                                 .intFundID)
                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                If mtnyDebitOrCredit = 1 Then
                    mtnyDebitOrCredit = 0
                Else
                    mtnyDebitOrCredit = 1
                End If
    End With

                    '----------------------------------------'
                    ' Data for TransactionChild From Grid   '
                    '----------------------------------------'
       With transChild
               mSlNo = 1
               For mLoop = 1 To vsGrid.Rows - 1
                 If vsGrid.TextMatrix(mLoop, 1) <> "" Then
                     mSlNo = mSlNo + 1
                 
                        .intTransactionID = mTransactionID
                        .intSerialNo = mSlNo
                        .intAccountHeadID = vsGrid.TextMatrix(mLoop, 0)
                        .fltAmount = vsGrid.TextMatrix(mLoop, 3)
                        .tinDebitOrCreditFlag = mtnyDebitOrCredit
                        .intByAccountHeadID = mintKeyID
                        .vchNarration = "OPening Balance"
                        .intFundID = mFundID
                        
                   arrInput = Array(.intTransactionID, _
                                 .intSerialNo, _
                                 .intAccountHeadID, _
                                 .fltAmount, _
                                 .tinDebitOrCreditFlag, _
                                 .intByAccountHeadID, _
                                 .vchNarration, _
                                 .intFundID _
                                 )
                                   
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
                 End If
               Next mLoop
      End With

                    '----------------------------------------'
                    '              Update VoucherNo          '
                    '----------------------------------------'
                Call GenerateVoucherNo(mDrCr)
                mSql = "Update faVouchers Set intVoucherNo = " & mVoucherNo & " Where intVoucherID=" & mVoucherID
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
'''''''                 '----------------------------------------'
'''''''                 '              Update faAccountHeads     '
'''''''                 '----------------------------------------'
'''''''                mSQL = "UPDATE    faAccountHeads SET    fltOpeningBalance = Where vchAccountHeadCode"
'''''''                objDB.ExecuteSP mSQL, , , , mcnn, adCmdText
    Exit Sub
ErrRollBack:
    MsgBox "Saankhya Error" & err.Description
'    End If
    End Sub
                   '-------------------------------------------------------------'
                   'To display As on date on the screen- and to set FinancialYear'
                   '-------------------------------------------------------------'
    Private Sub SetFinancialYear()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mCount As Integer
        Dim mStartingDate As String
        Dim mEndingDate As String
        Dim mNewstartDate As String
        Dim mNewEnddate As String
        Dim mLbID As Integer
        Dim mYearID As Integer
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "SELECT count(*) Ccount from faFinancialYear"
        Rec.Open mSql, mCnn, adOpenDynamic, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            mCount = Rec!Ccount
        End If
        Rec.Close
        If mCount = 1 Then
            mSql = "SELECT dtStartingDate,dtEndingDate,intLocalBodyID from faFinancialYear"
            Rec.Open mSql, mCnn, adOpenDynamic, adLockReadOnly
            If Not (Rec.EOF And Rec.BOF) Then
                mStartingDate = Rec!dtStartingDate
                mEndingDate = Rec!dtEndingDate
                mNewstartDate = DateAdd("yyyy", -1, mStartingDate)
                mNewstartDate = Format(mNewstartDate, "DD/MMM/YYYY")
                mNewEnddate = DateAdd("yyyy", -1, mEndingDate)
                mNewEnddate = Format(mNewEnddate, "DD/MMM/YYYY")
                mLbID = Rec!intLocalBodyID
            End If
            Rec.Close
            'MsgBox mLbID, vbInformation
            mYearID = gbFinancialYearID - 1
'            Rec.Open "INSERT faFinancialYear(intFinancialYearID,intFinancialYear,dtStartingDate,dtEndingDate,dtLastTransactionDate,tinCurrentFinancialYearFlag,intLocalBodyID) VALUES('gbFinancialYearID'-1,'gbFinancialYearID'-1,'mNewstartDate','mNewEnddate','mNewEnddate','0',mLBID)", mCnn, adOpenDynamic, adLockReadOnly
            mSql = "INSERT into faFinancialYear(intFinancialYearID,intFinancialYear,dtStartingDate,dtEndingDate,dtLastTransactionDate,tinCurrentFinancialYearFlag,intLocalBodyID) VALUES(" & mYearID & "," & mYearID & ",'" & mNewstartDate & "','" & mNewEnddate & "', '" & mNewEnddate & "' ," & 0 & "," & gbLocalBodyID & ")"
'            Rec.Open mSql, mCnn, adOpenDynamic, adLockReadOnly
            mCnn.Execute mSql
'            Rec.Close
        End If
        'Rec.Open "select Min(intFinancialYearID)as[FinancialYear] from faFinancialYear", mCnn, adOpenDynamic, adLockReadOnly
        Rec.Open "Select Min(dtEndingDate) FinancialYear From faFinancialYear", mCnn, adOpenDynamic, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            'mMinFinancialYear = Rec!FinancialYear
            'mFinancialYear = (Rec!FinancialYear)
            'mTransactionDate = ("31/Mar/ " & Rec!FinancialYear)
            'lblBalanceDate.Caption = " Balances as on 31/Mar/ " & Rec!FinancialYear
            
            mMinFinancialYear = Year(Rec!FinancialYear)
            If (Month(Rec!FinancialYear)) < 4 Then
               mMinFinancialYear = mMinFinancialYear - 1
            Else
                mMinFinancialYear = mMinFinancialYear
            End If
            mFinancialYear = Year(Rec!FinancialYear)
            mintFinancialYearID = Year(Rec!FinancialYear) - 1
            mTransactionDate = Rec!FinancialYear
            lblBalanceDate.Caption = " Balances as on " & DdMmmYy(Rec!FinancialYear)
        End If
        
    End Sub
    Private Sub FormInitialize()
        vsGridCredit.Clear (1)
        vsGridDebit.Clear (1)
        txtTotalDrAmount.Text = ""
        txtTotalCrAmount.Text = ""
    End Sub

    Private Sub vsGridCredit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If val(vsGridCredit.TextMatrix(Row, 1)) <> 0 And vsGridCredit.TextMatrix(Row, 2) <> "" And val(vsGridCredit.TextMatrix(Row, 3)) <> 0 Then
            vsGridCredit.Rows = vsGridCredit.Rows + 1
        End If
        vsGridCredit.TextMatrix(Row, 3) = Format(val(vsGridCredit.TextMatrix(Row, 3)), "0.00")
        If Len(vsGridCredit.TextMatrix(Row, 3)) > 15 Then
            MsgBox "Please check the Amount", vbInformation
            vsGridCredit.TextMatrix(Row, 3) = ""
        End If
        If val(vsGridCredit.TextMatrix(Row, 3)) < 0 Then
            vsGridCredit.TextMatrix(Row, 3) = ""
        End If
        Call Calculate
    End Sub
                    '----------------------------------------'
                    ' To search the  Liability Account Heads     '
                    '----------------------------------------'
    Private Sub vsGridCredit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim mToken As String
'        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinHiddenFlag <> 1 And intGroupID is Null Order by vchAccountHeadCode"
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType = 3 And vchAccountHeadCode Not IN('300100100','310100101') AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
        frmSearchAccountHeads.Show vbModal
        If gbSearchID <> -1 Then
            If vsGridCredit.FindRow(gbSearchID, 1, 0) > 0 Or vsGridDebit.FindRow(gbSearchID, 1, 0) > 0 Then
                MsgBox "Already selected this Account Head"
                Exit Sub
            End If
            vsGridCredit.TextMatrix(vsGridCredit.Row, 0) = gbSearchID
            vsGridCredit.TextMatrix(vsGridCredit.Row, 1) = Token(gbSearchStr, " ")
            vsGridCredit.TextMatrix(vsGridCredit.Row, 2) = Trim(gbSearchStr)
            gbSearchID = -1
            gbSearchStr = ""
            vsGridCredit.Rows = vsGridCredit.Rows + 1
        End If
    End Sub
    Private Sub vsGridCredit_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
'                Call DeleteRows(vsGridCredit)
                vsGridCredit.RemoveItem (vsGridCredit.Row)
            End If
        End If
        Call Calculate   'Added on 19/08/2011
    End Sub
    
    Private Sub vsGridCredit_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        End If
    End Sub

    Private Sub vsGridCredit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If KeyAscii > 0 And (Col = 1 Or Col = 2) Then KeyAscii = 0
        If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47 And Col = 3) Then KeyAscii = 0
    End Sub
    
    Private Sub vsGridDebit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If val(vsGridDebit.TextMatrix(Row, 1)) <> 0 And vsGridDebit.TextMatrix(Row, 2) <> "" And val(vsGridDebit.TextMatrix(Row, 3)) <> 0 Then
            vsGridDebit.Rows = vsGridDebit.Rows + 1
        End If
        vsGridDebit.TextMatrix(Row, 3) = Format(val(vsGridDebit.TextMatrix(Row, 3)), "0.00")
        If Len(vsGridDebit.TextMatrix(Row, 3)) > 15 Then
            MsgBox "Please check the Amount", vbInformation
            vsGridDebit.TextMatrix(Row, 3) = ""
        End If
        If val(vsGridDebit.TextMatrix(Row, 3)) < 0 Then
            vsGridDebit.TextMatrix(Row, 3) = ""
        End If
        Call Calculate

    End Sub

    Private Sub vsGridDebit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsGridDebit.Row - 1 <> 0 Then
            If vsGridDebit.TextMatrix(Row - 1, Col) = "" Then
                Cancel = True
            End If
        End If
        Call Calculate
    End Sub
    
    Private Sub vsGridCredit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsGridCredit.Row - 1 <> 0 Then
            If vsGridCredit.TextMatrix(Row - 1, Col) = "" Then
                Cancel = True
            End If
        End If
        Call Calculate
    End Sub
    
                    '----------------------------------------'
                    ' To search the  Asset Account Heads '
                    '----------------------------------------'
    Private Sub vsGridDebit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead)  as AccHead, intAccountHeadID From faAccountHeads Where tinType = 4 AND tinHiddenFlag = 0 ORDER BY vchAccountHeadCode"
        frmSearchAccountHeads.Show vbModal
        If gbSearchID <> -1 Then
            If vsGridDebit.FindRow(gbSearchID, 1, 0) > 0 Or vsGridCredit.FindRow(gbSearchID, 1, 0) > 0 Then
                MsgBox "Already selected this Account Head"
                Exit Sub
            End If
            vsGridDebit.TextMatrix(vsGridDebit.Row, 0) = gbSearchID
            vsGridDebit.TextMatrix(vsGridDebit.Row, 1) = Token(gbSearchStr, " ")
            vsGridDebit.TextMatrix(vsGridDebit.Row, 2) = Trim(gbSearchStr)
            gbSearchID = -1
            gbSearchStr = ""
            vsGridDebit.Rows = vsGridDebit.Rows + 1
        End If
    End Sub
                
                    '----------------------------------------'
                    ' ' To calculate Total amount            '
                    '----------------------------------------'
    Private Sub Calculate()
        Dim mLoop As Long
        Dim mAmtCr As Double
        Dim mAmtDr As Double
        For mLoop = 1 To vsGridCredit.Rows - 1
            If val(vsGridCredit.TextMatrix(mLoop, 1)) <> 0 And Trim(vsGridCredit.TextMatrix(mLoop, 2)) <> "" And val(vsGridCredit.TextMatrix(mLoop, 3)) <> 0 Then
                mAmtCr = mAmtCr + Format(val(vsGridCredit.TextMatrix(mLoop, 3)), "0.00")
            End If
        Next
        For mLoop = 1 To vsGridDebit.Rows - 1
            If val(vsGridDebit.TextMatrix(mLoop, 1)) <> 0 And Trim(vsGridDebit.TextMatrix(mLoop, 2)) <> "" And val(vsGridDebit.TextMatrix(mLoop, 3)) <> 0 Then
                mAmtDr = mAmtDr + Format(val(vsGridDebit.TextMatrix(mLoop, 3)), "0.00")
            End If
        Next
        txtTotalCrAmount.Text = Format(mAmtCr, "0.00")
        txtTotalDrAmount.Text = Format(mAmtDr, "0.00")
    End Sub
'    Private Sub DeleteRows(fg As VSFlexGrid)
'        Dim mLoop As Long
'        mLoop = 1
'        Do While (mLoop < fg.Rows)
'            If fg.IsSelected(mLoop) Then
'                fg.RemoveItem (mLoop)
'            Else
'                mLoop = mLoop + 1
'            End If
'        Loop
'
'    End Sub
    Private Sub FillGrid(vsGrid As VSFlexGrid)
            Dim mSql             As String
            Dim objdb            As New clsDB
            Dim mCnn             As New ADODB.Connection
            Dim Rec              As New ADODB.Recordset
            Dim mRow             As Double
            Dim mWHERE           As String
            Dim mBankDrawnFrom   As String
            Dim mCreditorDebit   As Integer
            
            If vsGrid.Name = vsGridCredit.Name Then
                mCreditorDebit = 0
                mVoucherNo = mCrVoucherNo
            ElseIf vsGrid.Name = vsGridDebit.Name Then
                mCreditorDebit = 1
                mVoucherNo = mDrVoucherNo
            End If
            
            
            If mCreditorDebit = 0 Then
                Call GenerateVoucherNo(0)
                mVoucherNo = mCrVoucherNo
            Else
                Call GenerateVoucherNo(1)
                mVoucherNo = mDrVoucherNo
            End If
                       
            mSql = "Select faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,Cast(Cast(faVoucherChild.fltAmount as numeric(18,2)) as varchar(18)) fltAmount ,faVouchers.tnyStatus From faAccountHeads  "
            mSql = mSql + " LEFT JOIN faVoucherChild ON faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID  "
            mSql = mSql + " LEFT JOIN faVouchers ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
'            mSQL = mSQL + " where faVouchers.intVoucherNo = " & mVoucherNo & " And faAccountHeads.intAccountHeadID <> 887 "
'            mSql = mSql + " where faVouchers.intVoucherNo = " & mVoucherNo & " And faAccountHeads.vchAccountHeadCode <>" & gbAcHeadCodeForCapitalFund
            mSql = mSql + " where faVouchers.intTransactionTypeID =3000  And faVoucherChild.tnyDebitOrCredit  = " & mCreditorDebit
            mSql = mSql + " ORDER BY faAccountHeads.intAccountHeadID "
            objdb.SetConnection mCnn
            Rec.CursorLocation = adUseClient
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            If Not (Rec.EOF And Rec.BOF) Then
                vsGrid.Rows = Rec.RecordCount + 1
'                If Rec!tnyStatus = 1 Then
'                    cmdApprove.Enabled = False
''                    cmdSave.Enabled = False
'                End If
                vsGrid.Col = 0
                vsGrid.Row = 1
                vsGrid.ColSel = 3
                vsGrid.RowSel = vsGrid.Rows - 1
                mSql = Rec.GetString(, , vbTab, Chr(13))
                vsGrid.Clip = mSql
                vsGrid.Row = 1
                vsGrid.Col = 0
            Call Calculate
            Rec.Close
            End If
            vsGrid.Rows = vsGrid.Rows + 1
            End Sub
    Public Sub PressTabKey()
            keybd_event ib_Tab, 0, 0, 0  ' press Tab
            keybd_event ib_Tab, 0, KEYEVENTF_KEYUP, 0  ' release Tab
    End Sub

    Private Sub vsGridDebit_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
'                Call DeleteRows(vsGridDebit)
                vsGridDebit.RemoveItem (vsGridDebit.Row)
            End If
        End If
        Call Calculate   'Added on 19/08/2011
    End Sub

    Private Sub vsGridDebit_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
            PressTabKey
            End If
    End Sub
    
    Public Property Let Mode(mData As Integer)
        intMode = mData
    End Property
    Public Property Get Mode() As Integer
        Mode = intMode
    End Property
        
    Private Sub vsGridDebit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If KeyAscii > 0 And (Col = 1 Or Col = 2) Then KeyAscii = 0
        If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47 And Col = 3) Then KeyAscii = 0
    End Sub
