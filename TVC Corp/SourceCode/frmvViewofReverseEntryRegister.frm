VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmViewofReverseEntryRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversed Vouchers"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8966.618
   ScaleMode       =   0  'User
   ScaleWidth      =   28332.82
   Begin VB.Frame Frame1 
      Height          =   6840
      Left            =   12900
      TabIndex        =   12
      Top             =   1440
      Width           =   7365
      Begin VB.TextBox txtInstrumentDate 
         Height          =   360
         Left            =   2700
         TabIndex        =   27
         Top             =   2610
         Width           =   3000
      End
      Begin VB.TextBox txtVoucherNo 
         Height          =   360
         Left            =   2700
         TabIndex        =   24
         Top             =   1335
         Width           =   3000
      End
      Begin VB.TextBox txtInstumentNo 
         Height          =   360
         Left            =   2700
         TabIndex        =   23
         Top             =   1800
         Width           =   3000
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Left            =   2700
         TabIndex        =   22
         Top             =   3075
         Width           =   3000
      End
      Begin VB.TextBox txtFromAmt 
         Height          =   360
         Left            =   2700
         TabIndex        =   21
         Top             =   3540
         Width           =   3000
      End
      Begin VB.ComboBox cmbInstrumentType 
         Height          =   315
         Left            =   2700
         TabIndex        =   20
         Text            =   " "
         Top             =   2220
         Width           =   3000
      End
      Begin VB.TextBox txtToAmt 
         Height          =   315
         Left            =   2700
         TabIndex        =   18
         Top             =   3960
         Width           =   3000
      End
      Begin VB.Label Label12 
         Caption         =   "Instrument Date :"
         Height          =   225
         Left            =   1110
         TabIndex        =   26
         Top             =   2670
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "&Search Options"
         Height          =   225
         Left            =   3000
         TabIndex        =   25
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label11 
         Caption         =   "Instrument Type :"
         Height          =   225
         Left            =   1110
         TabIndex        =   19
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label Label10 
         Caption         =   "To Amount     :"
         Height          =   225
         Left            =   1110
         TabIndex        =   17
         Top             =   3975
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "From Amount  :"
         Height          =   225
         Left            =   1110
         TabIndex        =   16
         Top             =   3570
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Voucher Date :"
         Height          =   225
         Left            =   1110
         TabIndex        =   15
         Top             =   3105
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Insrtument No :"
         Height          =   225
         Left            =   1110
         TabIndex        =   14
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Voucher No    :"
         Height          =   225
         Left            =   1110
         TabIndex        =   13
         Top             =   1395
         Width           =   1335
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   11880
      Top             =   9015
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify"
      Height          =   375
      Left            =   5355
      TabIndex        =   10
      Top             =   8430
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cmbVoucherType 
      Height          =   315
      Left            =   6180
      TabIndex        =   8
      Top             =   1095
      Width           =   2925
   End
   Begin VB.TextBox txtToDate 
      Height          =   330
      Left            =   3210
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtFromDate 
      Height          =   330
      Left            =   1110
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H8000000F&
      Height          =   780
      Left            =   -30
      ScaleHeight     =   720
      ScaleWidth      =   20280
      TabIndex        =   2
      Top             =   60
      Width           =   20340
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   6765
      Left            =   45
      TabIndex        =   1
      Top             =   1515
      Width           =   12900
      _cx             =   22754
      _cy             =   11933
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmvViewofReverseEntryRegister.frx":0000
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3855
      TabIndex        =   0
      Top             =   8430
      Width           =   1365
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   330
      Left            =   8070
      TabIndex        =   11
      Top             =   8430
      Width           =   1590
   End
   Begin VB.Label Label3 
      Caption         =   "Voucher Type"
      Height          =   240
      Left            =   4890
      TabIndex        =   9
      Top             =   1125
      Width           =   1290
   End
   Begin VB.Label lblCount 
      Caption         =   "#"
      Height          =   210
      Left            =   720
      TabIndex        =   7
      Top             =   8460
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   270
      Left            =   2910
      TabIndex        =   4
      Top             =   1140
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   225
      Left            =   540
      TabIndex        =   3
      Top             =   1110
      Width           =   495
   End
End
Attribute VB_Name = "frmViewofReverseEntryRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmbInstrumentType_Change()
        If cmbInstrumentType.ListIndex > -1 Then 'Note:- If any item is selected then
            cmbInstrumentType.Tag = cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex)
        End If
    End Sub
    Private Sub FillInstrumentType()
        Dim mSqlIns As String
        mSqlIns = "SELECT vchInstrumentType,intInstrumentTypeID from faInstrumentTypes"
        PopulateList cmbInstrumentType, mSqlIns, , True, True, True
    End Sub
    Private Sub cmdSearch_Click()
        Call fillGrid
    End Sub

'    Private Sub cmdVerify_Click()
'        frmViewVoucher.FormName = "frmViewofReverseEntryRegister"
'        'frmViewVoucher.FormName = "frmViewPaymentOrder"
'        frmViewVoucher.Show vbModal
'    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FormInitialize
        Call FillInstrumentType
        Call fillVoucherType
    End Sub
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        txtFromDate.Text = DdMmmYy(DateAdd("d", -30, gbTransactionDate))
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        VSGrid.Clear 1, 0
        lblCount.Caption = ""
        gbSearchCode = ""
        gbSearchStr = ""
        gbSearchID = -1
     End Sub
    Private Sub fillGrid()
        Dim mCnn As New ADODB.Connection
        Dim ObjDb As New clsDb
        Dim Rec As New ADODB.Recordset
        Dim msQl As String
        Dim mVoucherNo As Variant
        Dim mVoucherType As String
        Dim mAmount As Variant
        Dim dDate As Variant
        Dim mReason As Variant
        Dim mRowCnt As Integer
        Dim mRecCnt As Integer
        Dim mLoop As Long
        lblCount.Caption = "Rec: 0"
        
        ObjDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        msQl = "SELECT faVouchers.intVoucherNo, faVouchers.numLinkKeyID, faVouchers.fltAmount, faReverseReasons.vchReason, faVouchers.dtDate, "
        msQl = msQl + " faVouchers.tnyVoucherTypeID , faVouchers.tnyStatus, faReverseEntry.dtApprovedDate, faReverseEntry.tnyStatus,faInstrumentTypes.vchInstrumentType, faVouchers.vchInstrumentNo, "
        msQl = msQl + " faVouchers.dtInstrumentDate From faVouchers "
        msQl = msQl + " INNER JOIN  faReverseEntryChild ON faVouchers.numLinkKeyID = faReverseEntryChild.intVoucherID "
        msQl = msQl + " INNER JOIN  faReverseEntry ON faReverseEntryChild.intRequestID = faReverseEntry.intRequestID "
        msQl = msQl + " INNER JOIN  faReverseReasons ON faReverseEntry.intReasonID = faReverseReasons.intReasonID "
        msQl = msQl + " INNER JOIN  faInstrumentTypes ON faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
        msQl = msQl + " WHERE faReverseEntry.tnyStatus = 2 "
        If Trim(txtFromDate.Text) <> "" And Trim(txtToDate.Text) <> "" Then
            msQl = msQl + "AND dtDate Between '" & Trim(txtFromDate.Text) & "' AND '" & Trim(txtToDate.Text) & "'"
        End If
        If val(cmbVoucherType.ListIndex) > 0 Then
            msQl = msQl + " And faVouchers.tnyVoucherTypeID = " & val(cmbVoucherType.ItemData(cmbVoucherType.ListIndex))
        End If
        If Trim(txtVoucherNo.Text) <> "" Then
            msQl = msQl + " AND  faVouchers.intVoucherNo LIKE '%" & Trim(txtVoucherNo.Text) & "%'"
        End If
        If Trim(txtInstumentNo.Text) <> "" Then
            msQl = msQl + "And faVouchers.vchInstrumentNo LIKE '%" & Trim(txtInstumentNo.Text) & "%'"
        End If
        If val(cmbInstrumentType.Tag) > 0 Then
            msQl = msQl + " And faVouchers.intInstrumentTypeID = " & val(cmbInstrumentType.Tag)
        End If
        If Trim(txtInstrumentDate.Text) <> "" Then
            msQl = msQl + " And faVouchers.dtInstrumentDate = '" & Trim(txtInstrumentDate.Text) & " ' "
        End If
        If val(txtFromAmt.Text) <> 0 And val(txtToAmt.Text) <> 0 Then
            msQl = msQl + " And faVouchers.fltAmount BETWEEN " & val(txtFromAmt.Text) & " And " & val(txtToAmt.Text)
        ElseIf val(txtFromAmt.Text) > 0 Then
            msQl = msQl + " And faVouchers.fltAmount > " & val(txtFromAmt.Text) ''& " And " & val(txtToAmt.Text)
        End If
        
        Rec.CursorLocation = adUseClient
        Rec.Open msQl, mCnn
        mRowCnt = 1
        mRecCnt = 1
        VSGrid.Clear 1, 1
        'VSGrid.Rows = Rec.RecordCount + 1
        lblCount.Caption = "Rec:" & str(Rec.RecordCount)
        If Not (Rec.EOF Or Rec.BOF) Then
            While Not (Rec.EOF)
           VSGrid.TextMatrix(mRowCnt, 0) = mRowCnt  'Serial No.
           'VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!VoucherType), "", Rec!VoucherType)
           VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
           VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
           VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
           VSGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
           VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!fltAmount), "", Format(Rec!fltAmount, "0.00"))
           VSGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
           VSGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
           'VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!numLinkKeyID), "", Rec!numLimkKeyID)
           Select Case VSGrid.TextMatrix(mRowCnt, 8)
            Case 10
                mVoucherType = "R" 'Receipt Voucher
            Case 20
                mVoucherType = "P" 'Payment Voucher
            Case 30
                mVoucherType = "C"  'Contra Voucher
            Case 40
                mVoucherType = "J" 'Journal Voucher
             Case Else
                mVoucherType = "Nothing"
        End Select
        VSGrid.TextMatrix(mRowCnt, 1) = mVoucherType
        Rec.MoveNext
                VSGrid.Rows = VSGrid.Rows + 1
                mRowCnt = mRowCnt + 1
                mRecCnt = mRecCnt + 1
           Wend
        End If
        Rec.Close
    End Sub
    Private Sub Form_Resize()
        Label4.Caption = Me.Height & " - " & Me.Width
    End Sub
    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
    End Sub
    Private Sub txtInstrumentDate_LostFocus()
        If Trim(txtInstrumentDate.Text) <> "" Then
            txtInstrumentDate.Text = CheckDateInMMM(txtInstrumentDate.Text)
        End If
    End Sub
    Private Sub txtToDate_Click()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
    End Sub

    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
    End Sub
    Private Sub fillVoucherType()
        cmbVoucherType.AddItem ""
        cmbVoucherType.AddItem "Receipts"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 10
        cmbVoucherType.AddItem "Payments"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 20
        cmbVoucherType.AddItem "Contra"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 30
        cmbVoucherType.AddItem "Journal"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 40
    End Sub
    
    
