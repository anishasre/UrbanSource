VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchPaymentVoucher 
   BackColor       =   &H00D3F7EA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Search"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Tag             =   "20"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtVoucherID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cmbGroupID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VSFlex8LCtl.VSFlexGrid vsFgVoucherList 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   7095
      _cx             =   12515
      _cy             =   3413
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483634
      ForeColor       =   4194368
      BackColorFixed  =   -2147483632
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483634
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
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchPaymentVoucher.frx":0000
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
   Begin VB.TextBox txtTodate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No:"
      Height          =   240
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VoucherType : "
      Height          =   240
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "frmSearchPaymentVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim intVoucherTypeID As Integer
    Private mTransactionGroupId As Integer
    Public Property Let TransactionGroupId(mData As Long)
        mTransactionGroupId = mData
    End Property
    


    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdSearch_Click()
        Call FillFlexGrid
    End Sub

''    Private Sub cmdSearch_Click()
''        mTransactionGroupId = cmbGroupID.ItemData(cmbGroupID.ListIndex)
''        Call FillFlexGrid
''    End Sub
    Private Sub Form_Load()
        ClearAll
        ListGroupIDCombo
        FillFlexGrid
    End Sub
    Private Sub ClearAll()
        txtVoucherID.Text = ""
        txtFromDate.Text = DdMmmYy(gbStartingDate)
        txtTodate.Text = DdMmmYy(gbDate)
    End Sub
    Private Sub ListGroupIDCombo()
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        objDB.SetConnection mCnn
        Dim rs As New ADODB.Recordset
        Dim mSQL As String
        mSQL = "SELECT * from faAccountGroups where intGroupId = 10 or intGroupID = 20"
        rs.Open mSQL, mCnn
        Do Until rs.EOF
            cmbGroupID.AddItem rs(1)
            cmbGroupID.ItemData(cmbGroupID.NewIndex) = rs(0)
            rs.MoveNext
        Loop
        cmbGroupID.ListIndex = 1
    End Sub
    Private Sub FillFlexGrid()
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        objDB.SetConnection mCnn
        Dim Rec As New ADODB.Recordset
        vsFgVoucherList.Rows = 1
        vsFgVoucherList.Rows = 10
        Dim mSQL As String
        'mSQL = "SELECT faVouchers.intVoucherId,faVouchers.intVoucherNo, vchInstrumentNo,fltAmount,dtDate,vchInstrumentType,tnyVoucherTypeID FROM faVouchers LEFT OUTER JOIN faInstrumentTypes ON faVouchers.intInstrumentTypeID=faInStrumentTypes.intInstrumentTypeID Inner Join faTransactions On faTransactions.intVoucherId=faVouchers.intVoucherId WHERE dtDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & "' AND convert(varchar(10),intVoucherNo) like '" & txtVoucherID.Text & "%' AND faTransactions.intGroupId=" & cmbGroupID.ItemData(cmbGroupID.ListIndex) & " and faTransactions.intGroupId=" & mTransactionGroupId & "  Order by dtDate"
        mSQL = "SELECT "
        mSQL = mSQL + "faVouchers.intVoucherID , faVouchers.intVoucherNo, vchInstrumentNo, faVouchers.fltAmount, dtDate, vchInstrumentType, tnyVoucherTypeID "
        mSQL = mSQL + "From faVouchers LEFT OUTER JOIN faInstrumentTypes "
        mSQL = mSQL + " ON faVouchers.intInstrumentTypeID=faInStrumentTypes.intInstrumentTypeID "
        mSQL = mSQL + " Inner Join faTransactions "
        mSQL = mSQL + " On faTransactions.intVoucherId=faVouchers.intVoucherId "
        mSQL = mSQL + " Inner Join faTransactionChild "
        mSQL = mSQL + " On faTransactions.intTransactionId=faTransactionChild.intTransactionId "
        mSQL = mSQL + " Inner Join faAccountHeads "
        mSQL = mSQL + " On faTransactionChild.intAccountHeadId=faAccountHeads.intAccountHeadId "
        mSQL = mSQL + " Where dtDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtTodate.Text & "' "
        If cmbGroupID.ListIndex = 0 Then
            mSQL = mSQL + " And faVouchers.intInstrumentTypeID = " & 1
        Else
            mSQL = mSQL + " And faVouchers.intInstrumentTypeID = " & 5
        End If
        mSQL = mSQL + " AND faVouchers.intVoucherNo like '" & txtVoucherID.Text & "%' "
        mSQL = mSQL + " AND  faTransactions.intGroupId=" & mTransactionGroupId & ""
        mSQL = mSQL + " And intSerialNo =1 Order By dtDate Desc "
        
        Rec.Open mSQL, mCnn
        Dim i As Integer
        i = 0
        While Rec.EOF = False
            i = i + 1
            vsFgVoucherList.Rows = vsFgVoucherList.Rows + 1
            vsFgVoucherList.TextMatrix(i, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            vsFgVoucherList.TextMatrix(i, 1) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            vsFgVoucherList.TextMatrix(i, 2) = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
            vsFgVoucherList.TextMatrix(i, 3) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            vsFgVoucherList.TextMatrix(i, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            vsFgVoucherList.TextMatrix(i, 5) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            Rec.MoveNext
        Wend
        Rec.Close
    End Sub
    Private Sub txtFromDate_LostFocus()
        If Trim(txtFromDate.Text) = "" Then
            ClearAll
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub
    Private Sub txtTodate_LostFocus()
        If Trim(txtTodate.Text) = "" Then
            ClearAll
        Else
            txtTodate.Text = CheckDateInMMM(txtTodate.Text)
        End If
    End Sub
    Private Sub vsFgVoucherList_DblClick()
        gbSearchID = val(vsFgVoucherList.TextMatrix(vsFgVoucherList.Row, 5))
        gbSearchStr = vsFgVoucherList.TextMatrix(vsFgVoucherList.Row, 0)
        Unload Me
    End Sub
    
    
