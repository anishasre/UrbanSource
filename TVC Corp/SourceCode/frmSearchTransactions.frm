VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchTransactions 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Transactions"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbTransactionTypes 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1740
      Width           =   7425
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   7515
      TabIndex        =   20
      Top             =   1080
      Width           =   1740
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1860
      TabIndex        =   19
      Top             =   1080
      Width           =   4320
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   300
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1395
      Width           =   300
   End
   Begin VB.Frame fmeGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Left            =   9630
      TabIndex        =   12
      Top             =   735
      Width           =   1170
      Begin VB.CheckBox chkGroup 
         BackColor       =   &H80000016&
         Caption         =   "Journal"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   16
         Top             =   1035
         Width           =   1185
      End
      Begin VB.CheckBox chkGroup 
         BackColor       =   &H80000016&
         Caption         =   "Contra"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   750
         Width           =   1185
      End
      Begin VB.CheckBox chkGroup 
         BackColor       =   &H80000016&
         Caption         =   "Payment"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   450
         Width           =   1185
      End
      Begin VB.CheckBox chkGroup 
         BackColor       =   &H80000016&
         Caption         =   "Receipt"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   165
         Width           =   1185
      End
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1860
      TabIndex        =   6
      Top             =   780
      Width           =   1740
   End
   Begin VB.TextBox txtTodate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   780
      Width           =   1740
   End
   Begin VB.ComboBox cmbAccountHeads 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1395
      Width           =   7425
   End
   Begin VB.TextBox txtVoucherID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   7515
      TabIndex        =   2
      Top             =   780
      Width           =   1740
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "20"
      Top             =   2565
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2565
      Width           =   975
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3090
      Left            =   0
      TabIndex        =   4
      Top             =   2925
      Width           =   10905
      _cx             =   19235
      _cy             =   5450
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
      ForeColor       =   64
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   12582912
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
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
      Rows            =   20
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchTransactions.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Types"
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
      Height          =   270
      Left            =   225
      TabIndex        =   23
      Top             =   1770
      Width           =   1605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   270
      Left            =   6795
      TabIndex        =   21
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   270
      Left            =   1305
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "         Transactions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   180
      Width           =   11805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Height          =   270
      Left            =   900
      TabIndex        =   10
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Height          =   270
      Left            =   3735
      TabIndex        =   9
      Top             =   780
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head"
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
      Height          =   270
      Left            =   630
      TabIndex        =   8
      Top             =   1425
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No"
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
      Height          =   270
      Left            =   6495
      TabIndex        =   7
      Top             =   795
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00008000&
      Height          =   1950
      Left            =   90
      Top             =   570
      Width           =   10785
   End
End
Attribute VB_Name = "frmSearchTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVarReceipt     As Boolean
Private mVarPayment     As Boolean
Private mVarContra      As Boolean
Private mVarJournal     As Boolean
Dim mvarCategoryID      As Integer
Dim mPreviousYearMode   As Integer

    Public Property Let Receipt(mVal As Boolean)
        mVarReceipt = mVal
    End Property
    
    Public Property Let Payment(mVal As Boolean)
        mVarPayment = mVal
    End Property
    
    Public Property Let Contra(mVal As Boolean)
        mVarContra = mVal
    End Property
    
    Public Property Let Journal(mVal As Boolean)
        mVarJournal = mVal
    End Property
    Public Property Let FormSelectionType(mData As Integer)
        mvarCategoryID = mData 'Reverse Request
    End Property

    Private Sub cmdBrowse_Click()
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Order By vchAccountHeadCode"
        frmSearchAccountHeads.Show vbModal
        If gbSearchStr <> "" Then
            cmbAccountHeads.Text = gbSearchStr
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdsearch_Click()
        Dim mSql As String
        Dim cSQL As String
        Dim mType As String
        Dim mRow As Double
        Dim mIndex As Double
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connection Not Present"
            Exit Sub
        End If
        mType = ""
        For mIndex = 0 To chkGroup.count - 1
            If chkGroup(mIndex).Value = False Then
                mType = mType + "0," + CStr(mIndex + 1)
            End If
        Next mIndex
        mType = mType + "0"
        mIndex = 0
        cSQL = ""
        cSQL = "Select Distinct faTransactions.intTransactionID " & _
                "From faTransactions " & _
                "Left Join faTransactionType On faTransactions.intTransactionTypeID = faTransactionType.intTransactionTypeID " & _
                "Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                "Inner Join faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID " & _
                "Inner Join faVouchers On faVouchers.intVoucherID = faTransactions.intVoucherID " & _
                "Left Join faVoucherAddress On faVoucherAddress.intVoucherID  = faVouchers.intVoucherID " & _
                "Where faTransactions.intGroupID not in(0" & mType & ") And IsNull(faTransactions.tnyStatus,100) <> 4 And dtDate Between '" & Trim(txtFromDate.Text) & "' And '" & Trim(txtToDate.Text) & "'"
        If Trim(txtVoucherID.Text) <> "" Then
            cSQL = cSQL + " And faVouchers.intVoucherNo = " & Trim(txtVoucherID.Text)
        End If
        If Trim(txtName.Text) <> "" Then
            cSQL = cSQL + " And faVoucherAddress.vchName Like '%" & txtName.Text & "%'"
        End If
        If val(Trim(txtAmount.Text)) > 0 Then
            cSQL = cSQL + " And Convert(varchar(25),faTransactionChild.fltAmount) Like '" & val(Trim(txtAmount.Text)) & "%'"
        End If
        If cmbTransactionTypes.ListIndex > 0 Then
            cSQL = cSQL + " And IsNull(faTransactions.intTransactionTypeID,0) = " & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex)
        End If
        If cmbAccountHeads.ListIndex > 0 Then
            cSQL = cSQL + " And faTransactionChild.intAccountHeadID = " & cmbAccountHeads.ItemData(cmbAccountHeads.ListIndex)
        End If
        mSql = "Select  faTransactions.intTransactionID,faVouchers.intVoucherID,dtTransactionDate,faTransactions.vchGroup,faTransactionChild.intSerialNo,faTransactionChild.fltAmount," & _
                "Case When tinDebitOrCreditFlag = 1 Then 'Dr' Else 'Cr' End [DrOrCr],faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode," & _
                "faAccountHeads.vchAccountHead,faVouchers.intVoucherNo,intInstrumentTypeID,faVoucherAddress.vchName,faVoucherAddress.intWardNo,faVoucherAddress.intDoorNo,faVoucherAddress.vchDoorNo2," & _
                "faTransactionType.intTransactionTypeID , vchTransactionType,faVouchers.fltAdvAmtAdj " & _
                "From faTransactions " & _
                "Inner Join faTransactionType On faTransactions.intTransactionTypeID = faTransactionType.intTransactionTypeID " & _
                "Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                "Inner Join faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID " & _
                "Inner Join faVouchers On faVouchers.intVoucherID = faTransactions.intVoucherID " & _
                " Inner Join faVoucherAddress On faVoucherAddress.intVoucherID  = faVouchers.intVoucherID " & _
                "Where IsNull(faTransactions.tnyStatus,100) <> 4 And dtDate Between '" & Trim(txtFromDate.Text) & "' And '" & Trim(txtToDate.Text) & "' And faTransactions.intTransactionID in(" & cSQL & ")"
        

        mSql = mSql + " AND ISNULL(faVouchers.tnyVoucherGroupID,0)<>4"    'TO BLOCK INTERRUPTED RECEIPTS
        mSql = mSql + " Order By dtTransactionDate,faVouchers.intVoucherID Desc"
        Rec.Open mSql, mCnn
        mRow = 1
        vsGrid.MergeCells = flexMergeFree
        vsGrid.MergeCol(0) = True
        vsGrid.Rows = 1
        vsGrid.Rows = 50
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                If mIndex <> Rec!intTransactionID Then
                    mIndex = Rec!intTransactionID
                Else
                    vsGrid.Cell(flexcpForeColor, mRow, 0) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 3) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 5) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 7) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 8) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 9) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 10) = vbWhite
                    vsGrid.Cell(flexcpForeColor, mRow, 11) = vbWhite
                End If
                vsGrid.TextMatrix(mRow, 0) = Rec!intVoucherNo
                vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!dtTransactionDate), "", Rec!dtTransactionDate)
                vsGrid.TextMatrix(mRow, 5) = Rec!intVoucherID
                vsGrid.TextMatrix(mRow, 7) = IIf(IsNull(Rec!fltAdvAmtAdj), "", Rec!fltAdvAmtAdj)
                vsGrid.TextMatrix(mRow, 8) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                vsGrid.TextMatrix(mRow, 9) = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                vsGrid.TextMatrix(mRow, 10) = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                vsGrid.TextMatrix(mRow, 11) = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                vsGrid.TextMatrix(mRow, 1) = Rec!vchAccountHeadCode + "   " + Rec!vchAccountHead
                vsGrid.TextMatrix(mRow, 2) = Rec!DrOrCr
                vsGrid.TextMatrix(mRow, 4) = Format(Rec!fltAmount, "0.00")
                Rec.MoveNext
                mRow = mRow + 1
                vsGrid.Rows = vsGrid.Rows + 1
            Wend
        End If
        Rec.Close
    End Sub

    Private Sub Form_Load()
        Call MakeChecked
        txtFromDate.Text = DdMmmYy(DateAdd("M", -1, gbTransactionDate)) ' DdMmmYy(gbStartingDate)
        txtToDate.Text = DdMmmYy(gbDate)
        PopulateList cmbAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads", , True, , True
        PopulateList cmbTransactionTypes, "Select vchTransactionType,intTransactionTypeID From faTransactionType Order By vchTransactionType", , True, , True
    End Sub
   
    Private Sub MakeChecked()
        If mVarReceipt Then
            chkGroup(0).Value = 1
        Else
            chkGroup(0).Value = 0
        End If
        If mVarPayment Then
            chkGroup(1).Value = 1
        Else
            chkGroup(1).Value = 0
        End If
        If mVarContra Then
            chkGroup(2).Value = 1
        Else
            chkGroup(2).Value = 0
        End If
        If mVarJournal Then
            chkGroup(3).Value = 1
        Else
            chkGroup(3).Value = 0
        End If
    End Sub
    
    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate)
        'txtFromDate.Text = DdMmmYy(txtFromDate.Text)
    End Sub
    
    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        'txtToDate.Text = DdMmmYy(txtToDate.Text)
    End Sub
    
    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 1 Then
            If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
            gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 5)
            gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 0)
            gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 4) ''Now uses as Amount
            Unload Me
        End If
    End Sub
 
    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property
