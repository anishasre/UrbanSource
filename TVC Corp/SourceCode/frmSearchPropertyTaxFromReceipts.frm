VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchPropertyTaxFromReceipts 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Search Property Tax Receipts Isseued"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   10845
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   10500
      Top             =   6600
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid fgReceiptDetails 
      Height          =   4455
      Left            =   30
      TabIndex        =   9
      Top             =   2070
      Width           =   10725
      _cx             =   18918
      _cy             =   7858
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483624
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
      Rows            =   16
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchPropertyTaxFromReceipts.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      BackColor       =   &H00E0E0E0&
      Caption         =   "Receipt Details"
      Height          =   1725
      Left            =   90
      TabIndex        =   6
      Top             =   180
      Width           =   10605
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5640
         TabIndex        =   3
         Top             =   300
         Width           =   3435
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   435
         Left            =   8040
         TabIndex        =   5
         Top             =   1080
         Width           =   1035
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   435
         Left            =   6960
         TabIndex        =   4
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox txtDoorNo2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3540
         TabIndex        =   2
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtDoorNo1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2340
         TabIndex        =   1
         Top             =   720
         Width           =   1155
      End
      Begin VB.TextBox txtWardNo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2340
         TabIndex        =   0
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name"
         Height          =   225
         Left            =   4830
         TabIndex        =   10
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Door No"
         Height          =   225
         Left            =   1410
         TabIndex        =   8
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ward No"
         Height          =   225
         Left            =   1380
         TabIndex        =   7
         Top             =   330
         Width           =   720
      End
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   10530
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSearchPropertyTaxFromReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormInitialize()
    txtWardNo.Text = ""
    txtDoorNo1.Text = ""
    txtDoorNo2.Text = ""
    txtName.Text = ""
    fgReceiptDetails.Clear 1, 1
End Sub
Private Sub SearchReceiptNo()
    Dim mSQL As String
    Dim mSqlFa As String
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim RecReceipt As New ADODB.Recordset
    Dim RecFa   As New ADODB.Recordset
    Dim mYearIDReceipt As Variant
    Dim mYearPeriodReceipt As Variant
    Dim mFlagSWD As Boolean
    Dim mRowCount As Integer
    
    Dim mDoorNo As String
    Dim mDescription As String
    
    Dim mLen1 As Integer
    Dim mString As Variant
    Dim mLenStart As Integer
    Dim mStringCount As Integer
    Dim mStrTemp As String
    Dim mTemp As Long
    Dim mTempString As String
        objDB.SetConnection mCnn
    fgReceiptDetails.Clear 1, 1
    mSqlFa = " Select faVouchers.intTransactionTypeID as TRType, faVouchers.intVoucherNO, faVouchers.dtDate, "
    mSqlFa = mSqlFa + " faVouchers.fltAmount,  faVoucherAddress.intWardNo, faVoucherAddress.intDoorNo, faVoucherAddress.vchDoorNo2, "
    mSqlFa = mSqlFa + " faVoucherChild.intYearID, faVoucherChild.tnyPeriodID,faVoucherAddress.vchName "
    mSqlFa = mSqlFa + " from faVouchers "
    mSqlFa = mSqlFa + " Inner Join faVoucherAddress ON faVouchers.intVoucherID = faVoucherAddress.intVoucherID "
    mSqlFa = mSqlFa + " Inner Join faVoucherChild ON faVouchers.intVoucherID = faVoucherChild.intVoucherID "
    mSqlFa = mSqlFa + " Where (tnyVoucherTypeID = '10') AND (tnyPeriodID < 3) "
    
    If txtWardNo.Text <> "" Then
        mSqlFa = mSqlFa + " And (intWardNo = " & Val(txtWardNo.Text) & " ) "
    End If
    
    If txtDoorNo1.Text <> "" Then
        mSqlFa = mSqlFa + " And (intDoorNo = " & Val(txtDoorNo1.Text) & ")"
    End If
    
    If txtDoorNo2.Text <> "" Then
        mSqlFa = mSqlFa + " And (vchDoorNo2 = '" & txtDoorNo2.Text & "')"
    End If
    
    If txtName.Text <> "" Then
        mSqlFa = mSqlFa + " And (vchName Like '%" & txtName.Text & "%')"
    End If
    
    mSqlFa = mSqlFa + " Group By faVoucherChild.intVoucherID, faVouchers.intTransactionTypeID, "
    mSqlFa = mSqlFa + " faVouchers.intVoucherNO, faVouchers.dtDate, faVouchers.fltAmount,  faVoucherAddress.intWardNo, "
    mSqlFa = mSqlFa + " faVoucherAddress.intDoorNo, faVoucherAddress.vchDoorNo2, "
    mSqlFa = mSqlFa + " faVoucherChild.intYearID , faVoucherChild.tnyPeriodID, faVoucherAddress.vchName "
    mSqlFa = mSqlFa + " Order By faVoucherChild.intYearID desc "
    
    RecFa.Open mSqlFa, mCnn
    mRowCount = 1
    While Not (RecFa.EOF Or RecFa.BOF)
        fgReceiptDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(RecFa!vchName), "", RecFa!vchName)
        fgReceiptDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(RecFa!intWardNo), "", RecFa!intWardNo)
        mDoorNo = IIf(IsNull(RecFa!intDoorNo), "", RecFa!intDoorNo)
        If Not IsNull(RecFa!vchDoorNo2) Then
            mDoorNo = mDoorNo + " - " + RecFa!vchDoorNo2    'IIf(IsNull(RecFa!vchDoorNo2), "", RecFa!vchDoorNo2)
        End If
        fgReceiptDetails.TextMatrix(mRowCount, 3) = mDoorNo
        mDescription = RecFa!intYearID
        If RecFa!tnyPeriodID = 1 Then
            mDescription = mDescription + "1st Half"
        ElseIf RecFa!tnyPeriodID = 2 Then
            mDescription = mDescription + "2st Half"
        End If
        fgReceiptDetails.TextMatrix(mRowCount, 4) = mDescription
        
        fgReceiptDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(RecFa!fltAmount), "", RecFa!fltAmount)
        fgReceiptDetails.TextMatrix(mRowCount, 8) = "Ver(3)"
        mRowCount = mRowCount + 1
        fgReceiptDetails.Rows = fgReceiptDetails.Rows + 1
        RecFa.MoveNext
    Wend
    mCnn.Close
    '-------------------------------------------------------'
    '           Checking in Sahatha Database                '
    '-------------------------------------------------------'
    
        objDB.CreateNewConnection mCnn, enuSourceString.Sahatha
        
'''''    mSQL = "Select *,TblReceiptChild.Amount as [PTaxAmout] from tblReceiptBuildings "
'''''    mSQL = mSQL + " Inner Join tblReceipt "
'''''    mSQL = mSQL + " ON tblReceiptBuildings.ReceiptId = tblReceipt.ID"
'''''    mSQL = mSQL + " Inner Join TblReceiptChild "
'''''    mSQL = mSQL + " ON tblReceiptBuildings.ReceiptId = TblReceiptChild.ReceiptId"
    mSQL = "SELECT      TblReceipt.Id, TblReceiptChild.Amount AS PTaxAmout, "
    mSQL = mSQL + " tblReceiptBuildings.Description, tblReceiptBuildings.Amount, tblReceiptBuildings.HouseNo, "
    mSQL = mSQL + " tblReceiptBuildings.WardNo, tblReceiptBuildings.ReceiptId, TblReceipt.Payee, TblReceipt.ReceiptDate "
    mSQL = mSQL + " FROM         tblReceiptBuildings INNER JOIN "
    mSQL = mSQL + " TblReceipt ON tblReceiptBuildings.ReceiptId = TblReceipt.Id INNER JOIN "
    mSQL = mSQL + " TblReceiptChild ON tblReceiptBuildings.ReceiptId = TblReceiptChild.ReceiptId "
    
    If txtWardNo.Text <> "" Then
        mSQL = mSQL + " Where ( tblReceiptBuildings.wardNo = " & Val(txtWardNo.Text) & " )"
    End If
    
    If txtDoorNo2.Text = "" Then
        mSQL = mSQL + " And  tblReceiptBuildings.houseNo Like '" & txtDoorNo1.Text & "%'"
    Else
        mSQL = mSQL + " And  tblReceiptBuildings.houseNo Like '" & txtDoorNo1.Text & "%" & txtDoorNo2.Text & "%'"
    End If
    
    If txtName.Text <> "" Then
        mSQL = mSQL + " And  TblReceipt.Payee Like '%" & txtName.Text & "%'"
    End If
    
    mSQL = mSQL + " And TblReceipt.CancelFlag <> 1 "
    mSQL = mSQL + " And TblReceiptChild.Period is not null "
    
    mSQL = mSQL + " Group By TblReceipt.Id, TblReceiptChild.Amount , "
    mSQL = mSQL + " tblReceiptBuildings.Description, tblReceiptBuildings.Amount, tblReceiptBuildings.HouseNo, "
    mSQL = mSQL + " tblReceiptBuildings.WardNo, tblReceiptBuildings.ReceiptId, TblReceipt.Payee, TblReceipt.ReceiptDate "
    mSQL = mSQL + " Order By TblReceipt.ReceiptDate Desc "
    
    RecReceipt.Open mSQL, mCnn
    
    fgReceiptDetails.MousePointer = flexHand
    If Not (RecReceipt.EOF Or RecReceipt.BOF) Then
        While Not (RecReceipt.EOF Or RecReceipt.BOF)
           'fgReceiptDetails.WordWrap = True
           fgReceiptDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(RecReceipt!Payee), "", RecReceipt!Payee)
           fgReceiptDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(RecReceipt!WardNo), "", RecReceipt!WardNo)
           fgReceiptDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(RecReceipt!HouseNo), "", RecReceipt!HouseNo)
           fgReceiptDetails.TextMatrix(mRowCount, 2) = ""
           fgReceiptDetails.Cell(flexcpPicture, mRowCount, 4) = imgNote
           fgReceiptDetails.Cell(flexcpPictureAlignment, mRowCount, 4) = flexPicAlignRightTop
           fgReceiptDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(RecReceipt!Description), "", RecReceipt!Description)
           fgReceiptDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(RecReceipt!Amount), "", RecReceipt!Amount)
           fgReceiptDetails.TextMatrix(mRowCount, 7) = IIf(IsNull(RecReceipt!PTaxAmout), "", RecReceipt!PTaxAmout)
           fgReceiptDetails.TextMatrix(mRowCount, 8) = "Ver(1)"
           RecReceipt.MoveNext
           mRowCount = mRowCount + 1
           fgReceiptDetails.Rows = fgReceiptDetails.Rows + 1
        Wend
        fgReceiptDetails.AutoSize 4, , 1, fgReceiptDetails.Rows - 1
    End If
    mCnn.Close
End Sub

Private Sub cmdClear_Click()
    Call FormInitialize
End Sub

Private Sub cmdSearch_Click()
    If Val(txtWardNo) <= 0 And Val(txtDoorNo1) <= 0 Then
        MsgBox "Please specify any search Criteria", vbInformation
        txtWardNo.SetFocus
        Exit Sub
    End If

    Me.MousePointer = vbHourglass
    Call SearchReceiptNo
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = 0
    fgReceiptDetails.TextMatrix(0, 0) = "Name"
    fgReceiptDetails.TextMatrix(0, 1) = "WardNo"
    fgReceiptDetails.TextMatrix(0, 2) = "DNo 1"
    fgReceiptDetails.TextMatrix(0, 3) = "Door No"
    fgReceiptDetails.TextMatrix(0, 4) = "Description"
    fgReceiptDetails.TextMatrix(0, 5) = "Period"
    fgReceiptDetails.TextMatrix(0, 6) = "Tot.Amount"
    fgReceiptDetails.TextMatrix(0, 7) = "PTax Amt"
    fgReceiptDetails.TextMatrix(0, 8) = "SWD Ver."
End Sub

Private Sub Form_Load()
    Me.Height = 7260
    Me.Width = 10965
    Me.Refresh
    WindowsXPC1.InitIDESubClassing
End Sub

