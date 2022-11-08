VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchJournalVouchers 
   BackColor       =   &H00F4FBFA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Journal Vouchers"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
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
   ScaleHeight     =   6060
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D5F0EE&
      Caption         =   "Cance&L"
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
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1215
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00D5F0EE&
      Caption         =   "&Search"
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
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "20"
      Top             =   1215
      Width           =   975
   End
   Begin VB.TextBox txtVoucherID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7305
      TabIndex        =   8
      Top             =   255
      Width           =   1935
   End
   Begin VB.ComboBox cmbAccountHeadID 
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
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   675
      Width           =   7485
   End
   Begin VSFlex8LCtl.VSFlexGrid vsFgVoucherList 
      Height          =   3915
      Left            =   15
      TabIndex        =   4
      Top             =   1665
      Width           =   9735
      _cx             =   17171
      _cy             =   6906
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
      Rows            =   20
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchJournalVouchers.frx":0000
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
      Left            =   4140
      TabIndex        =   3
      Top             =   240
      Width           =   1395
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   240
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   1095
      Left            =   45
      Top             =   60
      Width           =   9675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   7
      Top             =   255
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   5
      Top             =   705
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3270
      TabIndex        =   1
      Top             =   255
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   510
      TabIndex        =   0
      Top             =   255
      Width           =   1140
   End
End
Attribute VB_Name = "frmSearchJournalVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Public mTransactionGroupId As Integer

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub
    Private Sub cmdSearch_Click()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim i As Integer
        Dim mStartID    As Variant
        Dim mEndID      As Variant
        Dim mStartDate  As String
        Dim mEndDate    As String
        
        objDb.SetConnection mCnn
        vsFgVoucherList.Rows = 1
        'vsFgVoucherList.Rows = 10
        
        mStartID = -1
        mEndID = 100000
        mStartDate = "01-Apr-" & gbFinancialYearID
        mEndDate = "31-Mar-" & gbFinancialYearID + 1
        
        mSql = "Select MIN(intVoucherID) As StartID, MAX(intVoucherID) As EndID From faVouchers"
        mSql = mSql + " Where tnyVoucherTypeID = 40"
        mSql = mSql + " And dtDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & "' "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStartID = IIf(IsNull(Rec!StartID), -1, Rec!StartID)
            mEndID = IIf(IsNull(Rec!EndID), -1, Rec!EndID)
        End If
        Rec.Close
        
        mSql = "SELECT "
        mSql = mSql + " faVouchers.intVoucherID , faVouchers.intVoucherNo, vchInstrumentNo, faVouchers.fltAmount, "
        mSql = mSql + " dtDate, vchInstrumentType, tnyVoucherTypeID, vchAccountHead, tinDebitOrCreditFlag "
        mSql = mSql + " From faVouchers LEFT OUTER JOIN faInstrumentTypes "
        mSql = mSql + " ON faVouchers.intInstrumentTypeID=faInStrumentTypes.intInstrumentTypeID "
        mSql = mSql + " Inner Join faTransactions "
        mSql = mSql + " On faTransactions.intVoucherId=faVouchers.intVoucherId "
        mSql = mSql + " Inner Join faTransactionChild "
        mSql = mSql + " On faTransactions.intTransactionId=faTransactionChild.intTransactionId "
        mSql = mSql + " Inner Join faAccountHeads "
        mSql = mSql + " On faTransactionChild.intAccountHeadId=faAccountHeads.intAccountHeadId "
        mSql = mSql + " Where dtDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & "' AND faVouchers.intVoucherNo like '" & txtVoucherID.Text & "%' AND  faTransactions.intGroupId = 40 AND  faVouchers.intTransactionTypeID <> 3000"
        mSql = mSql + " And intSerialNo =1 Order By dtDate Desc "
        
        
        '================================================================================"
        ' Changed by Aiby on 04-Jun-2010
        '================================================================================"
        mSql = ""
        mSql = mSql + " SELECT faVoucherCHild.*, faVouchers.intVoucherID, faVouchers.intVoucherNo, vchInstrumentNo, faVouchers.fltAmount, dtDate, vchInstrumentType, tnyVoucherTypeID, vchAccountHead, tnyDebitOrCredit"
        mSql = mSql + " From faVouchers"
        mSql = mSql + " LEFT JOIN faTransactions ON faTransactions.intVoucherId=faVouchers.intVoucherId"
        mSql = mSql + " LEFT JOIN faInstrumentTypes ON faVouchers.intInstrumentTypeID=faInStrumentTypes.intInstrumentTypeID"
        mSql = mSql + " LEFT JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faVouchers.intKeyID1"
        mSql = mSql + " LEFT JOIN faVoucherCHild ON faVoucherChild.intVoucherID = faVouchers.intVoucherID"
        'mSql = mSql + " Where dtDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & "' "
        mSql = mSql + " Where faVouchers.tnyVoucherTypeID = 40 AND intSLNO = 1 AND  faVouchers.intTransactionTypeID <> 3000"
        mSql = mSql + " And faVouchers.intVoucherID Between " & mStartID & " And " & mEndID
        If txtVoucherID.Text = "" Then
            mSql = mSql + " AND faVouchers.intVoucherNo like '%'"
        Else
            mSql = mSql + " AND faVouchers.intVoucherNo =" & val(txtVoucherID.Text)
        End If
        If cmbAccountHeadID.ListIndex > 0 Then
            mSql = mSql + " And faAccountHeads.intAccountHeadID = " & cmbAccountHeadID.ItemData(cmbAccountHeadID.ListIndex)
        End If
        'mSql = mSql + " AND faVouchers.tnyVoucherTypeID = 40 AND intSLNO = 1"
        mSql = mSql + " Order By faVouchers.intVoucherNo Desc"
        Rec.Open mSql, mCnn
        
        i = 0
        While Rec.EOF = False
            i = i + 1
            vsFgVoucherList.Rows = vsFgVoucherList.Rows + 1
            vsFgVoucherList.TextMatrix(i, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            vsFgVoucherList.TextMatrix(i, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            vsFgVoucherList.TextMatrix(i, 2) = IIf(Rec!tnyDebitOrCredit = 0, "Dr", "Cr")
            If Not (IsNull(Rec!dtDate)) Then
                vsFgVoucherList.TextMatrix(i, 3) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                If (CDate(Rec!dtDate) < CDate(txtFromDate.Text) Or CDate(Rec!dtDate) > CDate(txtToDate.Text)) Then
                    vsFgVoucherList.Cell(flexcpBackColor, i, 0, , 4) = &H80C0FF
                End If
                If (CDate(Rec!dtDate) < mStartDate Or CDate(Rec!dtDate) > mEndDate) Then
                    vsFgVoucherList.TextMatrix(i, 6) = 1                               'Transaction not in Current Financial Year
                Else
                    vsFgVoucherList.TextMatrix(i, 6) = 0                               'Transaction in Current Financial Year
                End If
            End If
            vsFgVoucherList.TextMatrix(i, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            vsFgVoucherList.TextMatrix(i, 5) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            Rec.MoveNext
        Wend
        Rec.Close
    End Sub
    Private Sub Form_Load()
        ClearAll
        ListGroupIDCombo
        FillFlexGrid
    End Sub
    Private Sub ClearAll()
        txtVoucherID.Text = ""
        txtFromDate.Text = DdMmmYy(gbStartingDate)
        txtToDate.Text = DdMmmYy(gbDate)
    End Sub
    Private Sub ListGroupIDCombo()
        Dim mSql As String
        mSql = " Select Distinct vchAccountHeadCode + '  ' + vchAccountHead,  faAccountHeads.intAccountHeadID From faTransactionChild Inner Join"
        mSql = mSql + " faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
        mSql = mSql + " faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID"
        mSql = mSql + " Where intSerialNo = 1 And faTransactions.intGroupID = 40"
        'mSQL = mSQL + " Order By vchAccountHead"
        Call PopulateList(cmbAccountHeadID, mSql, , True, True, True)
        
    End Sub
    Private Sub FillFlexGrid()
       
    End Sub
    Private Sub txtFromDate_LostFocus()
        If Trim(txtFromDate.Text) = "" Then
            ClearAll
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
        If CDate(txtFromDate.Text) Then
            If CDate(txtToDate.Text) Then
                If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtFromDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub
    Private Sub txtTodate_LostFocus()
        If Trim(txtToDate.Text) = "" Then
            ClearAll
        Else
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        End If
        If CDate(txtToDate.Text) Then
            If CDate(txtFromDate.Text) Then
                If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtToDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtToDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub

    Private Sub txtVoucherID_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub


    Private Sub vsFgVoucherList_DblClick()
        If val(vsFgVoucherList.TextMatrix(vsFgVoucherList.Row, 6)) <> 1 Then 'Transaction not in Current Financial Year
            gbSearchID = val(vsFgVoucherList.TextMatrix(vsFgVoucherList.Row, 5))
            gbSearchStr = vsFgVoucherList.TextMatrix(vsFgVoucherList.Row, 0)
            Unload Me
        End If
    End Sub
   
