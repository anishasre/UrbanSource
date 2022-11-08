VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmLinkAllotmentsWithReceipts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " LINK ALLOTMENTS WITH RECEIPTS"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13350
   ForeColor       =   &H00000000&
   Icon            =   "frmLinkAllotmentsWithReceipts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   12735
      Top             =   7290
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -45
      ScaleHeight     =   285
      ScaleWidth      =   13380
      TabIndex        =   2
      Top             =   0
      Width           =   13380
   End
   Begin VB.Frame frmSearch 
      Height          =   735
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   13290
      Begin VB.ComboBox cmbCategory 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9225
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   2175
      End
      Begin VB.ComboBox cmbSourceOfFund 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   3480
      End
      Begin VB.ComboBox cmbYear 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   8
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Source Of Fund"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3015
         TabIndex        =   5
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label lblYear 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   255
         Width           =   465
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5055
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   13290
      _cx             =   23442
      _cy             =   8916
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
      Rows            =   1
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLinkAllotmentsWithReceipts.frx":1CCA
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
   Begin VB.Frame frmVouchers 
      Height          =   690
      Left            =   45
      TabIndex        =   9
      Top             =   6120
      Width           =   13290
      Begin VB.TextBox txtDemandNo 
         Height          =   285
         Left            =   13080
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10980
         TabIndex        =   17
         Top             =   180
         Width           =   1050
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8415
         TabIndex        =   13
         Top             =   180
         Width           =   2265
      End
      Begin VB.TextBox txtVoucherNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   12
         Top             =   180
         Width           =   2940
      End
      Begin VB.CommandButton cmdSearchVoucher 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   4230
         TabIndex        =   11
         Top             =   180
         Width           =   285
      End
      Begin VB.TextBox txtVrDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         TabIndex        =   10
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label Label7 
         Caption         =   "AMOUNT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7695
         TabIndex        =   16
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "VOUCHER NO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   15
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label8 
         Caption         =   "DATE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5085
         TabIndex        =   14
         Top             =   180
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmLinkAllotmentsWithReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mYearID As Integer
Private Sub FillCombo()
    Dim mSql As String
        
    If gbLBPanchayat = 1 Then
        mSql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(10,11,12,13,14)"
    Else
        mSql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(10,11,12,13,14)"
    End If
    PopulateList cmbSourceOfFund, mSql, True, True, True, True, enuSourceString.Saankhya
    
    mSql = "SELECT vchTransactionCategory,intCategoryID FROM faTransactionCategory"
    PopulateList cmbCategory, mSql, True, True, True, True
    
    PopulateList cmbYear, "Select Cast(intFinancialYearID as varchar(4)) + '-' + Right(Cast(intFinancialYearID+1 as varchar(4)),2),intFinancialYearID  From faFinancialYear WHERE intFinancialYearID > 2011", , , , True
    cmbYear.ListIndex = cmbYear.ListCount - 1
    cmbYear.Enabled = False
    vsGrid.SelectionMode = flexSelectionByRow
        
End Sub

Private Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCnt As Integer
    Dim mLoop   As Integer
    
    On Error GoTo err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
''''    mSQL = " SELECT faAllotmentLetters.intAllotmentID, faAllotmentLetters.vchAllotmentNo, faAllotmentLetters.dtAllotmentDate, faAllotmentLetters.intSourceOfFundID, "
''''    mSQL = mSQL + "   faAllotmentLetters.intCategoryID, faAllotmentLetters.intGrossAccountHeadID, faAllotmentLetters.fltAmount, faAllotmentLetters.intFinancialYearID,"
''''    mSQL = mSQL + "   faAllotmentLetters.tnyCancelledFlag,faAllotmentLetters.tnyStatus, suSourceOfFund.vchSourceFundName, faTransactionCategory.vchTransactionCategory,"
''''    mSQL = mSQL + "   faVouchers.intVoucherID,faVouchers.intVoucherNo,faVouchers.dtDate,"
''''    mSQL = mSQL + "   faIDemandTBL.numDEmandID,faIDemandTBL.vchDemandNo , faIDemandTBL.dtDemandDate"
''''    mSQL = mSQL + "   FROM faAllotmentLetters"
''''    mSQL = mSQL + "   LEFT JOIN faVouchers ON faVouchers.vchInstrumentNo=faAllotmentLetters.vchAllotmentNO AND faVouchers.intInstrumentTypeID=6 AND tnyVoucherTypeID=10 "
''''    mSQL = mSQL + "   LEFT JOIN faIDemandTBL ON faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
''''    mSQL = mSQL + "   INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faAllotmentLetters.intSourceOfFundID"
''''    mSQL = mSQL + "   INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faAllotmentLetters.intCategoryID"
''''    mSQL = mSQL + "   WHERE faAllotmentLetters.intFinancialYearID =  " & cmbYear.ItemData(cmbYear.ListIndex) & " "
''''    mSQL = mSQL + "   AND faAllotmentLetters.intSourceOfFundID In(10,11,12,13,14)"
    
''''
''''    mSQL = "SELECT faAllotmentLetters.intAllotmentID, faAllotmentLetters.vchAllotmentNo, faAllotmentLetters.dtAllotmentDate, faAllotmentLetters.intSourceOfFundID,"
''''    mSQL = mSQL + "   faAllotmentLetters.intCategoryID, faAllotmentLetters.intGrossAccountHeadID, faAllotmentLetters.fltAmount, faAllotmentLetters.intFinancialYearID,"
''''    mSQL = mSQL + "   faAllotmentLetters.tnyCancelledFlag,faAllotmentLetters.tnyStatus, suSourceOfFund.vchSourceFundName, faTransactionCategory.vchTransactionCategory,"
''''    mSQL = mSQL + "   faVouchers.intVoucherID,faVouchers.intVoucherNo,faVouchers.dtDate,"
''''    mSQL = mSQL + "   faIDemandTBL.numDemandID , faIDemandTBL.vchDemandNo, faIDemandTBL.dtDemandDate"
''''    mSQL = mSQL + "   From faAllotmentLetters"
''''    mSQL = mSQL + "   LEFT JOIN faIDemandTBL ON faIDemandTBL.numSubLedgerID=faAllotmentLetters.intAllotmentID AND  faIDemandTBL.vchInstrumentNo=faAllotmentLetters.vchAllotmentNO"
''''    mSQL = mSQL + "   LEFT JOIN faVouchers ON faVouchers.intVoucherID=faIDemandTBL.intVoucherID"
''''    mSQL = mSQL + "   INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faAllotmentLetters.intSourceOfFundID"
''''    mSQL = mSQL + "   INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faAllotmentLetters.intCategoryID"
''''    mSQL = mSQL + "   Where faAllotmentLetters.intFinancialYearID = 2015"
''''    mSQL = mSQL + "   AND faAllotmentLetters.intSourceOfFundID In(10,11,12,13,14)"

        
    mSql = " SELECT  faAllotmentLetters.intAllotmentID, faAllotmentLetters.vchAllotmentNo, faAllotmentLetters.dtAllotmentDate, faAllotmentLetters.intSourceOfFundID,"
    mSql = mSql + "   faAllotmentLetters.intCategoryID, faAllotmentLetters.intGrossAccountHeadID, faAllotmentLetters.fltAmount, faAllotmentLetters.intFinancialYearID,"
    mSql = mSql + "   faAllotmentLetters.tnyCancelledFlag , faAllotmentLetters.tnyStatus, suSourceOfFund.vchSourceFundName, faTransactionCategory.vchTransactionCategory"
    mSql = mSql + "   From faAllotmentLetters"
    mSql = mSql + "   LEFT JOIN faIDemandTBL ON faIDemandTBL.numSubLedgerID=faAllotmentLetters.intAllotmentID"
    mSql = mSql + "   INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faAllotmentLetters.intSourceOfFundID"
    mSql = mSql + "   INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faAllotmentLetters.intCategoryID"
    mSql = mSql + "   Where faAllotmentLetters.intFinancialYearID = 2015 And IsNull(faAllotmentLetters.tnyStatus, 0) = 1"
    mSql = mSql + "   AND ISNULL(faIDemandTBL.intVoucherID,0)=0 AND vchAllotmentNo IS NOT NULL"
    mSql = mSql + "   AND ISNULL(faAllotmentLetters.tnyStatus,0)<>8"
    mSql = mSql + "   AND intSourceofFundID NOT IN (4,3)"
    
    If cmbSourceOfFund.ListIndex > 0 Then
        mSql = mSql + " AND faAllotmentLetters.intSourceOfFundID= " & cmbSourceOfFund.ItemData(cmbSourceOfFund.ListIndex) & " "
    End If
    
    If cmbCategory.ListIndex > 0 Then
        mSql = mSql + " AND faAllotmentLetters.intCategoryID= " & cmbCategory.ItemData(cmbCategory.ListIndex) & " "
    End If
    
    mSql = mSql + " ORDER BY faAllotmentLetters.dtAllotmentDate"
    
    Rec.CursorLocation = adUseClient
    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
        vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate))
        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
        vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        'vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
        'vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        'vsGrid.TextMatrix(mRowCnt, 7) = Format(Rec!dtDate, "dd-mmm-yyyy")
        
          
        If Rec!tnyStatus = 8 Then
            For mLoop = 0 To vsGrid.Cols - 1
                vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HC0E0FF
            Next mLoop
            vsGrid.TextMatrix(mRowCnt, 8) = "CANCELLED"
        ElseIf Rec!tnyStatus = 1 Then
            vsGrid.TextMatrix(mRowCnt, 8) = "Letter Of Authority APPROVED"
'            If IsNull(Rec!intVoucherNo) Then
'                vsGrid.TextMatrix(mRowCnt, 8) = "NO  RECEIPT GENERATED"
'                If IsNull(Rec!vchDemandNo) Then
'                    vsGrid.TextMatrix(mRowCnt, 8) = "NO  DEMANDGENERATED"
'                End If
'            End If
        ElseIf Rec!tnyStatus = 0 Then
             vsGrid.TextMatrix(mRowCnt, 8) = "Letter Of Authority NOT APPROVED"
        End If
        
        vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
        'vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
        vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intSourceOfFundID), "", Rec!intSourceOfFundID)
        vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!intCategoryID), "", Rec!intCategoryID)
        vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
        'vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
        vsGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!intGrossAccountHeadID), "", Rec!intGrossAccountHeadID)
        vsGrid.TextMatrix(mRowCnt, 16) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
        
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    Rec.Close
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmbCategory_Click()
    Call FillGrid
End Sub

Private Sub cmbSourceOfFund_Click()
    Call FillGrid
End Sub

Private Sub cmbYear_Click()
    Call FillGrid
End Sub

Private Sub cmdAdd_Click()
    Dim mSql    As String
    Dim objDB   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    
    If val(txtVoucherNo.Tag) = 0 Then
        MsgBox "Voucher number cannot be Balnk", vbInformation
        txtVoucherNo.SetFocus
        Exit Sub
    End If
'''    mSql = "UPDATE faVouchers SET  intInstrumentTypeID=6,vchInstrumentNo='" & Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) & "'  "
'''    'mSql = mSql + " ,intKeyID2=" & vsGrid.TextMatrix(vsGrid.Row, 14) & "  "
'''    mSql = mSql + " Where intVoucherID = " & val(txtVoucherNo.Tag) & ""
'''    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
   
'''    mSql = "UPDATE faIDemandTBL SET  tnyStatus=1,intInstrumentTypeID=6,vchInstrumentNo='" & Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) & "' "
'''    mSql = mSql + "  ,intVoucherID=" & val(txtVoucherNo.Tag) & " ,dtVoucherDate='" & (txtVrDate.Text) & " ' "
'''    mSql = mSql + "  WHERE numSubLedgerID=" & val(vsGrid.TextMatrix(vsGrid.Row, 9)) & ""
'''    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
   
    mSql = "UPDATE faIDemandTBL SET  tnyStatus=1"
    'intInstrumentTypeID=6,vchInstrumentNo='" & Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) & "' "
    mSql = mSql + "  ,intVoucherID=" & val(txtVoucherNo.Tag) & " ,dtVoucherDate='" & (txtVrDate.Text) & " ' "
    mSql = mSql + "  ,numSubLedgerID=" & val(vsGrid.TextMatrix(vsGrid.Row, 9)) & ""
    mSql = mSql + "  WHERE numDemandID=" & Trim(txtDemandNo.Text) & ""
    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
   
   
    mSql = "UPDATE faAllotmentLetters SET  intAgreementID=1  WHERE intAllotmentID=" & val(vsGrid.TextMatrix(vsGrid.Row, 9))
    objDB.ExecuteSP mSql, , , , mCnn, adCmdText
    
    cmdAdd.Enabled = False
    Call FillGrid
End Sub

Private Sub cmdSearchVoucher_Click()
    Dim mCnn        As New ADODB.Connection
    Dim objDB       As New clsDB
    Dim Rec         As New ADODB.Recordset
    Dim mSql        As String
    Dim mDemandNo   As Variant
    'Dim mSql        As String


    If cmbYear.ItemData(cmbYear.ListIndex) > 2011 Then
        mYearID = cmbYear.ItemData(cmbYear.ListIndex)
    Else
        Exit Sub
    End If

    frmSearchVouchers.CheckMode = 10
    frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateSerial(mYearID, 4, 1))
    frmSearchVouchers.txtToDate.Text = DdMmmYy(DateSerial(mYearID + 1, 3, 31))
    
    frmSearchVouchers.chkContra.Visible = False
    frmSearchVouchers.chkReceipt.Visible = True
    frmSearchVouchers.chkReceipt.value = 1
    frmSearchVouchers.chkJournal.Visible = False
    frmSearchVouchers.chkPayment.Visible = False
    frmSearchVouchers.txtFromDate.Enabled = False
    frmSearchVouchers.txtToDate.Enabled = False
    frmSearchVouchers.Show vbModal
    If gbSearchID <> -1 Then
       txtVoucherNo.Text = gbSearchCode
       txtVoucherNo.Tag = gbSearchID
       gbSearchCode = ""
       gbSearchID = -1
    End If


    If val(txtVoucherNo.Tag) > 0 Then
        If objDB.SetConnection(mCnn) Then
            mSql = " SELECT * FROM faVouchers "
            mSql = mSql + " WHERE intVoucherID = " & txtVoucherNo.Tag & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                 txtAmount.Text = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                 txtVrDate.Text = DdMmmYy(IIf(IsNull(Rec!dtDate), 0, Rec!dtDate))
                 mDemandNo = IIf(IsNull(Rec!intKeyID2), 0, Rec!intKeyID2)
                 txtDemandNo.Text = mDemandNo
            End If
            Rec.Close
         End If
         
         If mDemandNo < 1 Then
            mSql = "No Demand Issued for the Receipt." & vbCrLf
            mSql = mSql + " Either Cancel Or Reverse the Receipt " & vbCrLf
            mSql = mSql + " And Issue a new Receipt with Demand"
            MsgBox mSql, vbInformation, "Saankhya"
            txtVoucherNo.Text = ""
            txtVoucherNo.Tag = ""
            txtVrDate.Text = ""
            txtAmount.Text = ""
            Exit Sub
            
         End If
         If val(txtAmount.Text) <> val(vsGrid.TextMatrix(vsGrid.Row, 4)) Then
            MsgBox "Amount Of the receipt is not matching with Letter Of Authority", vbInformation, "Saankhya"
            txtVoucherNo.Text = ""
            txtVoucherNo.Tag = ""
            txtVrDate.Text = ""
            txtAmount.Text = ""
            Exit Sub
         End If
         If val(vsGrid.TextMatrix(vsGrid.Row, 15)) <> 0 Then
             If ValidateCreditAccHead = False Then
                MsgBox "Credit Account Head Of the receipt is not matching with Letter Of Authority", vbInformation, "Saankhya"
                txtVoucherNo.Text = ""
                txtVoucherNo.Tag = ""
                txtVrDate.Text = ""
                txtAmount.Text = ""
                Exit Sub
            End If
            
         End If
        
    End If
        
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
End Sub
Private Sub FormInitialize()
   Dim ctrl As Control
   For Each ctrl In Me.Controls
       If TypeOf ctrl Is TextBox Then
           ctrl.Text = ""
           ctrl.Tag = ""
       ElseIf TypeOf ctrl Is OptionButton Then
           ctrl.value = False
       ElseIf TypeOf ctrl Is ComboBox Then
           If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
           ctrl.Tag = ""
       End If
   Next
End Sub

Private Sub Form_Load()
    XPC.InitSubClassing
    Call FormInitialize
    Call FillCombo
    Call FillGrid
End Sub



Private Sub vsGrid_Click()
    If vsGrid.Row > 0 Then
           If vsGrid.TextMatrix(vsGrid.Row, 10) = "" And val(vsGrid.TextMatrix(vsGrid.Row, 16)) = 1 Then
                txtVoucherNo.Enabled = True
                txtVrDate.Enabled = True
                txtAmount.Enabled = True
                cmdSearchVoucher.Enabled = True
                cmdAdd.Enabled = True
            Else
                txtVoucherNo.Enabled = False
                txtVrDate.Enabled = False
                txtAmount.Enabled = False
                cmdSearchVoucher.Enabled = False
                cmdAdd.Enabled = False
            End If

    End If
End Sub
Private Function ValidateCreditAccHead() As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mGrAccHeadId As Integer
    Dim mVrCrAccHeadId As Integer
    
    mGrAccHeadId = val(vsGrid.TextMatrix(vsGrid.Row, 15))
    If objDB.SetConnection(mCnn) Then
        If val(txtVoucherNo.Tag) > 0 Then
            mSql = " SELECT * FROM faTransactionChild "
            mSql = mSql + " WHERE intTransactionID = (Select intTransactionID from faTransactions Where "
            mSql = mSql + " intVoucherID= " & val(txtVoucherNo.Tag) & " ) And tinDebitOrCreditFlag=0 "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                 mVrCrAccHeadId = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
            End If
            Rec.Close
        End If
        If mGrAccHeadId = mVrCrAccHeadId Then
            ValidateCreditAccHead = True
        Else
            ValidateCreditAccHead = False
        End If
    End If
End Function

