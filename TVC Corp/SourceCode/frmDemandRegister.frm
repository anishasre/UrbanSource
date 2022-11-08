VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmDemandRegister 
   Caption         =   "Demand Register"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   Icon            =   "frmDemandRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11820
   Begin VB.CommandButton Command1 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8430
      TabIndex        =   30
      Top             =   6090
      Width           =   1010
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   9540
      TabIndex        =   9
      Top             =   1050
      Width           =   1830
   End
   Begin VB.TextBox txtWard 
      Height          =   315
      Left            =   9540
      TabIndex        =   8
      Top             =   690
      Width           =   1830
   End
   Begin VB.TextBox txtDemandSuffix 
      Height          =   300
      Left            =   7425
      TabIndex        =   6
      Top             =   1050
      Width           =   975
   End
   Begin VB.TextBox txtDemandPrefix 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   6270
      TabIndex        =   5
      Top             =   1050
      Width           =   990
   End
   Begin VB.TextBox txtFromDate 
      Height          =   315
      Left            =   6270
      TabIndex        =   3
      Top             =   330
      Width           =   1830
   End
   Begin VB.TextBox txtToDate 
      Height          =   315
      Left            =   6270
      TabIndex        =   4
      Top             =   690
      Width           =   1830
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   9540
      TabIndex        =   7
      Top             =   330
      Width           =   1830
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11745
      TabIndex        =   15
      Top             =   6045
      Width           =   1620
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   1665
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   3315
   End
   Begin VB.ComboBox cmbTransactionType 
      Height          =   315
      Left            =   1665
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   3315
   End
   Begin VB.CommandButton cmdCancelDemand 
      Caption         =   "Cancel Demand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3450
      TabIndex        =   16
      Top             =   6075
      Width           =   1620
   End
   Begin VB.TextBox txtDemandNo 
      Height          =   285
      Left            =   5415
      TabIndex        =   12
      Top             =   5640
      Width           =   2070
   End
   Begin VB.CommandButton cmdSearchBySeat 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1515
      Width           =   1620
   End
   Begin VB.ComboBox cmbSeat 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1050
      Width           =   3300
   End
   Begin VB.CommandButton cmdViewDemand 
      Caption         =   "View Demand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5115
      TabIndex        =   13
      Top             =   6075
      Width           =   1620
   End
   Begin VB.CommandButton cmdPrintDemand 
      Caption         =   "Print Demand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6780
      TabIndex        =   14
      Top             =   6075
      Width           =   1620
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   11580
      Top             =   6315
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   8115
      TabIndex        =   17
      Top             =   315
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      _Version        =   393216
      Format          =   60686337
      CurrentDate     =   39697
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   8115
      TabIndex        =   18
      Top             =   675
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      _Version        =   393216
      Format          =   60686337
      CurrentDate     =   39698
   End
   Begin VSFlex8LCtl.VSFlexGrid vsDetails 
      Height          =   3510
      Left            =   180
      TabIndex        =   11
      Top             =   2025
      Width           =   11505
      _cx             =   20294
      _cy             =   6191
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
      BackColorBkg    =   -2147483633
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDemandRegister.frx":1CCA
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7305
      TabIndex        =   29
      Top             =   825
      Width           =   90
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
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
      Height          =   270
      Left            =   8790
      TabIndex        =   28
      Top             =   1050
      Width           =   675
   End
   Begin VB.Label lblWard 
      Caption         =   "Ward"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9000
      TabIndex        =   27
      Top             =   690
      Width           =   495
   End
   Begin VB.Label lblDemandNo1 
      AutoSize        =   -1  'True
      Caption         =   "Demand No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5220
      TabIndex        =   26
      Top             =   1050
      Width           =   990
   End
   Begin VB.Label lblName 
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
      Height          =   225
      Left            =   8970
      TabIndex        =   25
      Top             =   330
      Width           =   495
   End
   Begin VB.Label lblFromDate 
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
      Height          =   225
      Left            =   5280
      TabIndex        =   24
      Top             =   330
      Width           =   915
   End
   Begin VB.Label lblToDate 
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
      Height          =   255
      Left            =   5505
      TabIndex        =   23
      Top             =   690
      Width           =   690
   End
   Begin VB.Label lblSection 
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   22
      Top             =   330
      Width           =   645
   End
   Begin VB.Label lblTransactionType 
      Caption         =   "Transaction Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   105
      TabIndex        =   21
      Top             =   675
      Width           =   1515
   End
   Begin VB.Label lblDemandNo 
      Caption         =   "Demand No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4365
      TabIndex        =   20
      Top             =   5625
      Width           =   990
   End
   Begin VB.Label lblSeatName 
      Caption         =   "Seat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1230
      TabIndex        =   19
      Top             =   1050
      Width           =   375
   End
End
Attribute VB_Name = "frmDemandRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Option Explicit
    Dim gSeatID As Variant
    
    Private Sub FillvsDetails(Rec As ADODB.Recordset)
        Dim mRowCount       As Double
        Dim mSerialNo       As Double
        Dim mStatus         As Variant
        Dim mReceiptCancel  As Variant
        
        mRowCount = 1
        mSerialNo = 1
        vsDetails.Rows = 1
        While Not Rec.EOF
            vsDetails.Rows = vsDetails.Rows + 1
            vsDetails.TextMatrix(mRowCount, 0) = mSerialNo
            vsDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            vsDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            vsDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
            vsDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!dtDemandDate), "", CheckDateInMMM(Rec!dtDemandDate))
            vsDetails.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            vsDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
'            If IsNull(Rec!intVoucherID) Then
            If (mStatus = 9) Then
                vsDetails.TextMatrix(mRowCount, 7) = "Demand Cancelled"
            Else
                vsDetails.TextMatrix(mRowCount, 7) = "Demand Generated"
            End If
'            Else
            mReceiptCancel = IIf(IsNull(Rec!tnyCancelFlag), "", Rec!tnyCancelFlag)
            If mReceiptCancel <> "" Then
                If mReceiptCancel = 1 Then
                    vsDetails.TextMatrix(mRowCount, 7) = "Receipt Cancelled"
                Else
                    vsDetails.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo) & " - " & IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
                End If
            End If
'            End If
            mRowCount = mRowCount + 1
            mSerialNo = mSerialNo + 1
            Rec.MoveNext
        Wend
    End Sub
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmbDepartment_Click()
        Dim mcnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCount      As Double
        Dim mRowCount   As Integer
        Dim mSerialNo   As Integer
        Dim mStatus     As Variant
        
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
'        vsDetails.Clear 1, 1
'        vsDetails.Rows = 1
'        mSql = "Select Count(numDemandID) as Count From faIDemandTBL "
'        mSql = mSql + " Where intSectionID='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'        Rec.Open mSql, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            mCount = Rec!Count
'        End If
'        Rec.Close
'
'        mSql = "Select faVouchers.tnyCancelFlag,faVouchers.intVoucherNo,faVouchers.dtDate,faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate,Sum(faIDemandChild.fltAmount) as Amount"
'        mSql = mSql + " From faIDemandTBL"
'        mSql = mSql + " Left Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
'        mSql = mSql + " Left Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
'        mSql = mSql + " Left Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'        mSql = mSql + " Left Join faVouchers On faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
'        mSql = mSql + " Where faIDemandTBL.intSectionID='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'        mSql = mSql + " Group By faVouchers.tnyCancelFlag,faVouchers.intVoucherNo,faVouchers.dtDate,faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate"
'        mSql = mSql + " Order By faIDemandTBL.dtDemandDate Desc"
'        Rec.Open mSql, mCnn
'        Call FillvsDetails(Rec, mCount)
'        Rec.Close
        '**********************************************'
        '   Commented By Poornima On 07-June-2010
        '**********************************************'
'        If cmbDepartment.ItemData(cmbDepartment.ListIndex) = 99 Then
'            mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID =10 Order By vchTransactionType"
'        Else
'            mSql = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intSectionID='" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "'"
'        End If
        '**********************************************'
        '   Modified By Poornima On 07-June-2010
        '**********************************************'
            If cmbDepartment.ItemData(cmbDepartment.ListIndex) = 99 Then
                mSQL = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
                mSQL = mSQL + " FROM faSectionWiseTransactionTypes INNER JOIN "
                mSQL = mSQL + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                mSQL = mSQL + " Where (faTransactionType.intGroupID = 10) And faSectionWiseTransactionTypes.tnyList = 1"
                mSQL = mSQL + " ORDER BY faTransactionType.vchTransactionType"
            Else
                mSQL = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
                mSQL = mSQL + " FROM faSectionWiseTransactionTypes INNER JOIN "
                mSQL = mSQL + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                mSQL = mSQL + " Where (faTransactionType.intGroupID = 10)And faSectionWiseTransactionTypes.tnyList =1 And  faSectionWiseTransactionTypes.intSectionID =  " & cmbDepartment.ItemData(cmbDepartment.ListIndex)
                mSQL = mSQL + " ORDER BY faTransactionType.vchTransactionType"
            End If
        PopulateList cmbTransactionType, mSQL, , True, , True, enuSourceString.Saankhya
        mcnn.Close
    End Sub
    
    Private Sub cmbSeat_KeyPress(KeyAscii As Integer)
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub cmbSeat_LostFocus()
        Dim mcnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL    As String
        Dim Rec     As New ADODB.Recordset
        
        gSeatID = ""
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        mSQL = "Select numSeatID From DB_Masters..GL_Seats Where DB_Masters..GL_Seats.chvSeatTitle='" & cmbSeat.Text & "'"
        Rec.Open mSQL, mcnn
        If Not (Rec.EOF And Rec.BOF) Then
            gSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
        End If
        Rec.Close
        mcnn.Close
    End Sub

    Private Sub cmbTransactionType_Click()
'        Dim objDb       As New clsDB
'        Dim mCnn        As New ADODB.Connection
'        Dim Rec         As New ADODB.Recordset
'        Dim mSql        As String
'        Dim mRowCount   As Integer
'        Dim mCount      As Double
'        Dim mSerialNo   As Integer
'        Dim mStatus     As Variant
'
'        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        If cmbDepartment.ListIndex = -1 Then
'            MsgBox "Please select the Department", vbInformation
'            cmbDepartment.SetFocus
'            Exit Sub
'        End If
'        vsDetails.Clear 1, 1
'        vsDetails.Rows = 1
'        mSql = "Select Count(numDemandID) as Count From faIDemandTBL "
'        mSql = mSql + " Where intTransactionTypeID='" & cmbTransactionType.itemData(cmbTransactionType.ListIndex) & "'"
'        Rec.Open mSql, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            mCount = Rec!Count
'        End If
'        Rec.Close
'
'        mSql = "Select faVouchers.tnyCancelFlag,faIDemandTBL.tnyStatus,faVouchers.intVoucherNo,faVouchers.dtDate,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate,Sum(faIDemandChild.fltAmount) as Amount"
'        mSql = mSql + " From faIDemandTBL"
'        mSql = mSql + " Left Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
'        mSql = mSql + " Left Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
'        mSql = mSql + " Left Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'        mSql = mSql + " Left Join faVouchers On faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
'        mSql = mSql + " Where faIDemandTBL.intTransactionTypeID='" & cmbTransactionType.itemData(cmbTransactionType.ListIndex) & "'"
'        mSql = mSql + " Group By faVouchers.tnyCancelFlag,faVouchers.intVoucherNo,faVouchers.dtDate,faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate"
'        mSql = mSql + " Order By faIDemandTBL.dtDemandDate Desc"
'        Rec.Open mSql, mCnn
'        Call FillvsDetails(Rec, mCount)
'        Rec.Close
'        mCnn.Close
    End Sub

    Private Sub cmdCancelDemand_Click()
        Dim mcnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim mSQL        As String
        Dim Rec         As New ADODB.Recordset
        Dim mStatus     As Integer
        Dim mSeatID     As Variant
        
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        
        If txtDemandNo.Text = "" Then
            MsgBox "Please Enter the Demand No", vbInformation
            txtDemandNo.SetFocus
            Exit Sub
        Else
            mSQL = "Select numSeatID From faIDemandTBL Where vchDemandNo='" & txtDemandNo.Text & "'"
            Rec.Open mSQL, mcnn
            If Not (Rec.EOF And Rec.BOF) Then
                mSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
            Else
                MsgBox "Demand doesn't exists", vbInformation
                txtDemandNo.SetFocus
                Exit Sub
            End If
            Rec.Close
            If mSeatID <> "" Then
                If Trim(mSeatID) = Trim(gbSeatID) Or gbUserTypeID = 2 Then
                    mSQL = "Select tnyStatus From faIDemandTBL Where vchDemandNo='" & txtDemandNo.Text & "'"
                    Rec.Open mSQL, mcnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
                    Else
                        MsgBox "Demand Number doesn't exists", vbInformation
                        txtDemandNo.SetFocus
                        txtDemandNo.SelStart = 0
                        txtDemandNo.SelLength = Len(txtDemandNo.Text)
                        Exit Sub
                    End If
                    Rec.Close
                Else
                    MsgBox "You are not authorized to Cancel this Demand!!", vbCritical
                    Exit Sub
                End If
            End If
        End If
        mcnn.Close
        
        If mStatus = 9 Then
            MsgBox "This Demand is already Cancelled!", vbInformation
            Exit Sub
        End If
        If mStatus = 1 Then
            MsgBox "Can't cancel this Demand (Receipt Issued)"
            Exit Sub
        End If
        If mStatus = 0 Then
            objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
            If MsgBox("Are You Sure to Cancel the Demand ?", vbYesNo + vbDefaultButton2, "Confirm Cancellation") = vbNo Then
                Exit Sub
            End If
            mSQL = "Update faIDemandTBL"
            mSQL = mSQL + " Set tnyStatus='" & 9 & "'"
            mSQL = mSQL + " Where vchDemandNo='" & txtDemandNo.Text & "'"
            mcnn.Execute mSQL
            mcnn.Close
            MsgBox "Demand Cancelled Successfully!", vbInformation
            vsDetails.TextMatrix(vsDetails.Row, 7) = "Demand Cancelled"
        End If
    End Sub
    
    Private Sub cmdPrintDemand_Click()
        Dim objDB       As New clsDB
        Dim mcnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mDemandID   As Variant
        Dim mSQL        As String
        
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        If txtDemandNo.Text = "" Then
            MsgBox "Please enter the DemandNo", vbInformation
            txtDemandNo.SetFocus
            Exit Sub
        End If
        mSQL = "Select numDemandID From faIDemandTBL Where vchDemandNo='" & txtDemandNo.Text & "'"
        Rec.Open mSQL, mcnn
        If Not (Rec.EOF And Rec.BOF) Then
            mDemandID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
            If (MsgBox("Do You want to Print the Demand Slip ?", vbYesNo + vbDefaultButton2, "Confirm Printing") = vbYes) Then
                Call PrintDemandSlip(mDemandID, mcnn)
            Else
                Exit Sub
            End If
        Else
            MsgBox "Demand doesn't exists", vbInformation
        End If
        Rec.Close
    End Sub
    
    Private Sub cmdSearchBySeat_Click()
        Dim objDB               As New clsDB
        Dim mcnn                As New ADODB.Connection
        Dim mSQL                As String
        Dim Rec                 As New ADODB.Recordset
        Dim mCount              As Double
        Dim mDepartment         As String
        Dim mTransactionType    As String
        Dim mFromDate           As String
        Dim mToDate             As String
        Dim mName               As String
        Dim mSeatID             As String
       
        vsDetails.Clear 1, 1
        vsDetails.Rows = 1
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        
'        If txtSeat.Text = "" Then
'            MsgBox "Please Enter a Seat Name", vbCritical
'            txtSeat.SetFocus
'            Exit Sub
'        End If

'        If txtSeat.Tag <> "" Then
        If cmbSeat.Text = "" Then
            mSeatID = "%"
        Else
            mSeatID = gSeatID
        End If
        
        If cmbDepartment.ListIndex < 1 Then
            mDepartment = "%"
        Else
            mDepartment = CStr(cmbDepartment.ItemData(cmbDepartment.ListIndex))
        End If
        
        If cmbTransactionType.ListIndex < 1 Then
            mTransactionType = "%"
        Else
            mTransactionType = CStr(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
        End If
        
        If txtFromDate.Text = "" Then
            mSQL = "Select dtStartingDate From  faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Rec.Open mSQL, mcnn
            If Not (Rec.EOF And Rec.BOF) Then
                mFromDate = IIf(IsNull(Rec!dtStartingDate), "", CheckDateInMMM(Rec!dtStartingDate))
            End If
            Rec.Close
        Else
            mFromDate = txtFromDate.Text
        End If
        
        If txtToDate.Text = "" Then
            mSQL = "Select dtEndingDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Rec.Open mSQL, mcnn
            If Not (Rec.EOF And Rec.BOF) Then
                mToDate = IIf(IsNull(Rec!dtEndingDate), "", CheckDateInMMM(Rec!dtEndingDate))
            End If
            Rec.Close
        Else
            mToDate = txtToDate.Text
        End If
        
        If txtName.Text = "" Then
            mName = ""
        Else
            mName = CStr(txtName.Text)
        End If
           
'        mSql = "Select Count(faIDemandTBL.numDemandID) as Count From faIDemandTBL "
'       ' mSql = mSql + " Left Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
'        mSql = mSql + " Left Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
'        mSql = mSql + " Where numSeatID LIKE '" & mSeatID & "' "
'        mSql = mSql + " And faIDemandTBL.intSectionID LIKE '" & mDepartment & "'"
'        mSql = mSql + " And faIDemandTBL.intTransactionTypeID LIKE '" & mTransactionType & "'"
'        mSql = mSql + " And faIDemandAddress.vchName LIKE '" & "%" & mName & "%" & "'"
'        mSql = mSql + " And dtDemandDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'        If txtDemandPrefix.Text <> "" And txtDemandSuffix.Text <> "" Then
'            mSql = mSql + " And vchDemandNo  ='" & txtDemandPrefix.Text + "-" + txtDemandSuffix.Text & "'"
'        End If
'        If txtWard.Text <> "" Then
'            mSql = mSql + " And faIDemandTBL.intWardNo = " & txtWard.Text
'        End If
'        If txtAmount.Text <> "" Then
'            'mSql = mSql + " And faIDemandChild.fltAmount =" & txtAmount.Text
'        End If
'
'        Rec.Open mSql, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            mCount = Rec!count
'        End If
'        Rec.Close
        
        mSQL = "Select faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate,faVouchers.intVoucherNo,faVouchers.tnyCancelFlag,faVouchers.dtDate,Sum(faIDemandChild.fltAmount) as Amount"
        mSQL = mSQL + " From faIDemandTBL"
        mSQL = mSQL + " Left Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
        mSQL = mSQL + " Left Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
        mSQL = mSQL + " Left Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
        mSQL = mSQL + " Left Join faVouchers On faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
        mSQL = mSQL + " Where Isnull(faIDemandTBL.numSeatID,0) LIKE'" & mSeatID & "'"
        mSQL = mSQL + " And Isnull(faIDemandTBL.intSectionID,0) LIKE '" & mDepartment & "'"
        mSQL = mSQL + " And Isnull(faIDemandTBL.intTransactionTypeID,0) LIKE '" & mTransactionType & "'"
        mSQL = mSQL + " And ISNULL(faIDemandAddress.vchName,'') LIKE '" & "%" & mName & "%" & "'"
        mSQL = mSQL + " And dtDemandDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
        If txtDemandPrefix.Text <> "" And txtDemandSuffix.Text <> "" Then
            mSQL = mSQL + " And vchDemandNo  ='" & txtDemandPrefix.Text + "-" + txtDemandSuffix.Text & "'"
        End If
        If txtWard.Text <> "" Then
            mSQL = mSQL + " And faIDemandTBL.intWardNo = " & txtWard.Text
        End If
        If txtAmount.Text <> "" Then
            mSQL = mSQL + " And faIDemandChild.fltAmount =" & txtAmount.Text
        End If
        mSQL = mSQL + " Group By faVouchers.tnyCancelFlag,faVouchers.intVoucherNo,faVouchers.dtDate,faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate"
        mSQL = mSQL + " Order By faIDemandTBL.dtDemandDate Desc"
    
        Rec.Open mSQL, mcnn
        If Rec.EOF Or Rec.BOF Then
            MsgBox "No Records Exist!!", vbInformation
            Exit Sub
        End If
        Call FillvsDetails(Rec)
        Rec.Close
'        Else
'            MsgBox "Please Enter a valid Seat", vbInformation
'            txtSeat.SetFocus
'            Exit Sub
'        End If
    End Sub
    
    Private Sub cmdViewDemand_Click()
        Dim mcnn                As New ADODB.Connection
        Dim objDB               As New clsDB
        Dim mSQL                As String
        Dim Rec                 As New ADODB.Recordset
        
        Dim mDemandID           As Variant
        Dim mAccountHeadCode    As Variant
        Dim mAccountHead        As String
        Dim mAmount             As Variant
        Dim HeadRec             As New ADODB.Recordset
        Dim mRemarks            As String
        
        Dim mName               As String
        Dim mHouseName          As String
        Dim mStreet             As String
        Dim mLocalPlace         As String
        Dim mMainPlace          As String
        Dim mPost               As String
        Dim mPin                As String
        Dim mPhone              As String
        
        Dim mInstrumentTypeID   As Variant
        Dim mInstrumentNo       As String
        Dim mInstrumentDate     As String
        Dim mDrawnFrom          As String
        Dim mDrawnPlace         As String
        
        mAccountHead = ""
        mAmount = 0
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        If txtDemandNo.Text = "" Then
            MsgBox "Please enter the DemandNo", vbInformation
            txtDemandNo.SetFocus
            Exit Sub
        End If
        mSQL = "Select numDemandID,vchRemarks From faIDemandTBL Where vchDemandNo='" & txtDemandNo.Text & "'"
        Rec.Open mSQL, mcnn
        If Not (Rec.EOF And Rec.BOF) Then
            mDemandID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
            mRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
        Else
            MsgBox "Demand doesn't exists", vbInformation
        End If
        Rec.Close
        If mDemandID <> "" Then
            mSQL = "Select  * From faIDemandChild"
            mSQL = mSQL + " Inner Join faIDemandTBL On faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
            mSQL = mSQL + " Inner Join faIDemandAddress On faIDemandChild.numDemandID=faIDemandAddress.numDemandID"
            mSQL = mSQL + " Where faIDemandChild.numDemandID='" & mDemandID & "'"
            Rec.Open mSQL, mcnn
            While Not Rec.EOF
                mAccountHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                mHouseName = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
                mStreet = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                mLocalPlace = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                mMainPlace = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                mPost = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
                mPin = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
                mPhone = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                
                mInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                
                If mInstrumentTypeID = 5 Then
                    mInstrumentNo = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    mInstrumentDate = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    mDrawnFrom = IIf(IsNull(Rec!vchDrawnFrom), "", Rec!vchDrawnFrom)
                    mDrawnPlace = IIf(IsNull(Rec!vchDrawnPlace), "", Rec!vchDrawnPlace)
                End If
                
                If mAccountHeadCode <> "" Then
                    mSQL = "Select vchAccountHead From faAccountHeads Where vchAccountHeadCode='" & mAccountHeadCode & " '"
                    HeadRec.Open mSQL, mcnn
                    If Not (HeadRec.EOF And HeadRec.BOF) Then
                        mAccountHead = mAccountHead + IIf(IsNull(HeadRec!vchAccountHead), "", HeadRec!vchAccountHead) + "   : " + CStr(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)) + vbCrLf + vbCrLf
                        mAmount = mAmount + IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                    End If
                    HeadRec.Close
                End If
                Rec.MoveNext
            Wend
            Rec.Close
            If mInstrumentTypeID = 5 Then
                MsgBox mAccountHead + vbCrLf + "Total      : " & mAmount & vbCrLf + vbCrLf + vbCrLf + "                Name : " + mName + vbCrLf + "     House Name : " + mHouseName + vbCrLf + "               Street : " + mStreet + vbCrLf + "        Local Place : " + mLocalPlace + vbCrLf + "         Main Place : " + mMainPlace + vbCrLf + "        Post Office : " + mPost + vbCrLf + "            Pin Code : " + mPin + vbCrLf + "                Phone : " + mPhone + vbCrLf + "            Remarks : " + mRemarks + vbCrLf + vbCrLf + "    InstrumentNo : " + mInstrumentNo + vbCrLf + "Instrument Date : " + mInstrumentDate + vbCrLf + "       Drawn From : " + mDrawnFrom + vbCrLf + "       Drawn Place : " + mDrawnPlace, , "Details"
            Else
                MsgBox mAccountHead + vbCrLf + "Total      : " & mAmount & vbCrLf + vbCrLf + vbCrLf + "            Name : " + mName + vbCrLf + " House Name : " + mHouseName + vbCrLf + "           Street : " + mStreet + vbCrLf + "    Local Place : " + mLocalPlace + vbCrLf + "     Main Place : " + mMainPlace + vbCrLf + "    Post Office : " + mPost + vbCrLf + "        Pin Code : " + mPin + vbCrLf + "            Phone : " + mPhone + vbCrLf + "        Remarks : " + mRemarks, , "Details"
            End If
        End If
    End Sub

Private Sub Command1_Click()
    Dim mMonth  As Integer
    Dim frmNewRpt As New frmRptViewer
    Dim arInput As Variant
    Dim frmNewViewer As New frmRptViewer
        If txtFromDate.Text = "" Then
            MsgBox "Please Select From Date ", vbCritical
            txtFromDate.SetFocus
            Exit Sub
        End If
        If txtToDate.Text = "" Then
            MsgBox "Please select To Date ", vbCritical
            txtToDate.SetFocus
            Exit Sub
        End If
        arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text))
        frmNewRpt.rptFileName = App.Path & "\Reports\rptDemandRegister.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
End Sub

    Private Sub dtpFromDate_CloseUp()
        txtFromDate.Text = CheckDateInMMM(dtpFromDate.Value)
    End Sub

    Private Sub dtpToDate_CloseUp()
        txtToDate.Text = CheckDateInMMM(dtpToDate.Value)
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Dim objDB       As New clsDB
        Dim mcnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCount      As Double
        Dim mRowCount   As Integer
        Dim mSerialNo   As Integer
        Dim mStatus     As Variant
        
        frmDemandRegister.Width = 11940
        frmDemandRegister.Height = 7155
        XPC.InitIDESubClassing
        cmdSearch.Visible = False
        
        PopulateList cmbSeat, "SELECT chvSeatTitle,numSeatID FROM GL_Seats ORDER BY chvSeatTitle", , , True, , enuSourceString.DBMaster
        
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        vsDetails.Clear 1, 1
        vsDetails.Rows = 1
        mSQL = "Select Count(numDemandID) as Count From faIDemandTBL "
        Rec.Open mSQL, mcnn
        mCount = Rec!count
        Rec.Close
        
'        mSql = "Select faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate,faVouchers.intVoucherNo,faVouchers.tnyCancelFlag,faVouchers.dtDate,Sum(faIDemandChild.fltAmount) as Amount"
'        mSql = mSql + " From faIDemandTBL"
'        mSql = mSql + " Left Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
'        mSql = mSql + " Left Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
'        mSql = mSql + " Left Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'        mSql = mSql + " Left Join faVouchers On faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
'        mSql = mSql + " Group By faVouchers.tnyCancelFlag,faVouchers.intVoucherNo,faVouchers.dtDate,faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate"
'        mSql = mSql + " Order By faIDemandTBL.dtDemandDate Desc"
'        Rec.Open mSql, mCnn
'        Call FillvsDetails(Rec, mCount)
'        Rec.Close
        
        mSQL = "Select vchSectionName,intSectionID From faSection"
        PopulateList cmbDepartment, mSQL, , True, , True, enuSourceString.Saankhya
        mcnn.Close
    End Sub
    
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

     Private Sub txtDemandNo_GotFocus()
        txtDemandNo.SelStart = 0
        txtDemandNo.SelLength = Len(txtDemandNo.Text)
    End Sub

    Private Sub txtDemandPrefix_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDemandSuffix_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
    End Sub

    Private Sub txtFromDate_LostFocus()
        If Trim(txtFromDate.Text) <> "" Then
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub

    Private Sub txtName_GotFocus()
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName)
    End Sub
   
    Private Sub txtToDate_GotFocus()
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate)
    End Sub

    Private Sub txtToDate_LostFocus()
        If Trim(txtToDate.Text) <> "" Then
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        End If
    End Sub
 
    Private Sub txtWard_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsDetails_Click()
        vsDetails.SelectionMode = flexSelectionByRow
        txtDemandNo.Text = vsDetails.TextMatrix(vsDetails.Row, 3)
        'txtDemandNo.SetFocus
    End Sub

    Private Sub vsDetails_DblClick()
        If vsDetails.TextMatrix(vsDetails.Row, 3) <> "" Then
            frmDemand.DemandNo = vsDetails.TextMatrix(vsDetails.Row, 3)
            frmDemand.Visible = True
        End If
    End Sub
