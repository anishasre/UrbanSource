VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfDemands 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand Register"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmListOfDemands.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00F4FAFA&
      Height          =   585
      Left            =   15
      ScaleHeight     =   525
      ScaleWidth      =   11760
      TabIndex        =   7
      Top             =   6090
      Width           =   11820
      Begin VB.CommandButton cmdSearchTransactionType 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5910
         TabIndex        =   2
         Top             =   135
         Width           =   315
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         TabIndex        =   1
         Top             =   135
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   7020
         TabIndex        =   3
         Top             =   135
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61931523
         CurrentDate     =   40429
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10155
         TabIndex        =   5
         Top             =   45
         Width           =   1395
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   0
         Top             =   30
         Width           =   1395
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   3270
         Top             =   510
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   8670
         TabIndex        =   4
         Top             =   135
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   40429
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   6555
         TabIndex        =   10
         Top             =   135
         Width           =   450
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   8460
         TabIndex        =   9
         Top             =   135
         Width           =   180
      End
      Begin VB.Label lblTransactionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1830
         TabIndex        =   8
         Top             =   135
         Width           =   1515
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsDetails 
      Height          =   6030
      Left            =   15
      TabIndex        =   6
      Top             =   15
      Width           =   11805
      _cx             =   20823
      _cy             =   10636
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfDemands.frx":1CCA
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
End
Attribute VB_Name = "frmListOfDemands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    '*********************************************************************************************'
    '               Form to list all the Demands generated by a particular Seat                   '
    '               In the case of Approving Officer, Lists only the demands need Approval        '
    '*********************************************************************************************'
    Private Sub FillDemand()
        Dim mCnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mFromDate       As String
        Dim mToDate         As String
        Dim mStatus         As Variant
        Dim mReceiptCancel  As Variant
        
        '*********************************************************************************************'
        '                                   Procedure to List Demands                                 '
        '*********************************************************************************************'
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mRowCount = 1
        vsDetails.Rows = 1
        vsDetails.Clear 1, 1
        mFromDate = CheckDateInMMM(dtpFromDate.value)
        mToDate = CheckDateInMMM(dtpToDate.value)
        
        If CDate(mFromDate) > CDate(mToDate) Then
            MsgBox "Please enter valid To Date", vbInformation
            dtpToDate.SetFocus
            Exit Sub
        End If
        mSql = "Select faIDemandTBL.tnyStatus,faIDemandTBL.intTransactionTypeID, faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate,faVouchers.intVoucherNo,faVouchers.tnyCancelFlag,faVouchers.dtDate,Sum(faIDemandChild.fltAmount) as Amount,faIDemandTBL.vchDemandNo As DemandNo,faIDemandTbl.intTransactionTypeID As TransactionTypeID,faIDemandTBL.numForwardedSeatID As ForwardedSeatID"
        'mSql = mSql + " ,dtTransactiondate,intDemandMode "
        mSql = mSql + " From faIDemandTBL"
        mSql = mSql + " Left Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
        mSql = mSql + " Left Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
        mSql = mSql + " Left Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
        mSql = mSql + " Left Join faVouchers On faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
        mSql = mSql + " Where dtDemandDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then 'In the case of Approving Officer, we need only demands need approval
            'mSQL = mSQL + " And faIDemandTBL.numForwardedSeatID =" & gbSeatID
            mSql = mSql + " And faIDemandTBL.numForwardedSeatID Is Not Null"
        Else
            mSql = mSql + " And faIDemandTBL.numSeatID =" & gbSeatID
        End If
        If txtTransactionType.Tag <> "" Then
            mSql = mSql + " And faIDemandTBL.intTransactionTypeID = " & txtTransactionType.Tag
        End If
        mSql = mSql + " Group By faVouchers.tnyCancelFlag,faVouchers.intVoucherNo,faVouchers.dtDate,faIDemandTBL.tnyStatus,faTransactionType.vchTransactionType,faIDemandTBL.intVoucherID,faIDemandAddress.intWardNo,faIDemandAddress.vchName,faIDemandTBL.vchDemandNo,faIDemandTBL.dtDemandDate,faIDemandTBL.vchDemandNo,faIDemandTbl.intTransactionTypeID,faIDemandTBL.numForwardedSeatID"
        mSql = mSql + " Order By faIDemandTBL.dtDemandDate Desc"
        Rec.Open mSql, mCnn
        While Not Rec.EOF
            vsDetails.Rows = vsDetails.Rows + 1
            vsDetails.TextMatrix(mRowCount, 0) = mRowCount
            vsDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            vsDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            vsDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
            vsDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!dtDemandDate), "", CheckDateInMMM(Rec!dtDemandDate))
            vsDetails.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            vsDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!amount), "", Rec!amount)
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
'            If IsNull(Rec!intVoucherID) Then
            If (mStatus = 9) Then
                vsDetails.TextMatrix(mRowCount, 7) = "Demand Cancelled"
            Else
                vsDetails.TextMatrix(mRowCount, 7) = "Demand Generated"
                If (mStatus = 8) Then
                    vsDetails.TextMatrix(mRowCount, 7) = "Waiting for Approval"
                    vsDetails.Cell(flexcpChecked, mRowCount, 8) = vbUnchecked
                Else
                    vsDetails.Cell(flexcpChecked, mRowCount, 8) = vbChecked
                End If
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
            vsDetails.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!DemandNo), "", Rec!DemandNo)
            vsDetails.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!TransactionTypeID), "", Rec!TransactionTypeID)
            vsDetails.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!ForwardedSeatID), "", Rec!ForwardedSeatID)
'            End If
            mRowCount = mRowCount + 1
            Rec.MoveNext
        Wend
    End Sub
    
    Private Sub cmdNew_Click()
        frmDemandInterface.PreviousYearMode = 0
        frmDemandInterface.PendingTaskReqID = -1
        frmDemandInterface.DemandNo = ""
        frmDemandInterface.Show vbModal
        FillDemand
    End Sub

    Private Sub cmdSearch_Click()
        FillDemand
    End Sub

    Private Sub cmdSearchTransactionType_Click()
        frmSearchTransactionType.Show vbModal
        txtTransactionType.SetFocus
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        dtpFromDate.value = gbTransactionDate
        dtpToDate.value = gbTransactionDate
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdNew.Enabled = False
        End If
        FillDemand
    End Sub
    
    Private Sub txtTransactionType_GotFocus()
        If gbSearchStr <> "" Then
            txtTransactionType.Text = gbSearchStr
            txtTransactionType.Tag = gbSearchID
            gbSearchCode = ""
            gbSearchID = -1
            gbSearchStr = ""
        Else
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        End If
    End Sub

    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyDelete Then
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        Else
            txtTransactionType.Locked = True
        End If
    End Sub
    
    Private Sub vsDetails_DblClick()
        If vsDetails.Row > 0 Then
            'If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                If vsDetails.TextMatrix(vsDetails.Row, 9) <> "" Then
                    If (vsDetails.TextMatrix(vsDetails.Row, 10) = gbTransactionTypeBFundSSSFund Or vsDetails.TextMatrix(vsDetails.Row, 10) = gbTransactionTypeMoneyOrderReturns) Then 'Demands need approval
                        If (vsDetails.Cell(flexcpChecked, vsDetails.Row, 8) = 2) Then
                            If gbSeatGroupID <> gbSeatGroupAccountsOfficer And gbSeatGroupID <> gbSeatGroupAccountsSuperintended Then
                                frmDemandInterface.DemandNo = vsDetails.TextMatrix(vsDetails.Row, 9)
                                frmDemandInterface.Show vbModal
                            Else
                                If vsDetails.TextMatrix(vsDetails.Row, 11) <> "" Then
                                    If vsDetails.TextMatrix(vsDetails.Row, 11) = gbSeatID Then
                                        frmDemandInterface.DemandNo = vsDetails.TextMatrix(vsDetails.Row, 9)
                                        frmDemandInterface.cmdSave.Caption = "Approve"
'                                        frmDemandInterface.cmdReject.Enabled = True      'ADDED BY MINU FOR REJECTIONS
                                        frmDemandInterface.Show vbModal
                                    Else
                                        MsgBox "You can't approve this Demand", vbInformation
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            MsgBox "Can't edit this Demand (Already approved)"
                        End If
                    Else
                        frmDemandInterface.DemandNo = vsDetails.TextMatrix(vsDetails.Row, 9)
                        frmDemandInterface.Show vbModal
                    End If
                    'End If
                End If
            'End If
            FillDemand
        End If
    End Sub
