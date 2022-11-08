VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCancelReceipt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cancel Receipt"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3450
      Top             =   5370
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   405
      Left            =   6690
      TabIndex        =   24
      Top             =   1920
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancelReceipt 
      Caption         =   "Cance&L Receipt"
      Height          =   405
      Left            =   8445
      TabIndex        =   9
      Top             =   5160
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clos&E"
      Height          =   405
      Left            =   10020
      TabIndex        =   10
      Top             =   5160
      Width           =   1485
   End
   Begin VB.TextBox txtStationaryNo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8880
      TabIndex        =   8
      Top             =   4305
      Width           =   1485
   End
   Begin VB.TextBox txtStationaryCount 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8880
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3840
      Width           =   1485
   End
   Begin VB.ComboBox cmbReason 
      Height          =   390
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3270
      Width           =   2595
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2955
      Left            =   90
      TabIndex        =   16
      Top             =   1950
      Width           =   6555
      _cx             =   11562
      _cy             =   5212
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
      BackColorFixed  =   13750782
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCancelReceipt.frx":0000
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   90
      TabIndex        =   12
      Top             =   300
      Width           =   11415
      Begin VB.ComboBox cmbInstrumentType 
         Height          =   390
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   3015
      End
      Begin VB.ComboBox cmbCounters 
         Height          =   390
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   3015
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   9870
         TabIndex        =   5
         Top             =   780
         Width           =   1365
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   405
         Left            =   8460
         TabIndex        =   4
         Top             =   780
         Width           =   1365
      End
      Begin VB.TextBox txtReceiptNoP2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   3
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtReceiptNoP1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8700
         MaxLength       =   6
         TabIndex        =   2
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9930
         TabIndex        =   21
         Top             =   270
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Type"
         Height          =   270
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter Description"
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Number"
         Height          =   270
         Left            =   6930
         TabIndex        =   13
         Top             =   330
         Width           =   1350
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   6720
      TabIndex        =   23
      Top             =   1620
      Width           =   4785
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Cancel Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   6720
      TabIndex        =   22
      Top             =   2820
      Width           =   4770
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   11730
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stationary Number"
      Height          =   270
      Left            =   6780
      TabIndex        =   20
      Top             =   4350
      Width           =   1545
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stationary Count"
      Height          =   270
      Left            =   6780
      TabIndex        =   19
      Top             =   3870
      Width           =   1395
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Cancellation"
      Height          =   270
      Left            =   6780
      TabIndex        =   18
      Top             =   3330
      Width           =   1965
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Receipt Information"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Top             =   1590
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Receipt Cancellation Form"
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
      Height          =   420
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   11415
   End
End
Attribute VB_Name = "frmCancelReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private intLoadMode         As Integer
    Private strReceiptNo        As String
    Private intInstrumentTypeID As Integer
    Private mRejected           As Boolean
    Dim mZonal As Integer 'Added by sunil

    Private Sub FormInitialize()
        cmbCounters.ListIndex = -1
        cmbInstrumentType.ListIndex = -1
        cmbReason.ListIndex = -1
        txtReceiptNoP1.Text = ""
        txtReceiptNoP2.Text = ""
        txtUserName.Text = ""
        txtStationaryCount.Text = ""
        txtStationaryNo.Text = ""
        vsGrid.Clear 1
        txtStationaryCount.Enabled = True
    End Sub
    
    Private Sub cmbCounters_Click()
        txtReceiptNoP1.Text = ""
        txtReceiptNoP2.Text = ""
        If cmbCounters.ListIndex <> -1 And cmbInstrumentType.ListIndex <> -1 And cmbCounters.ListIndex <> 0 And cmbInstrumentType.ListIndex <> 0 Then
            GetReceiptNoFirstPart
        End If
    End Sub
    
    Private Sub cmbInstrumentType_Click()
        txtReceiptNoP1.Text = ""
        txtReceiptNoP2.Text = ""
        If cmbCounters.ListIndex <> -1 And cmbInstrumentType.ListIndex <> -1 And cmbCounters.ListIndex <> 0 And cmbInstrumentType.ListIndex <> 0 Then
            GetReceiptNoFirstPart
        End If
    End Sub

    Private Sub cmbReason_Click()
        If cmbReason.ListIndex = -1 Then Exit Sub
        If cmbReason.ItemData(cmbReason.ListIndex) <> 5 Then
            txtStationaryCount.Text = 1
            txtStationaryCount.Enabled = False
        Else
            txtStationaryCount.Enabled = True
        End If
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdCancelReceipt_Click()
        On Error GoTo err:
            Dim mReceiptNo As Variant
            mReceiptNo = Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
            If SaveValidation = True Then
                If CancelReceipt = True Then
                    If (val(vsGrid.Tag) = gbTransactionTypePTax And val(cmbReason.ItemData(cmbReason.ListIndex)) <> 5) Then
                        Call CancelPropertyTax(CDbl(mReceiptNo), val(vsGrid.TextMatrix(8, 1)))
                    End If
                    
                    If val(vsGrid.Tag) = gbTransactionTypeZonalCollection Then
                        Call CancelHODemand(val(vsGrid.TextMatrix(8, 1)))
                    End If
                    
                    If (val(vsGrid.Tag) = gbTransactionTypePermitFeeFromKMBR And val(cmbReason.ItemData(cmbReason.ListIndex)) <> 5) Then
                        Call CancelKMBR(val(vsGrid.TextMatrix(8, 1)))
                    End If
                    If (val(vsGrid.Tag) = gbTransactionTypeRentOnBuilding And val(cmbReason.ItemData(cmbReason.ListIndex)) <> 5) Then
                        If gbLinkWithRentOnLand Then
                            Call CancelRLB(CStr(mReceiptNo))
                        End If
                    End If
                    If (val(vsGrid.Tag) = gbTransactionTypeProfTaxTrade And val(cmbReason.ItemData(cmbReason.ListIndex)) <> 5) Then
                        If gbLinkWithProfTaxEmp Then
                            Call cancelProfTaxInsts(CStr(mReceiptNo))
                        End If
                    End If
                    If ((val(vsGrid.Tag) = gbTransactionTypeDandO Or val(vsGrid.Tag) = gbTransactionTypePFA) And val(cmbReason.ItemData(cmbReason.ListIndex)) <> 5) Then
                        If gbLinkWithDandOPFA Then
                            If Not (CancelDOPFA(CStr(mReceiptNo))) Then
                                GoTo err
                            End If
                        End If
                    End If
                    
                    MsgBox "This Receipt Cancelled Successfully", vbInformation
                End If
                
                If cmbReason.ItemData(cmbReason.ListIndex) = 5 Then
                    frmReceiptsCounter.PRPReprintFlag = 1
                Else
                    frmReceiptsCounter.PRPReprintFlag = 0
                End If
                'Unload Me
                Call FormInitialize
                Call Form_Load
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdClear_Click()
        FormInitialize
    End Sub

    Private Sub cmdsearch_Click()
        If txtReceiptNoP1.Text <> "" And txtReceiptNoP2.Text <> "" Then
            Call GetReceiptDetails
        End If
    End Sub
    
    Private Sub Form_Activate()
        Me.Top = 1300
    End Sub

    Private Sub Form_Load()
        On Error GoTo err:
            Call FillCombos
            'Added By sunil babu For Zonal Collection  on 22-08-2011
            If mZonal = 1 Then
                cmbReason.ListIndex = 5
                cmbReason.Enabled = False
            End If
            CheckLastPostingDate
            WindowsXPC1.InitIDESubClassing
            Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Function FillCombos()
        On Error GoTo err:
            Call PopulateList(cmbCounters, "Select vchDescription, intCounterID From faCounters WHERE intSectionID = 99 Order By vchDescription", , True, True, True, enuSourceString.Saankhya)
            Call PopulateList(cmbInstrumentType, "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Order By vchInstrumentType", , True, True, True, enuSourceString.Saankhya)
            ShowComboTextForCounter
            If intLoadMode = 1 Then
                Call PopulateList(cmbReason, "Select vchCancelReason, intCancelID From faCancelReason Order By vchCancelReason", , True, True, True, enuSourceString.Saankhya)
                ShowComboTextForInstrument
                txtReceiptNoP1.Text = Left(strReceiptNo, 6)
                If Len(strReceiptNo) > 12 Then
                    txtReceiptNoP2.Text = Right(strReceiptNo, 6)
                Else
                    txtReceiptNoP2.Text = Right(strReceiptNo, 5)
                End If
                Frame1.Enabled = False
                GetReceiptDetails
                'cmbReason.SetFocus
            ElseIf intLoadMode = 2 Then
                Call PopulateList(cmbReason, "Select vchCancelReason, intCancelID From faCancelReason Where intCancelID<5 Order By vchCancelReason", , True, True, True, enuSourceString.Saankhya)
                txtStationaryCount.Text = 1
                txtStationaryCount.Enabled = False
                cmbInstrumentType.Text = "CASH"
            End If
            Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function GetReceiptNoFirstPart() As String
        On Error GoTo err:
            Dim mVoucherTypeID As String
            Dim mFinancialYearId As String
            Dim mCounterID As String
            Dim mInstTypeID As String
            Dim mVoucherIDP1 As String
            Dim mArrIn As Variant
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objdb As New clsDB
            Dim mArrOut As Variant
            
            mVoucherTypeID = "1"
            mFinancialYearId = CStr(Right(gbFinancialYearID, 2))
            mCounterID = Right("00" + LTrim(CStr(cmbCounters.ItemData(cmbCounters.ListIndex))), 2)
            mInstTypeID = Right("0" + LTrim(CStr(cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex))), 1)
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mArrIn = Array(cmbCounters.ItemData(cmbCounters.ListIndex), cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex), mFinancialYearId)
                Set Rec = objdb.ExecuteSP("spGetNextReceiptNo", mArrIn, mArrOut, , mCnn, adCmdStoredProc)
            End If
            'mVoucherIDP1 = mVoucherTypeID + mFinancialYearId + mCounterID + mInstTypeID
            mVoucherIDP1 = mArrOut(0, 0)
            txtReceiptNoP1.Text = Left(mVoucherIDP1, 6)
            GetReceiptNoFirstPart = mVoucherIDP1
            Exit Function
err:
        MsgBox (Error$)
        GetReceiptNoFirstPart = ""
    End Function
    
    Private Sub GetReceiptDetails()
        On Error GoTo err:
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim mSql As String
            Dim mVocherNo As String
            Dim objdb As New clsDB
            
            mVocherNo = Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
            mSql = "Select fltAmount, dtDate, vchDescription, numWardID, intDoorNoP1, vchDoorNoP2, vchName, vchTransactionType,faVouchers.intVoucherID, faVouchers.intTransactionTypeID,faVouchers.numZoneID,tnyVoucherGroupID From faVouchers "
            mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID "
            mSql = mSql + " Inner Join faTransactionType on faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID "
            mSql = mSql + " Where intVoucherNo = " & CDbl(mVocherNo)
            If objdb.SetConnection(mCnn) Then
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!dtDate < gbTransactionDate Then
                        MsgBox "You are not allowed to Cancel an Old Receipt!!!", vbInformation
                        Exit Sub
                    End If
                    If Rec!tnyVoucherGroupID = 4 Then
                        MsgBox "You are not allowed to Cancel Interrupted Receipt !!!", vbInformation
                        Exit Sub
                    End If
                    vsGrid.TextMatrix(0, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsGrid.TextMatrix(3, 1) = IIf(IsNull(Rec!numWardId), "", Rec!numWardId)
                    vsGrid.TextMatrix(4, 1) = IIf(IsNull(Rec!intDoorNoP1), "", Rec!intDoorNoP1)
                    vsGrid.TextMatrix(5, 1) = IIf(IsNull(Rec!vchDoorNoP2), "", Rec!vchDoorNoP2)
                    vsGrid.TextMatrix(6, 1) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    vsGrid.TextMatrix(7, 1) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                    vsGrid.TextMatrix(8, 1) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    vsGrid.Tag = IIf(IsNull(Rec!intTransactionTypeID), -1, Rec!intTransactionTypeID)
                    vsGrid.TextMatrix(9, 1) = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                    vsGrid.RowHidden(9) = True
                End If
                txtUserName.Text = gbUserName
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
            End If
            Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Function CancelReceipt() As Boolean
        On Error GoTo err:
            Dim aryIn As Variant
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim mStatCount As Integer
            Dim mStatNo As Long
            Dim Rec As New ADODB.Recordset
            Dim mFormat As String
            Dim mSql    As String
            
            mStatCount = val(txtStationaryCount.Text)
            mStatNo = val(txtStationaryNo.Text)
            If objdb.SetConnection(mCnn) Then
                If mRejected = False Then
                    Rec.Open "Select Count(*) as Cnt from faCancelledVouchers", mCnn
                    If Not (Rec.EOF Or Rec.BOF) Then
                        mFormat = Rec!Cnt
                    End If
                    If Rec.State = 1 Then Rec.Close
                    While (mStatCount)
                        aryIn = Array(vsGrid.TextMatrix(8, 1), _
                                    10, _
                                    Null, _
                                    Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text), _
                                    gbUserID, _
                                    gbCounterID, _
                                    gbSeatID, _
                                    cmbReason.ItemData(cmbReason.ListIndex), _
                                    gbTransactionDate, _
                                    mStatNo, _
                                    mFormat)
                        objdb.ExecuteSP "spSaveCancelledVouchers", aryIn, , , mCnn, adCmdStoredProc
                        mStatCount = mStatCount - 1
                        mStatNo = mStatNo + 1
                    Wend
                Else
                    mSql = "Update faCancelledVouchers Set tnyRemoveCancel=0 Where numReceiptNo =" & Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                End If
            End If
            CancelReceipt = True
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function ShowComboTextForCounter() 'ControlName As ComboBox, ByRef intIndex As Variant) As String
        On Error GoTo err:
            Dim intcnt As Integer
            For intcnt = 0 To cmbCounters.ListCount - 1
                If cmbCounters.ItemData(intcnt) = gbCounterID Then
                    cmbCounters.ListIndex = intcnt
                    Exit Function
                End If
            Next
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function ShowComboTextForInstrument() 'ControlName As ComboBox, ByRef intIndex As Variant) As String
        On Error GoTo err:
            Dim intcnt As Integer
            For intcnt = 0 To cmbInstrumentType.ListCount - 1
                If cmbInstrumentType.ItemData(intcnt) = intInstrumentTypeID Then
                    cmbInstrumentType.ListIndex = intcnt
                    Exit Function
                End If
            Next
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function SaveValidation() As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
        
            If cmbReason.ListIndex = -1 Then
                MsgBox "Please Select the Reason for Cancellation", vbInformation
                cmbReason.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If val(txtStationaryCount.Text) < 1 Then
                MsgBox "Please Give the Stationary Count", vbInformation
                txtStationaryCount.SetFocus
                SaveValidation = False
                Exit Function
            End If
            
            If txtStationaryNo.Text = "" Then
                MsgBox "Please Give the Stationary Number", vbInformation
                txtStationaryNo.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If val(vsGrid.TextMatrix(8, 1)) = 0 Then
                MsgBox "Please Select a particular Receipt for Cancellation", vbInformation
                cmdSearch.SetFocus
                SaveValidation = False
                Exit Function
            End If
            
            If cmbReason.ItemData(cmbReason.ListIndex) <> 5 Then
                If objdb.SetConnection(mCnn) Then
                    mSql = "Select * from faCancelledVouchers Where numReceiptNo = " & Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
                    Rec.Open mSql, mCnn
                    While Not (Rec.EOF Or Rec.BOF)
                        If Rec!tnyRemoveCancel = 1 Then
                            mRejected = True
                        Else
                            If Rec!intReasonID <> 5 Then
                                MsgBox "This Receipt is already in the list of Cancellation, Waiting for the Approval. So you are not allowed to Cancel again.", vbInformation
                                SaveValidation = False
                                Exit Function
                            End If
                        End If
                        Rec.MoveNext
                    Wend
                    Rec.Close
                Else
                
                    MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
                End If
            End If
            ''''---------Added On 26 Mar 2015 By Anisha C------------------------------------------
            If cmbReason.ItemData(cmbReason.ListIndex) <> 5 Then
                If objdb.SetConnection(mCnn) Then
                    mSql = "Select * from faVouchers Where tnyVoucherGroupID=2 and numLinkKeyID = " & Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF Or Rec.BOF) Then
                        MsgBox "This Receipt is done Adjustment Journal. So you are not allowed to Cancel again.", vbInformation
                        SaveValidation = False
                        Exit Function
                    Else
                        SaveValidation = True
                        'Exit Function
                    End If
                    Rec.Close
                Else
                    MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
                End If
            End If
            
            ''''----------------------------------------------------------------------------------------------------------------------------------
            
             ''''---------RECEIPTS GENERATED THROUGH NEW ACR MODE------------------------------------------
            If cmbReason.ItemData(cmbReason.ListIndex) <> 5 Then
                If objdb.SetConnection(mCnn) Then
                    mSql = "Select * from faVouchers Where intExternalModuleID=1 and intVoucherNo = " & Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF Or Rec.BOF) Then
                        MsgBox "This is an Autogenerated receipt of Development Expenditure.So you are not allowed to Cancel.", vbInformation
                        SaveValidation = False
                        Exit Function
                    Else
                        SaveValidation = True
                        'Exit Function
                    End If
                    Rec.Close
                Else
                    MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
                End If
            End If
            
            ''''----------------------------------------------------------------------------------------------------------------------------------
    
             ''''---------Reversed Voucher ------------------------------------------
            If cmbReason.ItemData(cmbReason.ListIndex) <> 5 Then
                If objdb.SetConnection(mCnn) Then
                    mSql = "Select isNull(tnyReversed,0)tnyReversed,isNull(intExternalApplicationID,0) ExtApp,* from faVouchers Where intVoucherNo = " & Trim(txtReceiptNoP1.Text) + Trim(txtReceiptNoP2.Text)
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF Or Rec.BOF) Then
                        If Rec!tnyReversed = 1 Then
                            MsgBox "This is Reversed Receipt can't do Cancellation", vbInformation
                            SaveValidation = False
                            Exit Function
                        End If
                        If Rec!ExtApp = 118 Then
                            MsgBox "This is Autogenerated Receipt of E bill can't do Cancellation", vbInformation
                            SaveValidation = False
                            Exit Function
                        End If
                    Else
                        SaveValidation = True
                        'Exit Function
                    End If
                    Rec.Close
                Else
                    MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
                End If
            End If
            
            ''''---------------------------------------------------------------------------------------------------------
            
            If cmbReason.ItemData(cmbReason.ListIndex) <> 5 And intInstrumentTypeID = 6 Then
                MsgBox "Letter of Allotment Vouchers cant cancel...Only Reverse..."
                cmbReason.SetFocus
                SaveValidation = False
                Exit Function
            End If
            
            SaveValidation = True
        Exit Function
err:
        MsgBox (Error$)
    End Function





    Private Sub txtStationaryCount_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub
        
    Private Sub txtStationaryCount_LostFocus()
        If val(txtStationaryCount) < 1 Then
            txtStationaryCount.Text = 1
        End If
    End Sub

    Private Sub txtStationaryNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub
    
    
    Private Function CancelPropertyTax(txtRecieptNo As Double, mVoucherID As Long)
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            'Dim mCnnSanchaya As New ADODB.Connection
            Dim Rec As New Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            Dim arrIn As Variant
            Dim mQry As String
            
            
            Dim blnConfig As Boolean
            Dim blnOtherZoneOfficeFlag As Boolean
            
            If objdb.SetConnection(mCnn) Then
                mQry = "Select tnyLinkWithPropertyTax from faConfig"
                Rec.Open mQry, mCnn
                If IsNull(Rec!tnyLinkWithPropertyTax) Then
                    blnConfig = False
                ElseIf val(Rec!tnyLinkWithPropertyTax) = 1 Then
                    blnConfig = True
                Else
                    blnConfig = False
                End If
                If Rec.State = 1 Then Rec.Close
                
                mSql = "Select numZoneID as ZoneID from faVouchers Where intVoucherNo = " & Trim(txtRecieptNo)
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!ZoneID <> gbLocationID Then
                        blnOtherZoneOfficeFlag = True
                    Else
                        blnOtherZoneOfficeFlag = False
                    End If
                End If
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
            
            If blnConfig = True Then
                Set mCnn = Nothing
                If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
                    If blnOtherZoneOfficeFlag = False Then
                        'arrIn = Array(Trim(txtRecieptNo))
                        arrIn = Array(gbLBID, gbLocationID, Trim(txtRecieptNo), gbTransactionDate)
                        objdb.ExecuteSP "spReverseDemandFromSaankhya", arrIn, , , mCnn
                  Else
                        '---------------------------------------------------------------'
                        ' Other Zone Office Collection Modified on 13-aug-2009 By cijith'
                        '---------------------------------------------------------------'
                        arrIn = Array(gbLocationID, mVoucherID)
                        objdb.ExecuteSP "HOSaanOtherCollectionCancel", arrIn, , , mCnn
                        '----------------------------------------------------------'
                    End If
                Else
                    MsgBox "Connection To Sanchaya Does not Exist, Please Contact your System Administrtor", vbInformation
                End If
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function CancelHODemand(ByVal mVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim objdb As New clsDB
            Dim mCnnHO As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim mDemandID As Variant
            Dim mLocationID As Variant
            Dim mCnn As New ADODB.Connection
            
            If objdb.SetConnection(mCnn) Then
                mSql = "Select * from faVouchers Where intVoucherID = " & mVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    mDemandID = IIf(IsNull(Rec!intKeyID2), -1, Rec!intKeyID2)
                    mLocationID = IIf(IsNull(Rec!numZoneID), -1, Rec!numZoneID)
                End If
                
                If mLocationID = gbLocationID Then
                    CancelHODemand = False
                    Exit Function
                End If
                
                If Rec.State = 1 Then Rec.Close
                If mCnn.State = 1 Then Set mCnn = Nothing
            Else
                MsgBox "Connection to Finance Does not Exist, Please Contact Your System Administrator", vbInformation
                CancelHODemand = False
                Exit Function
                
            End If
            
            If objdb.CreateNewConnection(mCnnHO, enuSourceString.SaankhyaHO) Then
                mSql = "Update faIDemandTBL Set tnyStatus = 0 Where numDemandID = " & mDemandID
                mCnnHO.Execute mSql
                CancelHODemand = True
            Else
                MsgBox "Connection To FinanceHO does not exist, Please contact your System Admininstrator", vbInformation
                CancelHODemand = False
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function

    Private Function CancelKMBR(ByVal mVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim objdb As New clsDB
            
            If objdb.CreateNewConnection(mCnn, enuSourceString.KMBR) Then
                aryIn = Array(mVoucherID)
                objdb.ExecuteSP "CancelPermitReceipt", aryIn, , , mCnn, adCmdStoredProc
                CancelKMBR = True
            Else
                MsgBox "Connection To KMMBR does not exist, Please contact your System Administrator", vbInformation
                CancelKMBR = False
            End If
            
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function CancelRLB(ByVal txtRecieptNo As Double) As Boolean
        'Cancel receipt for Rent On Land Integration
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim objdb As New clsDB
            
            If objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya) Then
                aryIn = Array(txtRecieptNo, vsGrid.TextMatrix(9, 1), 2)
                objdb.ExecuteSP "spSanSnRentDemandReverce", aryIn, , , mCnn, adCmdStoredProc
                CancelRLB = True
            Else
                MsgBox "Connection To Sanchaya does not exist, Please contact your System Administrator", vbInformation
                CancelRLB = False
            End If
            
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function CancelDOPFA(ByVal txtRecieptNo As Double) As Boolean
        'Cancel receipt for D&O and PFA Licence
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim objdb As New clsDB
            
            If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
'''                    (@numZoneId [decimal],
'''                    @intLBID   [int],
'''                    @numVoucherID  [numeric],
'''                    @vchDemandNo   [varchar](20),
'''                    @CancelDate    [int])
                aryIn = Array(vsGrid.TextMatrix(9, 1), gbLocalBodyID, txtRecieptNo, vsGrid.TextMatrix(8, 1), Format(vsGrid.TextMatrix(0, 1), "DD/mmm/yyyy"))
                objdb.ExecuteSP "spsnLicSanCancellationTBL_I", aryIn, , , mCnn, adCmdStoredProc
                CancelDOPFA = True
            Else
                MsgBox "Connection To SanchayaLite does not exist, Please contact your System Administrator", vbInformation
                CancelDOPFA = False
            End If
            
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function cancelProfTaxInsts(ByVal txtRecieptNo As Double) As Boolean
        'Cancel receipt for Prof.Tax Traders Integration
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim objdb As New clsDB
            
            If objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya) Then
                aryIn = Array(txtRecieptNo, vsGrid.TextMatrix(9, 1), 2)
                objdb.ExecuteSP "spSanSnProfTaxDemandReverse", aryIn, , , mCnn, adCmdStoredProc
                cancelProfTaxInsts = True
            Else
                MsgBox "Connection To Sanchaya does not exist, Please contact your System Administrator", vbInformation
                cancelProfTaxInsts = False
            End If
            
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Sub CheckLastPostingDate()   '-----------------LAST POSTING VALIDATION------------------
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim dtCurrentDate As Date
        
        Call SetgbLastPostingDate
        
        objdb.SetConnection mCnn
        mSql = "Select GETDATE()CurrentDate From faFinancialYear "
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            dtCurrentDate = Format(Rec!currentdate, "dd-mmm-yyyy")
            If CDate(dtCurrentDate) <= CDate(gbLastPostingDate) Then
                MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
                cmdCancelReceipt.Enabled = False
                Exit Sub
            End If
            
        End If
        
    End Sub
    Public Property Let LoadMode(mData As Integer)
        intLoadMode = mData
    End Property
    Public Property Let ReceiptNO(mData As String)
        strReceiptNo = mData
    End Property
    Public Property Let InstrumentTypeID(mData As Integer)
        intInstrumentTypeID = mData
    End Property
    Public Property Let ZonalCollection(mData As Integer) 'Added by Sunil Babu on 22-08-2011
        mZonal = mData
    End Property
