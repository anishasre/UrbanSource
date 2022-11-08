VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReceiptCancellationList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Receipt Cancellation List"
   ClientHeight    =   5685
   ClientLeft      =   465
   ClientTop       =   1620
   ClientWidth     =   11850
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   405
      Left            =   5445
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   1365
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   11670
      Top             =   5460
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   2670
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4830
      Width           =   2205
   End
   Begin VB.CheckBox chkCheckAll 
      Height          =   255
      Left            =   11250
      TabIndex        =   8
      Top             =   900
      Width           =   255
   End
   Begin VB.CommandButton cmdRemoveFromCancelList 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Remove From Cancel List"
      Height          =   405
      Left            =   90
      TabIndex        =   7
      Top             =   5130
      Width           =   2445
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   405
      Left            =   10200
      TabIndex        =   6
      Top             =   5160
      Width           =   1365
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   8820
      TabIndex        =   5
      Top             =   5160
      Width           =   1365
   End
   Begin VB.CommandButton cmdApproveCancellation 
      Caption         =   "&Approve Cancellation"
      Height          =   405
      Left            =   6840
      TabIndex        =   4
      Top             =   5160
      Width           =   1965
   End
   Begin VB.ComboBox cmbCounters 
      Height          =   345
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   450
      Width           =   3015
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3645
      Left            =   90
      TabIndex        =   1
      Top             =   870
      Width           =   11505
      _cx             =   20294
      _cy             =   6429
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
      Rows            =   13
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReceiptCancellationList.frx":0000
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   1
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
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   5220
      X2              =   5220
      Y1              =   4680
      Y2              =   5760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   225
      Left            =   1800
      TabIndex        =   9
      Top             =   4860
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   11700
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Counter Description"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Receipt Cancellation List"
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
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmReceiptCancellationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


    Private Sub FillCombo()
        On Error GoTo err:
            Dim mSql As String
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objdb As New clsDB
            
            Call PopulateList(cmbCounters, "Select vchDescription, intCounterID From faCounters WHERE intSectionID = 99 Order By vchDescription", , True, True, True, enuSourceString.Saankhya)
            
            If objdb.SetConnection(mCnn) Then
                mSql = "Select Count(*) as Cnt from faCounters WHERE intSectionID = 99"
                Rec.Open mSql, mCnn
                cmbCounters.AddItem "ALL COUNTERS"
                cmbCounters.ItemData(cmbCounters.NewIndex) = val(Rec!Cnt) + 1
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub chkCheckAll_Click()
        On Error GoTo err:
            Dim mRowCnt As Integer
            If chkCheckAll.Value = vbChecked Then
                For mRowCnt = 1 To vsGrid.Rows - 1
                    If val(vsGrid.TextMatrix(mRowCnt, 0)) <> 0 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked
                    End If
                Next
            ElseIf chkCheckAll.Value = vbUnchecked Then
                For mRowCnt = 1 To vsGrid.Rows - 1
                    If val(vsGrid.TextMatrix(mRowCnt, 0)) <> 0 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbUnchecked
                    End If
                Next
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmbCounters_Click()
        Call FillGrid
    End Sub

    Private Sub cmdApproveCancellation_Click()
        If GridValidations = False Then Exit Sub
        If MsgBox("Are you sure want to Approve the Cancel Request?", vbYesNo) = vbYes Then
            If ApproveCancelRequest = True Then
                MsgBox "Receipt Cancellation Request Approved Successfully", vbInformation
                Call FillGrid
            End If
        End If
    End Sub
    
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdClear_Click()
        vsGrid.Clear 1, 1
    End Sub

'''    Private Sub cmdReject_Click()           'ADDED BY MINU FOR REJECTIONS
'''        Dim mRowCnt As Integer
'''        If GridValidations = False Then Exit Sub
'''        For mRowCnt = 1 To vsGrid.Rows - 1
'''            If val(vsGrid.TextMatrix(mRowCnt, 0)) <> 0 And vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
'''                frmReject.Mode = 2
'''                frmReject.RequestTypeID = val(vsGrid.TextMatrix(mRowCnt, 7))
'''                frmReject.Show vbModal
'''                cmdReject.Enabled = False
'''                cmdApproveCancellation.Enabled = False
'''            End If
'''        Next
'''    End Sub

    Private Sub cmdRemoveFromCancelList_Click()
        If GridValidations = False Then Exit Sub
        If MsgBox("Are you sure want to Remove the Cancel Request?", vbYesNo) = vbYes Then
            If RemoveCancelRequest = True Then
                MsgBox "Successfully Removed from Cancellation List", vbInformation
                Call FillGrid
            End If
        End If
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Call FillCombo
        cmbCounters.Text = "ALL COUNTERS"
        WindowsXPC1.InitIDESubClassing
    End Sub
    
    Private Sub FillGrid()
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objdb As New clsDB
            Dim mSql As String
            Dim mRowCnt As Integer
            Dim mStatList As String
            Dim mCounterID As Integer
            
            mCounterID = cmbCounters.ItemData(cmbCounters.ListIndex)
            
            If objdb.SetConnection(mCnn) Then
                mSql = "Select faVouchers.intVoucherID as VchID, numReceiptNo, faCancelledVouchers.intUserID,faCancelledVouchers.intCounterID,faCancelledVouchers.numSeatID, faVouchers.intTransactionTypeID, "
                mSql = mSql + "intReasonID,Convert(varchar(12),dtCancellationDate,101)as CancelDate, vchCancelReason, faCounters.vchDescription as CounterName, fltAmount , min(numStationaryNo) as StatStartNo,Max(numstationaryNo) as StatEndNo, vchCancelSeries"
                mSql = mSql + " from faCancelledVouchers "
                mSql = mSql + " Inner Join faCancelReason On faCancelledVouchers.intReasonID = faCancelReason.intCancelID "
                mSql = mSql + " Inner Join faCounters on faCancelledVouchers.intCounterID = faCounters.intCounterID "
                mSql = mSql + " Inner Join faVouchers on faCancelledVouchers.intVoucherID = faVouchers.intVoucherID "
                'mSql = mSql + " Where faVouchers.tnyCancelFlag <> 1 and faVouchers.tnyStatus <> 4 and Convert(varchar(12),dtCancellationDate,101) = Convert(varchar(12),getdate(),101) and faCancelledVouchers.tnyRemoveCancel is null"
                mSql = mSql + " Where faCancelledVouchers.tnyApproveStatus = 0 and isnull(faCancelledVouchers.tnyRemoveCancel,0) <> 1 and Convert(varchar(12),dtCancellationDate,101) = Convert(varchar(12),getdate(),101) "
                If cmbCounters.Text = "ALL COUNTERS" Then
                    mSql = mSql + " and convert(varchar(10), faCancelledVouchers.intCounterID) Like '%' "
                Else
                    mSql = mSql + " and convert(varchar(10), faCancelledVouchers.intCounterID) Like '" & mCounterID & "'"
                End If
                mSql = mSql + " Group By faVouchers.intVoucherID , numReceiptNo, faCancelledVouchers.intUserID, "
                mSql = mSql + " faCancelledVouchers.intCounterID,faCancelledVouchers.numSeatID, intReasonID, "
                mSql = mSql + " dtCancellationDate, vchCancelReason, "
                mSql = mSql + "faCounters.vchDescription , fltAmount, vchCancelSeries, faVouchers.intTransactionTypeID "
                mSql = mSql + " Order By numReceiptNo "
                Rec.Open mSql, mCnn
                mRowCnt = 1
                vsGrid.Clear 1, 1
                vsGrid.Rows = 2
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!numReceiptNo), "", Rec!numReceiptNo)
                    If Not IsNull(Rec!StatStartNo) Then
                        If Rec!StatStartNo <> Rec!StatEndNo Then
                            mStatList = CStr(Rec!StatStartNo) + " - " + CStr(Rec!StatEndNo)
                        Else
                            mStatList = CStr(Rec!StatStartNo)
                        End If
                    End If
                    vsGrid.TextMatrix(mRowCnt, 1) = mStatList 'IIf(IsNull(Rec!numStationaryNo), "", Rec!numStationaryNo)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchCancelReason), "", Rec!vchCancelReason)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intReasonID), "", Rec!intReasonID)
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!CounterName), "", Rec!CounterName)
                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!VchID), "", Rec!VchID)
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!vchCancelSeries), "", Rec!vchCancelSeries)
                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                    vsGrid.Rows = vsGrid.Rows + 1
                Wend
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Function ApproveCancelRequest() As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim mSql As String
            Dim objdb As New clsDB
            Dim mRowCnt As Integer
            
            If objdb.SetConnection(mCnn) Then
                For mRowCnt = 1 To vsGrid.Rows - 1
                    If val(vsGrid.TextMatrix(mRowCnt, 0)) <> 0 And vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
                        If val(vsGrid.TextMatrix(mRowCnt, 4)) = 5 Then
                        
'                            mSql = "Update faVouchers Set tnysync=Null,tnyStatus = 4 Where intVoucherID = " & val(vsGrid.TextMatrix(mRowCnt, 7))
'                            mCnn.Execute mSql
                        Else
                            mSql = "Update faVouchers Set tnysync=Null,tnyStatus = 4, tnyCancelFlag = 1,tnyChangeFlag=Null Where intVoucherID = " & val(vsGrid.TextMatrix(mRowCnt, 7))
                            mCnn.Execute mSql
                            
                            mSql = "Update faTransactions Set tnysync=Null,tnyStatus = 4 Where intVoucherID = " & val(vsGrid.TextMatrix(mRowCnt, 7))
                            mCnn.Execute mSql
                        End If
                        mSql = "Update faCancelledVouchers Set tnyApproveStatus = 1 Where vchCancelSeries = '" & vsGrid.TextMatrix(mRowCnt, 8) & "'"
                        mCnn.Execute mSql
                        
                        '**********************************************************************************************************************
                            Call UpdateVoucherIndex(val(vsGrid.TextMatrix(mRowCnt, 7)))    'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
                        '**********************************************************************************************************************
                        
                        
''                        If (Val(vsGrid.TextMatrix(mRowCnt, 9)) = gbTransactionTypePTax And Val(vsGrid.TextMatrix(mRowCnt, 4)) <> 5) Then
''                            Call CancelPropertyTax(Val(vsGrid.TextMatrix(mRowCnt, 0)), Val(vsGrid.TextMatrix(mRowCnt, 7)))
''                        End If
''
''                        If Val(vsGrid.TextMatrix(mRowCnt, 9)) = gbTransactionTypeZonalCollection Then
''                            Call CancelHODemand(Val(vsGrid.TextMatrix(mRowCnt, 7)), mCnn)
''                        End If
                    End If
                    If CheckZonal = True Then       'ADDED BY ANJU FOR ZONAL
                        mSql = "UPDATE faSyncLog SET tnyVouchers=NULL,tnyVoucherChild=NULL,tnyVoucherSub=NULL,"
                        mSql = mSql + " tnyVoucherAddress=NULL,tnyTransactions=NULL,tnyTransactionChild=NULL,intFinancialYearID=NULL,"
                        mSql = mSql + " tnySyncStatus=NULL WHERE intVoucherID=" & val(vsGrid.TextMatrix(mRowCnt, 7))
                        mCnn.Execute mSql
                    End If
                    '---------------For Cochin Corporation --------------
'''''                    If gbLocalBodyID = 169 Then     ''05/12/2015 for Cochin Corpn
'''''                        If (val(vsGrid.TextMatrix(mRowCnt, 9)) = gbTransactionTypePTax And val(vsGrid.TextMatrix(mRowCnt, 4)) <> 5) Then
'''''                            ProTaxTCSDemand (val(vsGrid.TextMatrix(mRowCnt, 7)))
'''''                        End If
                    '---------------For Cochin Corporation --------------
                   ' Else  anju
                   
                    If gbFetchDemandFromWeb = 1 Then
                        If (val(vsGrid.TextMatrix(mRowCnt, 9)) = gbTransactionTypePTax And val(vsGrid.TextMatrix(mRowCnt, 4)) <> 5) Then
                        
                            PTaxWebDemand (val(vsGrid.TextMatrix(mRowCnt, 7)))
                        End If
                    End If
                    If gbLinkWithDandOWeb = 1 And (val(vsGrid.TextMatrix(mRowCnt, 9))) = gbTransactionTypeDandO Then
                        CancelDAndODemand ((val(vsGrid.TextMatrix(mRowCnt, 7))))
                    End If
                    If gbLinkWithProfTradeWeb = 1 And (val(vsGrid.TextMatrix(mRowCnt, 9))) = gbTransactionTypeProfTaxTrade Then
                        Call ProfTaxTradeDemand((val(vsGrid.TextMatrix(mRowCnt, 7))), val(vsGrid.TextMatrix(mRowCnt, 9)))
                    End If
                    If gbLinkWithProfEmpWeb = 1 And (val(vsGrid.TextMatrix(mRowCnt, 9))) = gbTransactionTypeProfTaxEmp Then
                        Call ProfTaxTradeDemand((val(vsGrid.TextMatrix(mRowCnt, 7))), val(vsGrid.TextMatrix(mRowCnt, 9)))
                    End If
                Next

            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
            ApproveCancelRequest = True
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function PTaxDemand()
        
    End Function
    Private Function ProfTaxTradeDemand(mVoucherID As Double, mTrTypeID As Integer) As Boolean
        On Error GoTo err:
        Dim objSOAP             As Variant
        Dim mArrOutDemandRecpt  As Variant
        Dim mUrl                As String
        Dim mCollPost           As String
        Dim mArrOut             As Variant
        Dim Rec                 As New ADODB.Recordset
        Dim mInstID             As Variant
        Dim mInstType           As Integer
        Dim mRowCnt             As Integer
        
            Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
            If (mTrTypeID = gbTransactionTypeProfTaxTrade) Then
                mInstType = 1
            ElseIf (mTrTypeID = gbTransactionTypeProfTaxEmp) Then
                mInstType = 2
            End If
            
            If Not (Rec.EOF Or Rec.BOF) Then
                mInstID = Rec!numSubLedgerID
            End If
            
             On Error Resume Next
           Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
            mUrl = gbDefaultUrl
           On Error Resume Next
            objSOAP.MSSoapInit (mUrl + "?WSDL")
            
            'mCollPost = CStr(mInstID) + "#" + CStr(gbLocalBodyID) + "#" + CStr(mVoucherID) + "#" + CStr(mInstType) + "#"
            mCollPost = CStr(mInstID) + "#" + CStr(gbLocalBodyID) + "#" + CStr(mVoucherID) + "#"

            mArrOut = objSOAP.cancel_prof_receipt(mCollPost)

            ProfTaxTradeDemand = True
err:
        MsgBox (Error$)
    End Function
'Cancel D&O Demand in Sanchaya
'Created by Syalima S On Jan 2018
Private Function CancelDAndODemand(ByVal mVoucherID As Double) As Boolean
        On Error GoTo err:
        Dim objSOAP             As Variant
        Dim mArrOutDemandRecpt As Variant
        Dim flagSankhya As Integer
        Dim mUrl As String
        Dim arrInput As String
        Dim mArrOut As Variant
        Dim Rec As New ADODB.Recordset
        Dim mDemandID As Variant
        
           Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
            If Not (Rec.EOF Or Rec.BOF) Then
                mDemandID = Rec!numDemandID
            End If
            
           Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
           
            mUrl = gbDefaultUrl
            On Error Resume Next
            objSOAP.MSSoapInit (mUrl + "?WSDL")
           
            'arrInput = Array(mDemandID, gbLocalBodyID, vsGrid.TextMatrix(8, 1))
            arrInput = mDemandID & "#" & gbLocalBodyID & "#" & mVoucherID
            mArrOut = (objSOAP.savecancellreceipt(arrInput))
            CancelDAndODemand = True
err:
        MsgBox (Error$)
    End Function

'    --------------------------------------------------
'    For Cochin Corporation ANJU
'    Revoke Receipt
'    --------------------------------------------------
    Private Sub ProTaxTCSDemand(mVoucherID As Long)
        Dim mMode As String
        Dim mColAccID       As String
        Dim mColKeyID       As String
        Dim mRevDate        As Date
        Dim xmlHttp         As Object
        Dim mXmlString      As Variant
        Dim oRS             As ADODB.Recordset
        Dim oNode           As Object 'MSXML2.IXMLDOMNode
        Dim oSubNodes       As Object 'MSXML2.IXMLDOMSelection
        Dim oDoc            As Object
        Dim mCollPost       As String
        Dim mUrl            As String
        Dim params          As String
              
        Set xmlHttp = CreateObject("MSXML2.xmlHttp")
        'If mTransactionType = 1 Then
        mRevDate = gbTransactionDate            'dateOfRevoking
        mMode = txtRemarks.Text                 'modeOfRevoking
        mCollPost = CStr(mVoucherID) + "~" + CStr(mMode) + "~" + Format(mRevDate, "yyyy-mm-dd")
        mUrl = gbDefaultUrl + "/updatePaymentRevoking?paymentRevokeParam=" + mCollPost
        xmlHttp.Open "POST", mUrl, False
        xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
        xmlHttp.send
        'MsgBox xmlHttp.responseText
        '-----------------------------------------------
        '------------------For Kochi Corp----------
        '-----------------------------------------------
    End Sub
        
     Private Function PTaxWebDemand(mVoucherID As Long)
                    Dim Rec             As New ADODB.Recordset
                    Dim mCollPost       As String
                    Dim mColZoneID      As String
                    Dim mBuildingIdWeb  As String
                    Dim mColAmt            As String
                    Dim mColDate        As String
                    Dim mColReceiptNo   As String
                    Dim mColBookNo      As String
                    Dim mColPeriodId     As String
                    Dim mColYearID       As String
                    Dim mHash           As String
                    Dim mCollOut        As String
'                    Dim node            As IXMLDOMNode
'                    Dim DataNodes       As IXMLDOMNodeList
                    Dim mUrl            As String
                    Dim objSOAP         As Variant
                    Dim mLen            As Integer
                    Dim mColAccID       As String
                    Dim mColKeyID       As String
                    'Dim Rec             As New ADODB.Recordset
                mUrl = gbDefaultUrlSanchayaPost
                Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                objSOAP.MSSoapInit mUrl + "?WSDL"
          
                    Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                    If Not (Rec.EOF And Rec.BOF) Then
                        While Not Rec.EOF
                            
                            mColAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            mColDate = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                            mColReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                            mColBookNo = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
                            mColPeriodId = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                            mColYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                            mBuildingIdWeb = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                            mColZoneID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                            mColAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                            mColKeyID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                            If mColAccID <> gbAcHeadIDPenalInterest Then
                                mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(mColZoneID) + "#" + CStr(mBuildingIdWeb) + "#"
                                mCollPost = mCollPost + CStr(mColYearID) + "#" + CStr(mColPeriodId) + "#" + CStr(mVoucherID) + "#"
                                mCollPost = mCollPost + CStr(mColBookNo) + "#" + CStr(mColReceiptNo) + "#" + CStr(mColDate) + "#"
                                mCollPost = mCollPost + CStr(gbFinancialYearID) + "#" + CStr(mColAmt) + "#" + CStr(gbLBName) + "#"
                                mCollPost = mCollPost + CStr(mColAccID) + "#" + CStr(mColKeyID)
                            End If
                            Rec.MoveNext
                            mCollPost = mCollPost + "~"
                        Wend
                        mLen = Len(mCollPost) - 1
                        mCollPost = Left$(mCollPost, mLen - 1)
                        mHash = CStr(mVoucherID) + CStr(mBuildingIdWeb) + "ikm#9567" + CStr(mColDate) + "*ikm#9567"
                        mCollOut = objSOAP.Saankhyaa_CollectionPostingCancel(mCollPost, mHash)
                    End If
  
    End Function
    Private Function RemoveCancelRequest() As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            Dim mRowCnt As Integer
            
            If txtRemarks.Text = "" Then
                MsgBox "Please Give Proper Remarks for Removing the Receipt from Cancellation List", vbInformation
                txtRemarks.SetFocus
                RemoveCancelRequest = False
                Exit Function
            End If
            
            If objdb.SetConnection(mCnn) Then
                For mRowCnt = 1 To vsGrid.Rows - 1
                    If val(vsGrid.TextMatrix(mRowCnt, 0)) <> 0 And vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
                        If val(vsGrid.TextMatrix(mRowCnt, 4)) = 5 Then
                            MsgBox "Printer Fault Cancellation is not Possible to Remove from the list. The Stationary once Printed is Used!", vbInformation
                            RemoveCancelRequest = False
                            Exit Function
                        End If
                        mSql = "Update faCancelledVouchers Set vchRemarks = '" & Trim(txtRemarks.Text) & "',tnyRemoveCancel = 1 Where vchCancelSeries = '" & Trim(vsGrid.TextMatrix(mRowCnt, 8)) & "'"
                        mCnn.Execute mSql
                    End If
                Next
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
            RemoveCancelRequest = True
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function GridValidations() As Boolean
        Dim mRowCnt As Integer
        Dim mChkFlag As Boolean
        
        If vsGrid.TextMatrix(1, 1) = "" Then
            GridValidations = False
            Exit Function
        End If
        mChkFlag = False
        For mRowCnt = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
                mChkFlag = True
            End If
        Next
        If mChkFlag = False Then
            GridValidations = False
            Exit Function
        End If
        
        GridValidations = True
    End Function
    
'''    Private Function CancelPropertyTax(txtRecieptNo As Double, mVoucherID As Long)
'''        On Error GoTo Err:
'''            Dim mCnn As New ADODB.Connection
'''            'Dim mCnnSanchaya As New ADODB.Connection
'''            Dim Rec As New Recordset
'''            Dim mSql As String
'''            Dim objDb As New clsDB
'''            Dim arrIn As Variant
'''            Dim mQry As String
'''
'''
'''            Dim blnConfig As Boolean
'''            Dim blnOtherZoneOfficeFlag As Boolean
'''
'''            If objDb.SetConnection(mCnn) Then
'''                mQry = "Select tnyLinkWithPropertyTax from faConfig"
'''                Rec.Open mQry, mCnn
'''                If IsNull(Rec!tnyLinkWithPropertyTax) Then
'''                    blnConfig = False
'''                ElseIf Val(Rec!tnyLinkWithPropertyTax) = 1 Then
'''                    blnConfig = True
'''                Else
'''                    blnConfig = False
'''                End If
'''                If Rec.State = 1 Then Rec.Close
'''
'''                mSql = "Select numZoneID as ZoneID from faVouchers Where intVoucherNo = " & Trim(txtRecieptNo)
'''                Rec.Open mSql, mCnn
'''                If Not (Rec.EOF Or Rec.BOF) Then
'''                    If Rec!ZoneID <> gbLocationID Then
'''                        blnOtherZoneOfficeFlag = True
'''                    Else
'''                        blnOtherZoneOfficeFlag = False
'''                    End If
'''                End If
'''            Else
'''                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
'''            End If
'''
'''            If blnConfig = True Then
'''                Set mCnn = Nothing
'''                If objDb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
'''                    If blnOtherZoneOfficeFlag = False Then
'''                        arrIn = Array(Trim(txtRecieptNo))
'''                        objDb.ExecuteSP "spReverseDemandFromSaankhya", arrIn, , , mCnn
'''                    Else
'''                        '---------------------------------------------------------------'
'''                        ' Other Zone Office Collection Modified on 13-aug-2009 By cijith'
'''                        '---------------------------------------------------------------'
'''                        arrIn = Array(gbLocationID, mVoucherID)
'''                        objDb.ExecuteSP "HOSaanOtherCollectionCancel", arrIn, , , mCnn
'''                        '----------------------------------------------------------'
'''                    End If
'''                Else
'''                    MsgBox "Connection To Sanchaya Does not Exist, Please Contact your System Administrtor", vbInformation
'''                End If
'''            End If
'''        Exit Function
'''Err:
'''        MsgBox (Error$)
'''    End Function
'''
'''    Private Function CancelHODemand(ByVal mVoucherID As Long, ByRef mCnn As ADODB.Connection) As Boolean
'''        On Error GoTo Err:
'''            Dim objDb As New clsDB
'''            Dim mCnnHO As New ADODB.Connection
'''            Dim Rec As New ADODB.Recordset
'''            Dim mSql As String
'''            Dim mDemandID As Variant
'''            Dim mLocationID As Variant
'''
'''
'''            If mCnn.State = 0 Then mCnn.Open
'''
'''            mSql = "Select * from faVouchers Where intVoucherID = " & mVoucherID
'''            Rec.Open mSql, mCnn
'''            If Not (Rec.EOF Or Rec.BOF) Then
'''                mDemandID = IIf(IsNull(Rec!intKeyID2), -1, Rec!intKeyID2)
'''                mLocationID = IIf(IsNull(Rec!numZoneID), -1, Rec!numZoneID)
'''            End If
'''
'''            If mLocationID = gbLocationID Then
'''                CancelHODemand = False
'''                Exit Function
'''            End If
'''
'''            If Rec.State = 1 Then Rec.Close
'''            If mCnn.State = 1 Then Set mCnn = Nothing
'''
'''            If objDb.CreateNewConnection(mCnnHO, enuSourceString.SaankhyaHO) Then
'''                mSql = "Update faIDemandTBL Set tnyStatus = 0 Where numDemandID = " & mDemandID
'''                mCnnHO.Execute mSql
'''                CancelHODemand = True
'''            Else
'''                MsgBox "Connection To FinanceHO does not exist, Please contact your System Admininstrator", vbInformation
'''                CancelHODemand = False
'''            End If
'''        Exit Function
'''Err:
'''        MsgBox (Error$)
'''    End Function
    Private Function CheckZonal() As Boolean         'ADDED BY ANJU FOR ZONAL
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mDemandID As Variant
        Dim mLocationID As Variant
    
    
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "Select * from faLBSettings"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
            mLocationID = IIf(IsNull(Rec!intLocationID), -1, Rec!intLocationID)
        End If
        If Right(mLocationID, 2) <> "01" Then
            CheckZonal = True
        Else
            CheckZonal = False
        End If
        Rec.Close
        mCnn.Close
    End Function
