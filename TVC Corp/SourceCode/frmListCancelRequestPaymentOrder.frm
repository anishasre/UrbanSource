VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListCancelRequestPaymentOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancellation Request for Payment Order"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12870
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmListCancelRequestPaymentOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Refresh"
      Height          =   420
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   0
      Width           =   1140
   End
   Begin WinXPC_Engine.WindowsXPC winXpc 
      Left            =   0
      Top             =   6750
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fraNewCancellationRequests 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   12840
      Begin VB.Frame fraPayOrder 
         BackColor       =   &H80000016&
         Caption         =   "Payorder Details"
         Enabled         =   0   'False
         Height          =   3480
         Left            =   7470
         TabIndex        =   31
         Top             =   900
         Visible         =   0   'False
         Width           =   5280
         Begin VB.TextBox txtUser 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   2790
            Width           =   3435
         End
         Begin VB.CheckBox chkApproved 
            Alignment       =   1  'Right Justify
            Caption         =   "Approved"
            Height          =   240
            Left            =   540
            TabIndex        =   40
            Top             =   450
            Width           =   1275
         End
         Begin VB.TextBox txtSeat 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   2340
            Width           =   3435
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1890
            Width           =   3435
         End
         Begin VB.TextBox txtPayOrderDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   315
            Left            =   1575
            TabIndex        =   35
            Top             =   1440
            Width           =   3435
         End
         Begin VB.TextBox txtPaymentOrderNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   315
            Left            =   1575
            TabIndex        =   33
            Top             =   990
            Width           =   3435
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "User"
            Height          =   270
            Left            =   1080
            TabIndex        =   42
            Top             =   2790
            Width           =   390
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Seat"
            Height          =   270
            Left            =   1080
            TabIndex        =   38
            Top             =   2340
            Width           =   390
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   270
            Left            =   810
            TabIndex        =   37
            Top             =   1935
            Width           =   675
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payorder Date"
            Height          =   270
            Left            =   315
            TabIndex        =   34
            Top             =   1485
            Width           =   1230
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Payorder No."
            Height          =   270
            Left            =   405
            TabIndex        =   32
            Top             =   990
            Width           =   1125
         End
      End
      Begin VB.ComboBox cmbSeatID 
         Height          =   390
         Left            =   12105
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3420
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   5985
         TabIndex        =   8
         Top             =   3915
         Width           =   1275
      End
      Begin VB.CommandButton cmdSearchPaymentOrder 
         Caption         =   "..."
         Height          =   330
         Left            =   7020
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1080
         Width           =   375
      End
      Begin VB.ComboBox cmbForwardedSeat 
         Height          =   390
         Left            =   3795
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2925
         Width           =   3210
      End
      Begin VB.ComboBox cmbCancellationReason 
         Height          =   390
         Left            =   3795
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1530
         Width           =   3210
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   795
         Left            =   3795
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2025
         Width           =   3210
      End
      Begin VB.TextBox txtParOrderNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3795
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   3210
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   420
         Left            =   4680
         TabIndex        =   7
         Top             =   3915
         Width           =   1275
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   420
         Left            =   3375
         TabIndex        =   6
         Top             =   3915
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Remarks (if any..)"
         Height          =   270
         Left            =   1980
         TabIndex        =   16
         Top             =   2070
         Width           =   1620
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Forwarded to Seat"
         Height          =   270
         Left            =   1995
         TabIndex        =   15
         Top             =   3015
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cancellation Reason"
         Height          =   270
         Left            =   1845
         TabIndex        =   14
         Top             =   1575
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Payment order"
         Height          =   270
         Left            =   2325
         TabIndex        =   13
         Top             =   1125
         Width           =   1275
      End
   End
   Begin VB.Frame fraListofCancellations 
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   45
      TabIndex        =   10
      Top             =   765
      Width           =   12795
      Begin VB.Frame fraApprover 
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   -45
         TabIndex        =   27
         Top             =   5175
         Width           =   2760
         Begin VB.CommandButton cmdMyList 
            Caption         =   "My &List"
            Height          =   420
            Left            =   1395
            TabIndex        =   29
            Top             =   270
            Width           =   1320
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "&View"
            Height          =   420
            Left            =   90
            TabIndex        =   28
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame fraOperator 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   9630
         TabIndex        =   24
         Top             =   5175
         Width           =   3030
         Begin VB.CommandButton cmdNew 
            Caption         =   "&New"
            Height          =   420
            Left            =   225
            TabIndex        =   26
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   420
            Left            =   1530
            TabIndex        =   25
            Top             =   270
            Width           =   1275
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   5055
         Left            =   45
         TabIndex        =   11
         Top             =   90
         Width           =   12705
         _cx             =   22410
         _cy             =   8916
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
         BackColorBkg    =   -2147483626
         BackColorAlternate=   16777215
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
         FormatString    =   $"frmListCancelRequestPaymentOrder.frx":000C
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others' Request"
         Height          =   270
         Left            =   6525
         TabIndex        =   23
         Top             =   5175
         Width           =   1380
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000018&
         Height          =   195
         Left            =   6300
         TabIndex        =   22
         Top             =   5220
         Width           =   195
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelled"
         Height          =   270
         Left            =   6525
         TabIndex        =   21
         Top             =   5715
         Width           =   825
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   6300
         TabIndex        =   20
         Top             =   5760
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment order Approved"
         Height          =   270
         Left            =   6525
         TabIndex        =   19
         Top             =   5445
         Width           =   2160
      End
      Begin VB.Label Label7 
         BackColor       =   &H00008000&
         Height          =   195
         Left            =   6300
         TabIndex        =   18
         Top             =   5490
         Width           =   195
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12930
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "                      Cancellation Requests - Payorder"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   12930
   End
End
Attribute VB_Name = "frmListCancelRequestPaymentOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mEditMode As Boolean
    Private Sub cmdCancel_Click()
        fraListofCancellations.Visible = True
        fraNewCancellationRequests.Visible = False
    End Sub

    Private Sub cmdClear_Click()
        Call Clear
    End Sub

    Private Sub cmdEdit_Click()
        If vsGrid.Row > 0 Then
            If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
                Call cmdNew_Click
                mEditMode = True
                If vsGrid.TextMatrix(vsGrid.Row, 3) <> gbUserID Then
                    MsgBox "This payment order is generated by another user" & vbNewLine & "You cannot edit this", vbInformation
                    Exit Sub
                End If
                txtParOrderNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 0)
                txtParOrderNo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
                cmbCancellationReason.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
                cmbForwardedSeat.Text = vsGrid.TextMatrix(vsGrid.Row, 12)
                txtDescription.Text = vsGrid.TextMatrix(vsGrid.Row, 13)
            End If
        End If
    End Sub

    Private Sub cmdMyList_Click()
        If gbLBType = 4 And gbSeatGroupID = gbSeatGroupAccountsSuperintended Then     '   Accounts Supdt in a Corporation
            frmPaymentOrderCancellationRequest.Visible = True
            frmPaymentOrderCancellationRequest.ZOrder (0)
            frmPaymentOrderCancellationRequest.tabPaymentCancellationRequests.Tab = 0
        End If
    End Sub

    Private Sub cmdNew_Click()
        mEditMode = False
        fraListofCancellations.Visible = False
        fraNewCancellationRequests.Visible = True
        Call Clear
    End Sub

    Private Sub cmdRefresh_Click()
        Call FillGrid
    End Sub

    Private Sub cmdSearchPaymentOrder_Click()
        gbSearchID = -1
        frmSearchPaymentOrder.chkListToApprove.Visible = True
        frmSearchPaymentOrder.chkListToApprove.Enabled = True
        If gbLBType = 4 And gbSeatGroupID = gbSeatGroupAccountsClerk Then
            frmSearchPaymentOrder.chkListToApprove.value = 0
            frmSearchPaymentOrder.chkListToApprove.Enabled = False
        End If
        frmSearchPaymentOrder.Show vbModal
        If gbSearchID > 0 Then
            txtParOrderNo.Tag = gbSearchID
            txtParOrderNo.Text = gbSearchStr
            Call FetchPaymentOrder
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdSend_Click()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mArrayIn As Variant
        Dim mSQL As String
        
        Dim mPayOrderID         As Variant
        Dim mVoucherTypeID      As Integer
        Dim mPayOrderNo         As Variant
        Dim mUserID             As Variant
        Dim mCounterID          As Variant
        Dim mSeatID             As Variant
        Dim mReasonID           As Variant
        Dim mCancelDate         As Variant
        Dim mRemarks            As String
        Dim mApproveStatus      As Integer
        Dim mTinType            As Integer
        Dim mForwardedSeat      As Variant
        Dim mApproverID         As Variant
        
        '---------------------------------------------------------------------------'
        '                                Validations                                '
        '---------------------------------------------------------------------------'
        If gbSectionID <> 4 Then
            MsgBox "Accounts Section can only Apply for cancellation", vbInformation
            Exit Sub
        End If
        If gbSeatGroupID <> gbSeatGroupAccountsClerk Then
            MsgBox "An Operator can only Apply for cancellation", vbInformation
            Exit Sub
        End If
        If Trim(txtParOrderNo.Text) = "" Then
            MsgBox "Please Enter PayOrderNo", vbInformation
            txtParOrderNo.SetFocus
            Exit Sub
        End If
        If cmbCancellationReason.ListIndex < 1 Then
            MsgBox "Please select the Reason", vbInformation
            cmbCancellationReason.SetFocus
            Exit Sub
        End If
        If cmbForwardedSeat.ListIndex < 1 Then
            MsgBox "Please Refer the Forwarded Seat", vbInformation
            cmbForwardedSeat.SetFocus
            Exit Sub
        End If
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = "Select tnyStatus,intPayOrderID,tnyApproveStatus,dtCancellationDate,faPayOrder.numUserID,faPayOrder.tnyCancelled From faCancelledVouchers Right Join " & _
                " faPayOrder On faPayOrder.intPayOrderID = faCancelledVouchers.intVoucherID " & _
                " Where tnyRemoveCancel is Null And vchPayOrderNo = " & Trim(txtParOrderNo.Text)
        Rec.Open mSQL, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            txtParOrderNo.Tag = Rec!intPayOrderID
            If Rec!tnyCancelled = 1 Then
                MsgBox "This Paymentorder already cancelled", vbInformation
                Exit Sub
            End If
            If Rec!tnyStatus = 1 Then
                If gbLBType <> 4 Then                ''''' For Municipalities and Others
                    'In the case of Municipalities and Panchayats Approved Payorder's request is by an operator
                Else
                    MsgBox "Payment order already Approved," & vbNewLine & "an operator cannot request for this", vbInformation
                    Exit Sub
                End If
            End If
            If Rec!numUserID <> gbUserID Then
                MsgBox "Another User generated Payoder cannot be applied for cancellation", vbInformation
                Exit Sub
            Else
                If IsNull(Rec!tnyApproveStatus) = False Then
                    If Rec!tnyApproveStatus = 1 Then
                        MsgBox "This PaymentOrder is already cancelled", vbInformation
                        Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox "Invalid Payorder Number", vbInformation
            Exit Sub
        End If
        Rec.Close
        '---------------------------------------------------------------------------'
        '                           Validations Over                                '
        '---------------------------------------------------------------------------'
        mPayOrderID = IIf(Trim(txtParOrderNo.Tag) = "", -1, Trim(txtParOrderNo.Tag))
        mVoucherTypeID = 60
        mPayOrderNo = Trim(txtParOrderNo.Text)
        mUserID = gbUserID
        mCounterID = gbCounterID
        mSeatID = gbSeatID
        mReasonID = cmbCancellationReason.ItemData(cmbCancellationReason.ListIndex)
        mCancelDate = Null
        mRemarks = txtDescription.Text
        mApproveStatus = 0
        mTinType = 6
        mForwardedSeat = cmbSeatID.List(cmbForwardedSeat.ListIndex)
        mApproverID = Null
        
        mArrayIn = Array(mPayOrderID, _
                        mVoucherTypeID, _
                        Null, mPayOrderNo, _
                        mUserID, _
                        mCounterID, _
                        mSeatID, _
                        mReasonID, _
                        Null, Null, mRemarks, _
                        Null, Null, mApproveStatus, _
                        mTinType, _
                        mForwardedSeat, _
                        mApproverID)
        objDB.ExecuteSP "spSaveCancelledPaymentOrder", mArrayIn, , , mCnn
        Call Clear
        Call FillGrid
        If mEditMode Then
            fraListofCancellations.Visible = True
            fraNewCancellationRequests.Visible = False
        End If
    End Sub

    Private Sub cmdView_Click()
        If (gbSeatGroupID = gbSeatGroupAccountsSuperintended And gbLBType = 4) Or (gbSeatGroupID = gbSeatGroupAccountsOfficer And gbLBType = 3) Then
            If vsGrid.Row > 0 Then
                If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
                    frmViewPaymentorderCancellationRequest.PaymentOrderNo = Trim(vsGrid.TextMatrix(vsGrid.Row, 1))
                    frmViewPaymentorderCancellationRequest.UserType = 3
                    If vsGrid.TextMatrix(vsGrid.Row, 10) = 0 Then
                        frmViewPaymentorderCancellationRequest.Verified = False
                    Else
                        frmViewPaymentorderCancellationRequest.Verified = True
                    End If
                    frmViewPaymentorderCancellationRequest.Show vbModal
                    frmViewPaymentorderCancellationRequest.ZOrder (0)
                End If
            End If
        End If
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
        Call FillGrid
    End Sub

    Private Sub Form_Load()
        winXPC.InitSubClassing
        Call FormInitilise
        
        '''Case for User type in Accounts Section
        Call ApplyUserType
    End Sub

    Private Sub Clear()
        txtParOrderNo.Text = ""
        cmbCancellationReason.ListIndex = -1
        txtDescription.Text = ""
        cmbForwardedSeat.ListIndex = -1
        fraPayOrder.Visible = False
    End Sub
    
    Private Sub FormInitilise()
        PopulateList cmbCancellationReason, "SELECT vchCancelReason,intCancelID FROM  faCancelReason Order By vchCancelReason", , True, , True
        If gbLBType = 3 Then
            PopulateList cmbForwardedSeat, "SELECT chvSeatTitle,numSeatID FROM  faSeats Where intSectionID = 4 Order By chvSeatTitle", , True ' And intGroupID = " & gbSeatGroupAccountsOfficer & "
            PopulateList cmbSeatID, "SELECT numSeatID,numSeatID FROM  faSeats Where intSectionID = 4  Order By chvSeatTitle", , True ' And intGroupID = " & gbSeatGroupAccountsOfficer & "
        Else
            PopulateList cmbForwardedSeat, "SELECT chvSeatTitle,numSeatID FROM  faSeats Where intSectionID = 4 Order By chvSeatTitle", , True ' And intGroupID = " & gbSeatGroupAccountsSuperintended & " ' And intGroupID = " & gbSeatGroupAccountsSuperintended & "
            PopulateList cmbSeatID, "SELECT numSeatID,numSeatID FROM  faSeats Where intSectionID = 4 Order By chvSeatTitle", , True ' And intGroupID = " & gbSeatGroupAccountsSuperintended & "
        End If
    End Sub

    Private Sub FillGrid()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mArrayIn As Variant
        Dim mSQL As String
        Dim mCnt As Double
        vsGrid.Rows = 1
        vsGrid.Rows = 20
        mCnt = 1
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mArrayIn = Array(3, gbUserID)
        Set Rec = objDB.ExecuteSP("spGetListofCancelledPayments", mArrayIn, , , mCnn)
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                vsGrid.TextMatrix(mCnt, 0) = Rec!intVoucherID
                vsGrid.TextMatrix(mCnt, 1) = Rec!numReceiptNo
                vsGrid.TextMatrix(mCnt, 2) = Format(Rec!dtRequestDate, "dd/MMM/YYYY")
                vsGrid.TextMatrix(mCnt, 3) = Rec!intUserID
                vsGrid.TextMatrix(mCnt, 4) = Rec!vchUserName
                vsGrid.TextMatrix(mCnt, 5) = Rec!intReasonID
                vsGrid.TextMatrix(mCnt, 6) = Rec!vchCancelReason
                vsGrid.TextMatrix(mCnt, 7) = Rec!numApproverUserID
                vsGrid.TextMatrix(mCnt, 8) = Rec!ApproverName
                If Rec!SameUser = 2 Then                            'For Same User ToolTip Color
                    vsGrid.Cell(flexcpBackColor, mCnt, 1, mCnt, 13) = &H80000018
                End If
                If IsNull(Rec!dtCancellationDate) Then
                    vsGrid.TextMatrix(mCnt, 9) = ""
                Else
                    vsGrid.TextMatrix(mCnt, 9) = Format(Rec!dtCancellationDate, "dd/MMM/YYYY")
                End If
                vsGrid.TextMatrix(mCnt, 10) = Rec!Status
                If Rec!tnyStatus = 1 Then                           'For Payment Order Approved Green
                    vsGrid.Cell(flexcpBackColor, mCnt, 1, mCnt, 13) = vbGreen
                End If
                If Rec!Status = 0 Then
                    vsGrid.TextMatrix(mCnt, 11) = "Pending"
                ElseIf Rec!Status = 1 Then
                    vsGrid.TextMatrix(mCnt, 11) = "Approved 1 Level"
                Else
                    vsGrid.Cell(flexcpBackColor, mCnt, 1, mCnt, 13) = vbRed
                    vsGrid.Cell(flexcpForeColor, mCnt, 1, mCnt, 13) = vbWhite
                    vsGrid.TextMatrix(mCnt, 11) = "Cancelled"        'For Cancelled Red
                End If
                vsGrid.TextMatrix(mCnt, 12) = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                vsGrid.TextMatrix(mCnt, 13) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                vsGrid.AddItem ""
                mCnt = mCnt + 1
                Rec.MoveNext
            Wend
        End If
    End Sub

    Private Sub ApplyUserType()
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            fraOperator.Visible = True
            fraApprover.Visible = False
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer And gbLBType = 3 Then            '       Accounts Supdt
            fraOperator.Visible = False
            fraApprover.Visible = True
        ElseIf gbSeatGroupID = gbSeatGroupAccountsSuperintended And gbLBType = 4 Then            '       Accounts Supdt
            fraOperator.Visible = False
            fraApprover.Visible = True
        Else
            fraOperator.Visible = False
            fraApprover.Visible = False
        End If
    End Sub
    
    Private Sub FetchPaymentOrder()
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mSQL        As String
        If Trim(txtParOrderNo.Text) = "" Then
            fraPayOrder.Visible = False
            Exit Sub
        End If
        fraPayOrder.Visible = True
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = "Select faPayOrder.intPayOrderID,vchPayOrderNo,dtPayOrderDate,Max(numAmount) [numAmount],numSeatID,tnyStatus,dbo.fnGetUser(numUserID) [vchuserName] From faPayOrder " & _
                "Inner Join faPayOrderChild ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID " & _
                "Where vchPayOrderNo = " & Trim(txtParOrderNo.Text) & _
                " Group By faPayOrder.intPayOrderID,vchPayOrderNo,dtPayOrderDate,numSeatID,tnyStatus,numUserID"
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                txtPaymentOrderNo.Text = Rec!vchPayOrderNo
                txtPayOrderDate.Text = DdMmmYy(Rec!dtPayOrderDate)
                txtAmount.Text = Rec!numAmount
                txtSeat.Text = GetSeatName(Rec!numSeatID)
                  txtUser.Text = Rec!vchUserName
                If Rec!tnyStatus = 1 Then
                    chkApproved.value = vbChecked
                Else
                    chkApproved.value = vbUnchecked
                End If
                Rec.MoveNext
            Wend
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Function GetSeatName(mSeatID As Variant)
        Dim mCnnSeatName    As New ADODB.Connection
        Dim RecSeatName     As New ADODB.Recordset
        Dim objSeatName     As New clsDB
        Dim mSQLSeatName    As String
        
        On Error GoTo err:
        objSeatName.CreateNewConnection mCnnSeatName, enuSourceString.DBMaster
        
        mSQLSeatName = "Select * From GL_Seats"
        mSQLSeatName = mSQLSeatName + " Where numSeatID = " & mSeatID
        RecSeatName.Open mSQLSeatName, mCnnSeatName
        If Not (RecSeatName.EOF And RecSeatName.BOF) Then
            GetSeatName = IIf(IsNull(RecSeatName!chvSeatTitle), "", RecSeatName!chvSeatTitle)
        End If
        RecSeatName.Close
        Exit Function
err:
        MsgBox (Error$)
    End Function

    Private Sub txtParOrderNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
