VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCancelWebExtractVoucher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmCancelWebExtractVoucher"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   5310
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   840
      Width           =   3630
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9045
      TabIndex        =   18
      Top             =   4125
      Width           =   9105
      Begin VB.CommandButton cmdApprove 
         Caption         =   "&Approve"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1200
         TabIndex        =   29
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2985
         TabIndex        =   20
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdRequest 
         Caption         =   "Request"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4740
         TabIndex        =   19
         Top             =   45
         Width           =   1725
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   9585
         Top             =   450
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   4
         Common_Dialog   =   0   'False
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   9045
      TabIndex        =   16
      Top             =   0
      Width           =   9105
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reverse Entry for E bill"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   9690
      TabIndex        =   14
      Top             =   4095
      Width           =   9720
      Begin VB.Label lblMsgBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   45
         TabIndex        =   15
         Top             =   45
         Width           =   9600
      End
   End
   Begin VB.Frame fmeRequest 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   360
      Width           =   9690
      Begin VB.TextBox txtBillCCode 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   3330
      End
      Begin VB.CommandButton cmdSearchReason 
         Caption         =   "..."
         Height          =   315
         Left            =   4890
         TabIndex        =   6
         Top             =   1065
         Width           =   300
      End
      Begin VB.TextBox txtReason 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1065
         Width           =   3330
      End
      Begin VB.TextBox txtRemarks 
         Height          =   975
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   3330
      End
      Begin VB.TextBox txtVoucherNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   660
         Width           =   3330
      End
      Begin VB.TextBox txtSeat 
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
         Height          =   360
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1425
         Width           =   3330
      End
      Begin VB.CommandButton cmdSeat 
         Caption         =   "..."
         Height          =   315
         Left            =   4890
         TabIndex        =   1
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label lblMsg 
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
         Left            =   2760
         TabIndex        =   28
         Top             =   3360
         Width           =   4275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   5340
         TabIndex        =   25
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BillControl Code"
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
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason *"
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
         Left            =   765
         TabIndex        =   13
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lblReturnDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks *"
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
         Left            =   630
         TabIndex        =   12
         Top             =   1830
         Width           =   900
      End
      Begin VB.Label lblVoucherType 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Top             =   135
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No *"
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
         Left            =   360
         TabIndex        =   10
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded Seat *"
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
         Left            =   30
         TabIndex        =   9
         Top             =   1425
         Width           =   1500
      End
      Begin VB.Label lblDemand 
         Caption         =   "Demand No:"
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
         Left            =   45
         TabIndex        =   8
         Top             =   3510
         Width           =   915
      End
      Begin VB.Label lblDemandNo 
         Height          =   285
         Left            =   945
         TabIndex        =   7
         Top             =   3465
         Width           =   2715
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   3960
      TabIndex        =   26
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BillControl Code"
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
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1380
   End
End
Attribute VB_Name = "frmCancelWebExtractVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DispayWebExtractVoucherCancelLisstdetails(intBillCtrlCode As String)

    Dim objAcc As New clsAccounts
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mCount As Integer
    Dim VTypeId As Integer
    frmIntegratedPayments.mWebExtract = True
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
    End If
    mSql = "SELECT * FROM faReverseEntry INNER JOIN faWebExtracts ON faReverseEntry.numDemandID = faWebExtracts.intwebExtractID "
    mSql = mSql + " INNER JOIN faVouchers ON faWebExtracts.numKeyID = faVouchers.intVoucherID "
    mSql = mSql + " INNER JOIN faReasons ON faReverseEntry.intReasonID = faReasons.intReasonID INNER JOIN faSeats "
    mSql = mSql + " ON faVouchers.intLocalBodyID = faSeats.intLocalBodyID AND faReverseEntry.numForwardedSeatID = faSeats.numSeatID "
    mSql = mSql + " where faWebExtracts.numbillcontrolcode=" & intBillCtrlCode
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
       
        txtBillCCode.Text = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
        txtVoucherNo.Text = IIf(IsNull(Rec!intwebVoucherNo), "", Rec!intwebVoucherNo)
        txtVoucherNo.Tag = IIf(IsNull(Rec!intwebVoucherId), "", Rec!intwebVoucherId)
        VTypeId = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
        txtWebExtractID = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
        txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
        txtDescription.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
        txtReason.Text = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
        txtSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
        fmeRequest.Tag = IIf(IsNull(Rec!intRequestID), "", Rec!intRequestID)
    
    End If
End Sub



Public Sub DispayWebExtractVoucher(intWebExtractID As String)

    Dim objAcc As New clsAccounts
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mCount As Integer
    Dim VTypeId As Integer
    frmIntegratedPayments.mWebExtract = True
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
    End If
    
    mSql = " SELECT  *,isnull(faWebExtracts.numKeyID,0) as VrID From faWebExtracts  Inner Join faVouchers  On  faVouchers.intVoucherID=faWebExtracts.numKeyID"
    mSql = mSql + " Where faWebExtracts.intwebExtractID=" & intWebExtractID
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        
        txtBillCCode.Text = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
        txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        txtVoucherNo.Tag = IIf(IsNull(Rec!VrID), "", Rec!VrID)
        txtBillCCode.Tag = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
        VTypeId = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
        txtWebExtractID = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
        lblDemand.Caption = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
        txtRemarks.Tag = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
        txtDescription.Tag = IIf(IsNull(Rec!numKeyID), "", Rec!numKeyID)
     'vsGrid.TextMatrix(0, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
    
    End If
End Sub

Private Sub cmdApprove_Click()

    Dim mCnn        As New ADODB.Connection
    Dim objRev      As New clsReverseProcess
        objRev.VoucherID = val(txtVoucherNo.Tag)
        
        'MsgBox "This Voucher is Approved for Reverse," + vbCrLf + ""
        
        rVrID = objRev.ReverseTransaction(val(txtVoucherNo.Tag), mCnn)
            If rVrID = "" Then
                MsgBox "Transaction Failed", vbInformation
                'GoTo ErrRollBack
            End If
End Sub

    Private Sub cmdRequest_Click()
     
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim arrIn       As Variant
        Dim arrOut      As Variant
        Dim mRequestID  As Integer
        Dim mSql        As String
        Dim mReqID      As Double
        
        Dim VType       As String
        If val(fmeRequest.Tag) > 0 Then
            mReqID = fmeRequest.Tag
        Else
            mReqID = -1
        End If
        If val(txtSeat.Tag) < 0 Then
         MsgBox "Please select Seat", vbInformation
         Exit Sub
         End If
       
        If txtReason.Text = "" Then
            MsgBox "Please select a Reason ", vbInformation
            Exit Sub
        End If
        
        
        If objdb.SetConnection(mCnn) Then
                arrIn = Array(mReqID, _
                            gbTransactionDate, _
                            80, _
                            val(txtRemarks.Tag), _
                            val(txtReason.Tag), _
                            Trim(txtRemarks.Text), _
                            gbUserID, _
                            gbSeatID, _
                            Null, _
                            Null, _
                            txtSeat.Tag, _
                            gbFinancialYearID, _
                            0, _
                            Null, _
                            Null, _
                            Null, _
                            txtBillCCode.Tag, _
                             val(txtVoucherNo.Tag) _
                            )
                objdb.ExecuteSP "spSaveReverseEntry", arrIn, arrOut, , mCnn, adCmdStoredProc
                If Not IsNumeric(arrOut) Then
                    mRequestID = arrOut(0, 0)
                End If
                MsgBox "Successfully Saved", vbInformation
                cmdRequest.Enabled = False
          End If
    End Sub

Private Sub cmdSearchReason_Click()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objdb As New clsDB
            frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=80"
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Show vbModal
            txtReason.Text = gbSearchStr
            txtReason.Tag = gbSearchID
            gbSearchID = -1
            gbSearchStr = ""
End Sub

Private Sub cmdSeat_Click()
     Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mCnt    As Integer
        Dim mSql    As String
'        vsSeat.Visible = True
        txtSeat.Tag = ""
        txtSeat.Text = ""
        mSql = "Select chvSeatTitle, numSeatID From GL_Seats Where intGroupID in (5,6) And intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        frmSearchSeat.SQLString = mSql
        frmSearchSeat.Show vbModal
        If gbSearchID > -1 Then
            txtSeat.Tag = gbSearchID
            txtSeat.Text = gbSearchStr
        Else
            gbSearchID = -1
            gbSearchStr = ""
        End If
End Sub

Private Sub cmdVerify_Click()
     Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mCnt    As Integer
        Dim mSql    As String
        Dim mReqID      As Double
        If val(fmeRequest.Tag) > 0 Then
            mReqID = fmeRequest.Tag
            End If
        If objdb.SetConnection(mCnn) Then
            mSql = "Update faReverseEntry set tnyStatus=1 where intRequestID=" & mReqID
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            MsgBox "Verified Successfully", vbApplicationModal
            cmdVerify.Enabled = False
            cmdApprove.Enabled = True
        End If
       
End Sub


Private Sub Form_Load()
     If gbSeatGroupID = gbSeatGroupAccountsClerk Then
        cmdRequest.Enabled = True
        cmdVerify.Enabled = False
        cmdApprove.Enabled = False
     ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
        cmdRequest.Enabled = False
        cmdVerify.Enabled = True
   
'    ElseIf gbSeatGroupID = gbSeatGroupSecretary Then
'        cmdRequest.Enabled = True
'        cmdVerify.Enabled = False
'        cmdApprove.Enabled = False
'    Else
'        cmdVerify.Visible = False
    End If

End Sub

