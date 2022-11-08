VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReverseApproval 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmReverseApproval"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12765
   Icon            =   "frmReverseApproval.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   90
      TabIndex        =   25
      Top             =   1575
      Width           =   12570
      Begin VB.Label lblRemarks 
         Height          =   285
         Left            =   45
         TabIndex        =   26
         Top             =   180
         Width           =   12030
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   12705
      TabIndex        =   17
      Top             =   7815
      Width           =   12765
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify Modified Voucher"
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
         Left            =   2565
         TabIndex        =   24
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
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
         Left            =   7830
         TabIndex        =   21
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdApprove 
         Caption         =   "&Approve"
         Enabled         =   0   'False
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
         Left            =   4320
         TabIndex        =   20
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdReject 
         Caption         =   "Reject"
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
         Left            =   6075
         TabIndex        =   19
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdOld 
         Caption         =   "Requested Voucher"
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
         Left            =   11025
         TabIndex        =   18
         Top             =   45
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblTrType 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   150
         TabIndex        =   33
         Top             =   210
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   12705
      TabIndex        =   15
      Top             =   0
      Width           =   12765
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reverse Entry Report"
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
         Left            =   45
         TabIndex        =   16
         Top             =   45
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Approved By"
      Height          =   1230
      Left            =   8550
      TabIndex        =   10
      Top             =   360
      Width           =   4110
      Begin VB.TextBox txtApproveDate 
         Height          =   345
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   900
         Width           =   3045
      End
      Begin VB.TextBox txtApproverSeat 
         Height          =   345
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   3045
      End
      Begin VB.TextBox txtApprover 
         Height          =   345
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   3045
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   32
         Top             =   945
         Width           =   405
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   675
         Width           =   420
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Requested By"
      Height          =   1230
      Left            =   4140
      TabIndex        =   5
      Top             =   360
      Width           =   4290
      Begin VB.TextBox txtReqDate 
         Height          =   345
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   900
         Width           =   3045
      End
      Begin VB.TextBox txtUser 
         Height          =   345
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   3045
      End
      Begin VB.TextBox txtSeat 
         Height          =   345
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   3045
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Req Date"
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
         Left            =   15
         TabIndex        =   31
         Top             =   945
         Width           =   795
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         Height          =   195
         Left            =   405
         TabIndex        =   9
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         Height          =   195
         Left            =   405
         TabIndex        =   8
         Top             =   675
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.TextBox txtVrDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   855
         Width           =   2640
      End
      Begin VB.TextBox txtVoucherNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   135
         Width           =   2640
      End
      Begin VB.TextBox txtReason 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   495
         Width           =   2640
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vr Date"
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
         Left            =   345
         TabIndex        =   30
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   270
         Left            =   45
         TabIndex        =   4
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
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
         Left            =   405
         TabIndex        =   3
         Top             =   540
         Width           =   630
      End
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   5670
      Left            =   45
      TabIndex        =   22
      Top             =   2160
      Width           =   12630
      lastProp        =   500
      _cx             =   22278
      _cy             =   10001
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReportOld 
      Height          =   5670
      Left            =   5850
      TabIndex        =   23
      Top             =   2250
      Width           =   6555
      lastProp        =   500
      _cx             =   11562
      _cy             =   10001
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   13095
      Top             =   7875
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      Common_Dialog   =   0   'False
   End
End
Attribute VB_Name = "frmReverseApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private intVerify       As Integer  ' 0 = not verified;   1 = verified        '
    Dim mRequestID          As Integer  ' Set mRequest=mArray(3)
    Dim mDemandNo           As Variant
    Private mKeyID          As Integer
    Private mCategory       As Integer
    Private mVrType         As Integer
    Private mTrType         As Integer
    Private mDemandTotal    As Double
    Private mReceiptVrID    As Double  'Set From Receipt Counter form for Reason category =3 (Particulars)
    Private mPreviousYearID As Integer
    Private mPreviousYearRequestID As Integer
    Private mdtTrDate       As Date
    
    
''    Private strFormName As String
''        Dim mvarCategoryID  As Integer
''        Dim intInstrumentTypeID As Variant
''        Dim intMultipleVouchers As Integer

''        Private intUserType As Integer  '       0 = operator;       1 = approver        '
''        Private intRequestID As Long
''        Dim mDemandNo    As Variant
''        '---------------------------------
''        'To vrify demand is Saved or not 1=Saved   set this variable from DemandInterface Form
''        Public mRevDemand         As Boolean
''        '---------------------------------
''
''
''
''    Private Sub cmdChequeReturn_Click()
'''        Call FormInitialize
''        cmdRequest.Enabled = True
''        frmChequeBounceRequest.Show vbModal
''    End Sub
'''''''    Private Sub FillDemand(ByVal mVoucherID As Double)
'''''''        Dim objDb       As New clsDB
'''''''        Dim Rec         As New ADODB.Recordset
'''''''        Dim RecChild    As New ADODB.Recordset
'''''''        Dim mCnn        As New ADODB.Connection
'''''''        Dim mSql        As String
'''''''        Dim mRowCount   As Integer
'''''''        Dim mPeriodID   As Integer
'''''''        Dim mYearID     As Integer
'''''''        Dim mArrearFlag As Integer
'''''''
'''''''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'''''''        mSql = "Select * From faVouchers"
'''''''        mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
'''''''        mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
'''''''        mSql = mSql + " Inner Join faTransactions On faTransactions.intVoucherID=faVouchers.intVoucherID"
'''''''        mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'''''''        mSql = mSql + " Inner Join faSection On faTransactionType.intSectionID=faSection.intSectionID"
'''''''        mSql = mSql + " Inner Join faInstrumentTypes On faVouchers.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
'''''''        mSql = mSql + " Inner Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
'''''''        mSql = mSql + " Left Join faFunctionaries On faFunctionaries.intFunctionaryID=faTransactions.intFunctionaryID"
'''''''        mSql = mSql + " Left Join faFunctions On faFunctions.intFunctionID=faTransactions.intFunctionID"
'''''''        mSql = mSql + " Left Join faFunds On faFunds.intFundID=faTransactions.intFundID"
'''''''        mSql = mSql + " Left Join DB_Masters..GM_Zone On faVouchers.numZoneID=DB_Masters..GM_Zone.numZoneID"
'''''''        mSql = mSql + " Where faVouchers.intVoucherID=" & mVoucherID
'''''''        Rec.Open mSql, mCnn
'''''''        If Not (Rec.EOF And Rec.BOF) Then
'''''''            With frmDemandInterface
'''''''                .cmbSections.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
'''''''                .cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
'''''''                If .cmbTransactionType.ItemData(.cmbTransactionType.ListIndex) = gbTransactionTypeOutDoor Then
'''''''                    .cmbOutDoorStaff.Enabled = True
'''''''                    If IsNull(Rec!vchUserName) = False Then
'''''''                        .cmbOutDoorStaff.Text = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
'''''''                    End If
'''''''                End If
''''''''                If IsNull(Rec!chvZoneNameEnglish) = False Then
''''''''                    .cmbOutDoorStaff.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
''''''''                End If
'''''''                .txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
'''''''                .txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
'''''''                .txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
'''''''                If IsNull(Rec!vchInstrumentType) = False Then
'''''''                    .cmbInstrumentType.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
'''''''                    .cmbInstrumentType.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
'''''''                End If
'''''''                If .cmbInstrumentType.Tag <> "" Then
'''''''                    If .cmbInstrumentType.Tag <> 1 Then
'''''''                        .lblInstNo.Visible = True
'''''''                        .txtInstrumentNo.Visible = True
'''''''                        .txtInstrumentNo.Enabled = True
'''''''                        .lblInstDate.Visible = True
'''''''                        .txtInstrumentDate.Visible = True
'''''''                        .txtInstrumentDate.Enabled = True
'''''''                        .lblDrawnFrom.Visible = True
'''''''                        .txtDrawnFrom.Visible = True
'''''''                        .txtDrawnFrom.Enabled = True
'''''''                        .lblDrawnPlace.Visible = True
'''''''                        .txtDrawnPlace.Visible = True
'''''''                        .txtDrawnPlace.Enabled = True
'''''''                        .txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''''                        .txtInstrumentDate.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
'''''''                        .txtDrawnFrom.Text = IIf(IsNull(Rec!vchDrawnFrom), "", Rec!vchDrawnFrom)
'''''''                        .txtDrawnPlace.Text = IIf(IsNull(Rec!vchDrawnPlace), "", Rec!vchDrawnPlace)
'''''''                    End If
'''''''                End If
'''''''                .txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
'''''''                .txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
'''''''                .txtFunctionary.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
'''''''                .txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
'''''''                .txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
'''''''                .txtInitial1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
'''''''                .txtInitial2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
'''''''                .txtInitial3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
'''''''                .txtInitial4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
'''''''                .txtHouseName.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'''''''                '.txtStreet.Text = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
'''''''                .txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
'''''''                .txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'''''''                '.txtPost.Text = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
'''''''                '.txtPin.Text = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
'''''''                .txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
'''''''
'''''''                .txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
'''''''                .txtAdminNote.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
''''''''                If IsNull(Rec!numForwardedSeatID) = False Then
''''''''                    .cmbSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
''''''''                End If
'''''''                mSql = ""
'''''''                mSql = "Select * From faVoucherChild"
'''''''                mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
'''''''                mSql = mSql + " Where intVoucherID=" & mVoucherID
'''''''                RecChild.Open mSql, mCnn
'''''''                mRowCount = 1
'''''''                While Not Rec.EOF
'''''''                    While Not RecChild.EOF
'''''''                        .vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)
'''''''                        .vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecChild!vchAccountHead), "", RecChild!vchAccountHead)
'''''''                        mPeriodID = IIf(IsNull(RecChild!tnyPeriodID), "", RecChild!tnyPeriodID)
'''''''                        mYearID = IIf(IsNull(RecChild!intYearID), 0, RecChild!intYearID)
'''''''                        If mYearID <> 0 Then
'''''''                            .vsGrid.TextMatrix(mRowCount, 2) = mYearID & "-" & mYearID + 1
'''''''                        End If
'''''''                        If mPeriodID = 1 Then
'''''''                            .vsGrid.TextMatrix(mRowCount, 3) = "1st Half"
'''''''                        End If
'''''''                        If mPeriodID = 2 Then
'''''''                            .vsGrid.TextMatrix(mRowCount, 3) = "2nd Half"
'''''''                        End If
'''''''                        If mPeriodID = 3 Then
'''''''                            .vsGrid.TextMatrix(mRowCount, 3) = "Full Year"
'''''''                        End If
'''''''                        mArrearFlag = IIf(IsNull(RecChild!tnyArrearFlag), "", RecChild!tnyArrearFlag)
'''''''                        If mArrearFlag = 0 Then
'''''''                            .vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount)
'''''''                        End If
'''''''                        If mArrearFlag = 1 Then
'''''''                            .vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount)
'''''''                        End If
'''''''                        .vsGrid.Rows = .vsGrid.Rows + 1
'''''''                        mRowCount = mRowCount + 1
'''''''                        RecChild.MoveNext
'''''''                    Wend
'''''''                    Rec.MoveNext
'''''''                Wend
'''''''                RecChild.Close
'''''''
'''''''            End With
'''''''        End If
'''''''    End Sub
'    Private Sub cmdApprove_Click()
'            Dim objDb       As New clsDB
'            Dim objReverse  As New clsReverseProcess
'            Dim Rec         As New ADODB.Recordset
'            Dim mCnn        As New ADODB.Connection
'            Dim mSql        As String
'            Dim arrOut      As Variant
'            Dim mCnt        As Integer
'            Dim mFlag       As Boolean
'            Dim mVoucher    As Double
'            arrOut = objReverse.ReverseProcess(intRequestID, val(txtReason.Tag))
'            If IsNull(arrOut) Then
'                lblMsgBox.Visible = True
'                lblMsgBox.Caption = "Reverse Entry Process Failed for Voucher No " & txtVoucherNo.Text
'            Else
'                For mCnt = 0 To UBound(arrOut) - 1
'                    If UBound(arrOut) > 1 Then
'                        mFlag = True
'
'''                    Else
'''                         If lblVoucherType.Caption = "Receipt Voucher" Then
'''                            If (txtReason.Tag) > 2 Then
'''
'''                            End If
'                         End If
'    '                    lstVouchers.AddItem arrOut(mCnt)
'    '                    lstVouchers.ItemData(lstVouchers.NewIndex) = mCnt
'                    End If
'                Next
'''                If mFlag Then
'''
'''                    MsgBox "Multiple Vouchers Reversed : Please Verify"
'''
'''                Else
'''                    mVoucher = arrOut(0) 'lstVouchers.ItemData(lstVouchers.ListIndex)
'''                    'Call VoucherDetails(mVoucher)
'''                End If
'
'''                lblMsgBox.Visible = True
'''                lblMsgBox.Caption = "Reverse Entry Approved"
'                MsgBox "Reverse Entry Approved"
'                objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'                mSql = "Update faReverseEntry set  numApprovedUserID=" & gbUserID & ",dtApprovedDate='" & gbTransactionDate & "' Where intRequestID=" & intRequestID
'                mCnn.Execute mSql
'            End If
'    End Sub
''
''    Private Function RequsetValidation() As Boolean
''        If txtReason.Text = "" Then
''            lblMsgBox.Visible = True
''            lblMsgBox.Caption = "Please Select the Reason for Reverse Entry"
''            cmdSearchReason.SetFocus
''            RequsetValidation = False
''            Exit Function
''        End If
''        If txtRemarks.Text = "" Then
''            lblMsgBox.Visible = True
''            lblMsgBox.Caption = "Please give Remarks / Narration"
''            txtRemarks.SetFocus
''            RequsetValidation = False
''            Exit Function
''        End If
''        If txtSeat.Tag = "" Then
''            lblMsgBox.Visible = True
''            lblMsgBox.Caption = "Please Select the Seat"
''            cmdSeat.SetFocus
''            RequsetValidation = False
''            Exit Function
''        End If
''
''        RequsetValidation = True
''    End Function
''    Private Sub cmdRequest_Click()
''        Dim mCrl As Control
''            If RequsetValidation = False Then Exit Sub
''            If lblVoucherType.Tag = 10 Then
''                Select Case val(txtReason.Tag)
''                    Case 1, 2, 3
''                        Call SaveRequest
''                    Case 4  'Amount
''                        With frmDemandInterface
''                            .Reverse = 1
''                            .ReverseDemandDetails (val(txtVoucherNo.Tag))
''                            On Error Resume Next
''                            For Each mCrl In frmDemandInterface.Controls
''                                If TypeOf mCrl Is ComboBox Then
''                                    mCrl.Enabled = False
''                                ElseIf TypeOf mCrl Is CommandButton Then
''                                        mCrl.Enabled = False
''                                ElseIf TypeOf mCrl Is TextBox Then
''                                        mCrl.Enabled = False
''                                End If
''                            Next
'''                            .vsGrid.EditMask
''                            .cmdSave.Enabled = True
''                            .cmdCancel.Enabled = True
''                            .cmbTransactionType.Enabled = False
''                            .cmbInstrumentType.Enabled = False
''                            .cmbOutDoorStaff.Enabled = False
''                            .Show vbModal
''                        End With
''                        If mRevDemand = True Then
''                            Call SaveRequest
''                        Else
''                            MsgBox "Request Failed"
''                            Exit Sub
''                        End If
''                    Case 5  'Account Head
''                    Case 6  'Transaction Type
''                    Case 7  'Wrong demand
''                End Select
''            Else
''                Call SaveRequest
''            End If
''
''    End Sub
'
'
'    Private Sub ReportView()
'         Dim Rpt As New CRAXDRT.Report
'         Dim mApp As New CRAXDRT.Application
'         Dim rptFileName As String
'         Dim arrInput As Variant
'         Dim mLoop As Long
'
'         If FormName = "frmReverseRequest" Then
'            If MultipleVouchers Then
'               rptFileName = App.Path & "\Reports\rptMultipleVoucher.rpt"
'               crvReport.DisplayToolbar = True
'               crvReport.EnableNavigationControls = True
'               crvReport.EnableToolbar = True
'            Else
'               rptFileName = App.Path & "\Reports\rptVoucher.rpt"
'            End If
'         End If
'
''         If FormName = "frmSubsidiaryCashBook" Then
''            rptFileName = App.Path & "\Reports\rptVoucher.rpt"
''         End If
''
''          If FormName = "frmInterruptReceipt" Then
''            rptFileName = App.Path & "\Reports\rptVoucher.rpt"
''         End If
''
''         If FormName = "frmViewPaymentOrder" Then
''            rptFileName = App.Path & "\Reports\rptPOjournals.rpt"
''            crvReport.DisplayToolbar = True
''            crvReport.EnableNavigationControls = True
''            crvReport.EnableToolbar = True
''         End If
''         If FormName = "PrintPaymentOrder" Then
''            rptFileName = App.Path & "\Reports\rptPaymentOrder.rpt"
''            crvReport.DisplayToolbar = True
''            crvReport.EnableToolbar = True
''            cmdPrint.Visible = True
''            cmdVerify.Visible = False
''         End If
'         arrInput = ArrayIn
'         Screen.MousePointer = vbHourglass
'         crvReport.DisplayTabs = True
'
'         Set Rpt = Nothing
'         mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
'         Set Rpt = mApp.OpenReport(rptFileName, 1)
'
'         If IsArray(arrInput) Then
'             For mLoop = LBound(arrInput) To UBound(arrInput)
'                 Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
'                 Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
'             Next mLoop
'         End If
'         Screen.MousePointer = vbDefault
'         crvReport.ReportSource = Rpt
'         crvReport.ViewReport
'         crvReport.Zoom (1)
'    End Sub
'
'    Private Sub cmdClose_Click()
'        Unload Me
'    End Sub
'
'    Private Sub Form_Load()
'        Call ReportView
'        If FormName = "frmReverseRequest" Then
'            cmdApprove.Visible = True
'        Else
'            cmdApprove.Visible = False
'        End If
'    End Sub
''    Private Sub SaveRequest()
''        Dim objDb       As New clsDB
''        Dim Rec         As New ADODB.Recordset
''        Dim mCnn        As New ADODB.Connection
''        Dim arrIn       As Variant
''        Dim arrOut      As Variant
''        Dim mRequestID  As Integer
''        Dim mSql        As String
''
''
''        If objDb.SetConnection(mCnn) Then
''                        arrIn = Array(-1, _
''                                    gbTransactionDate, _
''                                    Null, _
''                                    lblVoucherType.Tag, _
''                                    val(txtReason.Tag), _
''                                    Trim(txtRemarks.Text), _
''                                    gbUserID, _
''                                    gbSeatID, _
''                                    Null, _
''                                    Null, _
''                                    txtSeat.Tag, _
''                                    gbFinancialYearID, _
''                                    0, _
''                                    Null, _
''                                    Null, _
''                                    mDemandNo)
''                        objDb.ExecuteSP "spSaveReverseEntry", arrIn, arrOut, , mCnn, adCmdStoredProc
''                        If Not IsNumeric(arrOut) Then
''                            mRequestID = arrOut(0, 0)
''                        End If
''                        arrIn = ""
''                        arrIn = Array(mRequestID, val(txtVoucherNo.Tag))
''                        objDb.ExecuteSP "spSaveReverseEntryChild", arrIn, , , mCnn, adCmdStoredProc
''                        lblMsgBox.Visible = True
''                        lblMsgBox.Caption = "Reverse Entry requested to Higher Authority"
''                        cmdRequest.Enabled = False
''    '                    FillDemand (val(txtVoucherNo.Tag))
''                        Unload Me
''                    Else
''                        MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
''                    End If
''    End Sub
''    Private Sub cmdSearchReason_Click()
''        On Error GoTo Err:
''            If txtVoucherNo.Text = "" Then
''                lblMsgBox.Visible = True
''                lblMsgBox.Caption = "Please Select Voucher Before Giving Reason"
''                Exit Sub
''
''            End If
''            lblMsgBox.Visible = False
''            If intInstrumentTypeID = gbInstrumentCash Then
''                frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReverseReasons Where intReasonID <>1 and intReasonID <>2"
''            Else
''                frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReverseReasons Where intReasonID <>1"
''            End If
''            frmSearchMasters.Connection = enuSourceString.Saankhya
''            frmSearchMasters.QrySP = Qyery
''            frmSearchMasters.Show vbModal
''            txtReason.Text = gbSearchStr
''            txtReason.Tag = gbSearchID
''
''            gbSearchID = -1
''            gbSearchStr = ""
''        Exit Sub
''Err:
''        MsgBox (Error$)
''    End Sub
''
''    Private Sub cmdSearchVoucher_Click()
''        Dim mSql    As String
''        Dim Rec     As New ADODB.Recordset
''        Dim mCnn    As New ADODB.Connection
''        Dim objDb   As New clsDB
''        lblMsgBox.Visible = False
''        frmSearchVouchers.Show vbModal
''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
''        If gbSearchID <> -1 Then
''
''            mSql = "Select tnyVoucherTypeID,isNull(intInstrumentTypeID,100) intInstrumentTypeID From faVouchers Where intVoucherID=" & gbSearchID
''            Rec.Open mSql, mCnn
''            If Not Rec.EOF Then
''                If Rec!tnyVoucherTypeID = 20 Then
''                    lblMsgBox.Visible = True
''                    lblMsgBox.Caption = "Payment Voucher is Not Allowed To Reverse"
''                    gbSearchCode = ""
''                    gbSearchStr = ""
''                    gbSearchID = -1
''                    Exit Sub
''                ElseIf Rec!tnyVoucherTypeID = 30 Then
''                        If CheckReverseRequestExist(gbSearchID) = 1 Then
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "Already sent Request for this Voucher"
''                            cmmVerify.Enabled = False
''                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
''                            Call GetVoucherDetails(gbSearchID)
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "This Voucher Already Reversed"
''                            cmmVerify.Enabled = False
''                        Else
''                            Call GetVoucherDetails(gbSearchID)
''                            cmmVerify.Enabled = True
''                        End If
''                ElseIf Rec!tnyVoucherTypeID = 40 Then
''                        If CheckReverseRequestExist(gbSearchID) = 1 Then
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "Already sent Request for this Voucher"
''                            cmmVerify.Enabled = False
''                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
''                            Call GetVoucherDetails(gbSearchID)
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "This Voucher Already Reversed"
''                            cmmVerify.Enabled = False
''                        Else
''                            Call GetVoucherDetails(gbSearchID)
''                            cmmVerify.Enabled = True
''                        End If
''                ElseIf Rec!tnyVoucherTypeID = 10 Then
''                    If Rec!intInstrumentTypeID = 5 Then
''                        If CheckReverseRequestExist(gbSearchID) = 1 Then
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "Already sent Request for this Voucher"
''                            cmmVerify.Enabled = False
''                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
''                            Call GetVoucherDetails(gbSearchID)
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "This Voucher Already Reversed"
''                            cmmVerify.Enabled = False
''                        Else
''                            Call GetVoucherDetails(gbSearchID)
''                            cmmVerify.Enabled = True
''                        End If
''                    Else
''                        If CheckReverseRequestExist(gbSearchID) = 1 Then
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "Already sent Request for this Voucher"
''                            cmmVerify.Enabled = False
''                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
''                            Call GetVoucherDetails(gbSearchID)
''                            lblMsgBox.Visible = True
''                            lblMsgBox.Caption = "This Voucher Already Reversed"
''                            cmmVerify.Enabled = False
''                        Else
''                            Call GetVoucherDetails(gbSearchID)
''                            cmmVerify.Enabled = True
''                        End If
'''''                        lblMsgBox.Visible = True
'''''                        lblMsgBox.Caption = "Receipt Voucher Other Than Cheque Instrument is Not Allowed To Reverse"
'''''                        gbSearchCode = ""
'''''                        gbSearchStr = ""
'''''                        gbSearchID = -1
'''''                        Exit Sub
''                    End If
''                End If
''            End If
''        End If
''        gbSearchCode = ""
''        gbSearchStr = ""
''        gbSearchID = -1
''    End Sub
''
''    Private Function CheckReverseRequestExist(ByVal VchID As Double) As Integer
''        On Error GoTo Err:
''            Dim mCnn As New ADODB.Connection
''            Dim Rec As New ADODB.Recordset
''            Dim mSql As String
''            Dim objDb As New clsDB
''            If objDb.SetConnection(mCnn) Then
''                mSql = " Select tnyStatus from faReverseEntry "
''                mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
''                mSql = mSql + " Where intVoucherID =  " & VchID
''                mSql = mSql + " And tnyStatus<>4"
''                Rec.Open mSql, mCnn
''                If Not (Rec.EOF Or Rec.BOF) Then
''                    If Rec!tnyStatus = 0 Then      'Requested
''                        CheckReverseRequestExist = 1
''                    ElseIf Rec!tnyStatus = 1 Then  ' Approved
''                        CheckReverseRequestExist = 2
''                    Else                           'Cancelled Status=4
''                        CheckReverseRequestExist = 3
''                    End If
''                    Exit Function
''                End If
''            Else
''                MsgBox "Connection to Finance does not Exist, Please Contact your System Administrator"
''            End If
''        Exit Function
''Err:
''        MsgBox (Error$)
''    End Function
''
''    Public Function GetVoucherDetails(ByVal intVoucherID As Long) As Boolean
''        On Error GoTo Err:
''            Dim mSql As String
''            Dim Rec As New ADODB.Recordset
''            Dim mCnn As New ADODB.Connection
''            Dim objDb As New clsDB
''
''            If objDb.SetConnection(mCnn) Then
''                mSql = "Select * from faVouchers Where intVoucherID = " & intVoucherID
''                Rec.Open mSql, mCnn
''                If Not (Rec.EOF Or Rec.BOF) Then
''                    Select Case Rec!tnyVoucherTypeID
''                        Case 10:
''                            lblVoucherType.Caption = "Receipt Voucher"
''                        Case 20:
''                            lblVoucherType.Caption = "Payment Voucher"
''                        Case 30:
''                            lblVoucherType.Caption = "Contra Voucher"
''                        Case 40:
''                            lblVoucherType.Caption = "Journal Voucher"
''                    End Select
''                    lblVoucherType.Tag = Rec!tnyVoucherTypeID
''                    txtVoucherNo.Tag = Rec!intVoucherID
''                    txtVoucherNo.Text = Rec!intVoucherNo
''                    lblNetAmount.Caption = Rec!fltAmount
''                    intInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
''                End If
''                If Rec.State = 1 Then Rec.Close
''            Else
''                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
''            End If
''        Exit Function
''Err:
''        MsgBox (Error$)
''    End Function
''
''    Private Sub cmdSeat_Click()
''        frmSearchSeat.Show vbModal
''        If gbSearchID = -1 Then
''            Exit Sub
''        Else
''            txtSeat.Text = gbSearchStr
''            txtSeat.Tag = gbSearchID
''        End If
''    End Sub
''
'''''    Private Sub cmdVoucherSearch_Click()
'''''        Dim mVoucherID      As Double
'''''        SelectionMode = 2
'''''        Call FormInitialize
'''''        cmdRequest.Enabled = True
'''''        frmSearchTransactions.FormSelectionType = 2
'''''        frmSearchTransactions.Show vbModal
'''''        mVoucherID = gbSearchID
'''''        frmSearchTransactions.FormSelectionType = -1
'''''        Call FillSearchData(mVoucherID)
'''''    End Sub
''
''    Private Sub cmmVerify_Click()
''        On Error GoTo Err:
''            If intMultipleVouchers > 1 Then
''                frmViewVoucher.MultipleVouchers = True
''                frmViewVoucher.ArrayIn = Array(RequestID)
''            Else
''
''                If txtVoucherNo.Tag = "" Then
''                    lblMsgBox.Visible = True
''                    lblMsgBox.Caption = "Please Select A Voucher to do Verification"
''                    Exit Sub
''                End If
''
''                frmViewVoucher.MultipleVouchers = False
''                frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
''            End If
'''            If gbSeatGroupID <> gbSeatGroupAccountsClerk Then 'UserType = 1 Then
'''                frmViewVoucher.FormName = "frmReverseEntryRequest"
'''                frmViewVoucher.Show vbModal
'''                If VerifyStatus = 1 Then
'''                    cmmVerify.Enabled = False
'''                    lblMsgBox.Visible = True
'''                    lblMsgBox.Caption = "Please Click Request Button to Approve Reverse Entry Request"
'''                Else
'''                   ' cmdApprove.Enabled = False
'''                    cmmVerify.Enabled = True
'''                    lblMsgBox.Visible = True
'''                    lblMsgBox.Caption = "Voucher Verification Failed"
'''                End If
''
'''            Else
''                frmViewVoucher.FormName = "frmReverseEntryRequest"
''                frmViewVoucher.Show vbModal
''                If VerifyStatus = 1 Then
''                    lblMsgBox.Visible = True
''                    lblMsgBox.Caption = "Please Click Request Button to Send Reverse Entry Request"
''                    cmdRequest.Enabled = True
''                    cmmVerify.Enabled = False
''                Else
''                    lblMsgBox.Visible = True
''                    lblMsgBox.Caption = "Voucher Verification Failed"
''                    cmdRequest.Enabled = False
''                    cmmVerify.Enabled = True
''                End If
'''            End If
''
''            If intMultipleVouchers > 1 Then
''                frmRequisition.Enabled = False
''            Else
''                frmRequisition.Enabled = True
''            End If
''
''        Exit Sub
''Err:
''        MsgBox (Error$)
''    End Sub
''
''    Private Sub Form_Load()
''        On Error GoTo Err:
''            intMultipleVouchers = 0
''            If UserType = 1 Then
''                'cmdApprove.Enabled = True
''                cmdRequest.Enabled = False
''                Call FillData(RequestID)
''                Call FormLock(True)
'''                cmdApprove.Enabled = False
''            ElseIf UserType = 0 Then
'''                cmdApprove.Enabled = False
''            ElseIf UserType = 2 Then
'''                cmdApprove.Enabled = False
''                cmdRequest.Enabled = False
''                lblMsgBox.Visible = True
''                lblMsgBox.Caption = "Reverse Entry Process Completed for this request"
''                Call FillData(RequestID)
''            End If
''            If intMultipleVouchers > 1 Then
''                lblMsgBox.Visible = True
''                lblMsgBox.Caption = "There are Multiple Vouchers in this Request"
'''                cmdApprove.Enabled = False
''                txtVoucherNo.Text = "Multiple Vouchers"
''                lblNetAmount.Caption = Format(frmListReverseEntryRequests.vsGridForCheque.TextMatrix(frmListReverseEntryRequests.VSGrid.Row, 1), "0.00")
''            End If
''
''        Exit Sub
''Err:
''        MsgBox (Error$)
''    End Sub
''    Private Sub FormLock(mFlag As Boolean)
''        'mflag =true for lock or disable
''        cmdSearchVoucher.Enabled = Not mFlag
''        txtVoucherNo.Locked = mFlag
''        cmdSearchReason.Enabled = Not mFlag
''        txtReason.Locked = mFlag
''        txtRemarks.Locked = mFlag
''        cmdSeat.Enabled = Not mFlag
''        txtSeat.Locked = mFlag
''    End Sub
''
''    Private Function FillData(ByVal intRequestID As Long) As Boolean
''         On Error GoTo Err:
''            Dim mCnn As New ADODB.Connection
''            Dim Rec As New ADODB.Recordset
''            Dim objDb As New clsDB
''            Dim mSql As String
''
''            If objDb.SetConnection(mCnn) Then
''                mSql = "Select * from faReverseEntry "
''                mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID "
''                mSql = mSql + " Inner Join faReverseReasons On faReverseReasons.intReasonID = faReverseEntry.intReasonID "
''                mSql = mSql + " Left JOIN faSeats ON faReverseEntry.numForwardedSeatID = faSeats.numSeatID "
''                mSql = mSql + " Where faReverseEntry.intRequestID = " & intRequestID
''
''                Rec.Open mSql, mCnn
''                While Not (Rec.EOF Or Rec.BOF)
''                    txtReason.Text = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
''                    txtReason.Tag = IIf(IsNull(Rec!intReasonID), "", Rec!intReasonID)
''                    txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
''                    txtSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
''                    txtSeat.Tag = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
''
''                    intMultipleVouchers = intMultipleVouchers + 1 '  To Count the No of Multiple Vouchers
''                    Call GetVoucherDetails(Rec!intVoucherID)
''                    Rec.MoveNext
''                Wend
''                If Rec.State = 1 Then Rec.Close
''            Else
''                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
''            End If
''        Exit Function
''Err:
''        MsgBox (Error$)
''    End Function
''
''
''    Private Function ReverEntry() As Boolean
''        On Error GoTo Err:
''            Dim objDb As New clsDB
''            Dim Rec As New ADODB.Recordset
''        Exit Function
''Err:
''        MsgBox (Error$)
''    End Function
''
''
''Private Sub Label17_Click()
''
''End Sub
''
''    Private Sub txtRemarks_LostFocus()
''        lblMsgBox.Visible = False
''    End Sub
''
''    Private Sub txtSeat_KeyPress(KeyAscii As Integer)
''       Call KeyPress(KeyAscii)
''    End Sub
''    Private Sub KeyPress(KeyAscii As Integer)
''        If KeyAscii = 13 Then
''            PressTabKey
''        Else
''            KeyAscii = 0
''        End If
''    End Sub
''

''
''    Public Property Let UserType(mData As Integer)
''        intUserType = mData
''    End Property
''
''    Public Property Get UserType() As Integer
''        UserType = intUserType
''    End Property
''
''    Public Property Let RequestID(mData As Integer)
''        intRequestID = mData
''    End Property
''
''    Public Property Get RequestID() As Integer
''        RequestID = intRequestID
''    End Property
''    Public Property Let DemandNo(mData As Variant)
''        mDemandNo = mData
''    End Property
''
''    Public Property Get DemandNo() As Variant
''        DemandNo = mDemandNo
''    End Property
''
'    Public Property Let FormName(mData As String)
'        strFormName = mData
'    End Property
'
'    Public Property Get FormName() As String
'        FormName = strFormName
'    End Property

    Public Property Let VerifyStatus(mData As Integer)
        intVerify = mData
    End Property

    Public Property Get VerifyStatus() As Integer
        VerifyStatus = intVerify
    End Property
    
    Public Property Let ReceiptVrID(mData As Integer)
        mReceiptVrID = mData
    End Property
    
    Public Property Let Request(mData As Variant)
        mRequestID = mData
    End Property

    Private Sub cmdApprove_Click()
        Dim mSql        As String
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objRev      As New clsReverseProcess
        Dim rVrID       As Variant
        Dim mStatus     As Boolean
        Dim nFlag       As Boolean
        Dim mCrl        As Control
        Dim mStat       As Boolean
            ''        If txtReqDate.Text <> "" Then
            ''            If gbTransactionDate < CDate(txtReqDate.Text) Then
            ''                MsgBox "Please Check the Requested Date..", vbApplicationModal
            ''                Exit Sub
            ''            End If
            ''        End If
        cmdVerify.Enabled = False
        If mRequestID > 0 Then
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            objRev.TransactionTypeID = mTrType
            If gbSeatGroupID = gbSeatGroupAccountsOfficer And (gbLBType = 3 Or gbLBType = 4) Then
                If mCnn.State Then
                    mSql = "Update faReverseEntry set tnyStatus=1 "
                    mSql = mSql + ",numAuthorisedByAO=" & gbUserID
                    mSql = mSql + ",dtAuthorisationDateAO='" & Format(gbTransactionDate, "dd/mmm/yy") & "'"
                    mSql = mSql + "Where intRequestID=" & mRequestID
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    MsgBox "Request for Reverse Entry is Recommended to Secretary ", vbInformation
                    cmdApprove.Enabled = False
                    cmdReject.Enabled = False
                End If
            ElseIf (gbSeatGroupID = gbSeatGroupSecretary And (gbLBType = 3 Or gbLBType = 4)) Or (gbSeatGroupID = gbSeatGroupAccountsOfficer And (gbLBPanchayat = 1)) Then
                On Error GoTo ErrRollBack
                mCnn.BeginTrans
                If mCnn.State Then
                    If IsDate(mdtTrDate) Then
                       objRev.TransactionDate = mdtTrDate
                    Else
                       MsgBox "TrasactionDate is not specified!", vbInformation
                       Exit Sub
                    End If
                    mSql = " Update faReverseEntry set tnyStatus=2 "
                    mSql = mSql + " ,numAuthorisedBySec=" & gbUserID
                    mSql = mSql + " ,dtAuthorisationDateSec='" & Format(gbTransactionDate, "dd/mmm/yy") & "'"
                    mSql = mSql + " Where intRequestID=" & mRequestID
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    objRev.VoucherID = val(txtVoucherNo.Tag)
                    If mVrType = 10 Then
                        If mCategory = 1 Then
                            
                            rVrID = objRev.ReverseTransaction(val(txtVoucherNo.Tag), mCnn)
                            If rVrID = "" Then
                                MsgBox "Transaction Failed", vbInformation
                                GoTo ErrRollBack
                            End If
                            '--------------------------------------------'
                            '
                            '--------------------------------------------'
                            mStatus = ReceiptSave(mCnn)
                            If mStatus = False Then
                                MsgBox "Transaction Failed", vbInformation
                                GoTo ErrRollBack:
                            End If
                        ElseIf mCategory = 2 Then 'Wrong Demand. Status will change after Receipt Save through receipt Screen
                            
                            mSql = ""
                            mSql = " Update faReverseEntry set tnyStatus=3 "
                            mSql = mSql + " Where intRequestID=" & mRequestID
                            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                            objRev.VoucherID = val(txtVoucherNo.Tag)
                            Call TotAmountCheck
                            mStat = True
                            MsgBox "This Voucher is Approved for Reverse," + vbCrLf + ""
                            rVrID = objRev.ReverseTransaction(val(txtVoucherNo.Tag), mCnn)
                            If rVrID = "" Then
                                MsgBox "Transaction Failed", vbInformation
                                GoTo ErrRollBack
                            End If
                        ElseIf mCategory = 3 Then
                            mStatus = ReceiptParticularsSave(mCnn)
                        Else
                            
                            rVrID = objRev.ReverseTransaction(val(txtVoucherNo.Tag), mCnn)
                            If rVrID = "" Then
                                MsgBox "Transaction Failed", vbInformation
                                GoTo ErrRollBack
                             End If
                        End If
                        
                    ElseIf mVrType = 30 Then
                        rVrID = objRev.ReverseTransaction(val(txtVoucherNo.Tag), mCnn)
                        If rVrID = "" Then
                            MsgBox "Transaction Failed", vbInformation
                            GoTo ErrRollBack
                        End If
                        nFlag = JournalCheckForCV(mCnn)
                    Else
                        rVrID = objRev.ReverseTransaction(val(txtVoucherNo.Tag), mCnn)
                        If rVrID = "" Then
                            MsgBox "Transaction Failed", vbInformation
                            GoTo ErrRollBack
                        End If
                    End If
                    
                    '**********************************************************************************************************************
                            'Call UpdateVoucherIndex(val(txtVoucherNo.Tag))    'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
                    mSql = "UPDATE faVoucherIndex SET tnyChangeFlag=1 WHERE intVoucherID = " & val(txtVoucherNo.Tag)
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    
                    '**********************************************************************************************************************
                    
                    If mStat = False Then
                        '--------------------Cochin Corporation For Reverse Entry--------------
'''''                        If gbLocalBodyID = 169 Then ''05/12/2015 for Cochin Corpn
'''''                                If gbTransactionTypePTax = 1 Then
'''''                                    Dim mMode As String
'''''                                    Dim mColAccID       As String
'''''                                    Dim mColKeyID       As String
'''''                                    Dim mCollPost       As String
'''''                                    Dim mUrl            As String
'''''                                    Dim mRevDate        As Date
'''''                                    Dim xmlHttp         As Object
'''''                                    Dim mXmlString      As Variant
'''''                                    Dim oRs             As ADODB.Recordset
'''''                                    Dim oNode           As Object 'MSXML2.IXMLDOMNode
'''''                                    Dim oSubNodes       As Object 'MSXML2.IXMLDOMSelection
'''''                                    Dim oDoc            As Object
'''''                                    Dim params          As String
'''''                                    'Dim mSql            As String
'''''
'''''                                    'objRev.VoucherID = val(txtVoucherNo.Tag)
'''''                                    'mSql = "SELECT "
'''''                                    Set xmlHttp = CreateObject("MSXML2.xmlHttp")
'''''                                    'If mTransactionType = 1 Then
'''''                                    mRevDate = gbTransactionDate            'dateOfRevoking
'''''                                    mMode = "Reverse Entry"               'modeOfRevoking
'''''                                    mCollPost = CStr(txtVoucherNo.Tag) + "~" + CStr(mMode) + "~" + CStr(mRevDate)
'''''                                    mUrl = gbDefaultUrl + "/updatePaymentRevoking?paymentRevokeParam=" + mCollPost
'''''                                    xmlHttp.Open "POST", mUrl, False
'''''                                    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
'''''                                    xmlHttp.send
'''''                                End If
'''''                            End If
'''''                            Exit Sub
                            '--------------------Cochin Corporation For Reverse Entry-------------
                        MsgBox "Reverse Process Done Successfully", vbInformation
                    End If
                    cmdApprove.Enabled = False
                    cmdReject.Enabled = False
                End If
              mCnn.CommitTrans
                If mVrType = 10 And rVrID <> "" Then
                    If gbFetchDemandFromWeb And mTrType = 1 Then
                     PTaxWebDemand (val(txtVoucherNo.Tag))
                    End If
                End If
            End If
        End If
        
        Exit Sub
ErrRollBack:
    mCnn.RollbackTrans
    MsgBox "Reverse Process Failed", vbApplicationModal
    End Sub
    Private Function JournalCheckForCV(mCn As ADODB.Connection) As Boolean
        Dim mSql            As String
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim objRev          As New clsReverseProcess
        Dim mVoucherID      As Variant
        Dim mRevVrID        As Variant
        Dim mFlag           As Boolean
        mFlag = False
        mSql = "Select intVoucherID From faVouchers Where tnyVoucherTypeID=40 And intKeyID2 in (Select intVoucherNo From faVouchers Where intVoucherID=" & val(txtVoucherNo.Tag) & ")"
        Rec.Open mSql, mCn
        If Not (Rec.EOF And Rec.BOF) Then
'            While Not (Rec.EOF And Rec.BOF)
                mVoucherID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                If mVoucherID <> "" Then
                    mRevVrID = objRev.ReverseTransaction(mVoucherID, mCn)
                End If
'                Rec.MoveNext
'            Wend
        End If
        If mRevVrID <> "" Then
            JournalCheckForCV = mFlag
        End If
    End Function
    Private Function ReceiptParticularsSave(mCn As ADODB.Connection) As Boolean
        Dim objdb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim mCnn                As New ADODB.Connection
        Dim mSql                As String
        Dim aryIn               As Variant
        Dim intInstrumentTypeID As Integer
        Dim vchInstrumentNo     As String
        Dim dtInstrumentDate    As String
        Dim vchDescription      As String
        Dim vchName             As String
        Dim vchHouseName        As String
        Dim vchStreetName       As String
        Dim vchMainPlace        As String
        Dim vchPostOffice       As String
        Dim vchDistrict         As Variant
        Dim vchPinNumber        As String
        Dim vchInit1            As String
        Dim vchInit2            As String
        Dim vchInit3            As String
        Dim vchInit4            As String
        Dim vchLocalPlace       As String
        Dim vchPhone            As String
        Dim intWardNo           As Integer
        Dim intDoorNo           As Integer
        Dim vchDoorNo2          As String
        Dim vchBank             As String
        Dim vchBankPlace        As String
        aryIn = Array(mDemandNo)
        Set Rec = objdb.ExecuteSP("spGetIDemandDetails", aryIn, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
'            @intVoucherID_1     [bigint],
'            @intLocalBodyID_2  [int],
'            Dim intInstrumentTypeID As Integer
            vchInstrumentNo = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            intInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
            dtInstrumentDate = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
            vchBank = IIf(IsNull(Rec!vchDrawnFrom), "", Rec!vchDrawnFrom)
            vchBankPlace = IIf(IsNull(Rec!vchDrawnPlace), "", Rec!vchDrawnPlace)
            vchDescription = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
            vchName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            vchHouseName = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            vchStreetName = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
            vchMainPlace = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            vchPostOffice = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
            vchDistrict = Null 'IIf(IsNull(Rec!vchPost, ""), Rec!vchPost)
            vchPinNumber = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
            vchInit1 = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
            vchInit2 = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
            vchInit3 = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
            vchInit4 = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
            vchLocalPlace = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
            vchPhone = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            intWardNo = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            intDoorNo = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
            vchDoorNo2 = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
            Set aryIn = Nothing
            aryIn = Array(val(txtVoucherNo.Tag), _
                        gbLocalBodyID, _
                        intInstrumentTypeID, _
                        vchInstrumentNo, _
                        dtInstrumentDate, _
                        vchBank, _
                        vchBankPlace, _
                        vchDescription, _
                        vchName, _
                        vchHouseName, _
                        vchStreetName, _
                        vchMainPlace, _
                        vchPostOffice, _
                        vchDistrict, _
                        vchPinNumber, _
                        vchInit1, _
                        vchInit2, _
                        vchInit3, _
                        vchInit4, _
                        vchLocalPlace, _
                        vchPhone, _
                        intWardNo, _
                        intDoorNo, _
                        vchDoorNo2)
                     objdb.ExecuteSP "spSaveReverseVrParticularsDetails", aryIn, , , mCn, adCmdStoredProc
        End If
    End Function
    Private Function ReceiptSave(mCn As ADODB.Connection) As Boolean
        Dim mFlag           As Boolean
        Dim mSql            As String
        Dim Rec             As New ADODB.Recordset
        Dim RecChild        As New ADODB.Recordset
        Dim RecSum          As New ADODB.Recordset
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim objDbs          As New clsDB
        Dim aryIn           As Variant
        Dim arrOutPut       As Variant
        Dim arrInput        As Variant
        Dim mintVoucherID   As Variant
        Dim mintVoucherNo   As Variant
        Dim mintTransactionID As Variant
        Dim mSlNo           As Integer
        Dim mTotal          As Double
        Dim mV              As uVoucher
        Dim mVC             As uVChild
        Dim mVA             As uVoucherAddress
        Dim mT              As uTr
        Dim mTC             As uTrChild
        Dim mKey            As Integer
        Dim mVoucherGroupID As Integer
        Dim numLinkKeyID      As Variant
        mFlag = False
        If mCn.State Then
            mSql = "Select Sum(fltAmount) as Sum From faIDemandChild Where numDemandID = " & mDemandNo
            Set RecSum = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (RecSum.EOF And RecSum.BOF) Then
                mTotal = RecSum!Sum
            Else
                GoTo ErrRollBack
            End If
            RecSum.Close
            aryIn = Array(mDemandNo)
            Set Rec = objdb.ExecuteSP("spGetIDemandDetails", aryIn, , , mCnn, adCmdStoredProc)
            If Not (Rec.EOF And Rec.BOF) Then
                With mV
                ''---------------------------------------------------------------------
                ''Assign Voucher details From DemandTable
                ''---------------------------------------------------------------------
                    .intVoucherID_1 = -1
                    .intLocalBodyID_2 = gbLocalBodyID
                    .intTransactionID_3 = Null
                    .intTransactionTypeID_4 = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    .tnyVoucherTypeID_5 = 10
                    .intVoucherNo_6 = Null
                    .intBookNo_7 = Null
                    .dtDate_8 = gbTransactionDate
                    .fltAmount_9 = mTotal
                    .intInstrumentTypeID_10 = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                    .vchInstrumentNo_11 = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                    .dtInstrumentDate_12 = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    .vchDescription_13 = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                    .numZoneID_14 = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                    .numWardID_15 = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                    .intDoorNoP1_16 = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                    .vchDoorNoP2_17 = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                    .vchDoorNoP3_18 = Null
                    .intUserID_19 = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
                    .intCounterID_20 = gbCounterID
                    .numSubLedgerID_21 = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                    .intKeyID1_22 = IIf(IsNull(Rec!intKeyID), "", Rec!intKeyID) ' val(txtCrHeadCode.Tag)
                    mKey = IIf(IsNull(Rec!intKeyID), -1, Rec!intKeyID)
                    .intKeyID2_23 = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                    .intExternalApplicationID_24 = Null
                    .intExternalModuleID_25 = 55
                    .intFinancialYearID_26 = gbFinancialYearID
                    .tnyShiftID_27 = gbShiftID
                    .tnyPrintFlag_28 = Null
                    .tnyCancelFlag_29 = 0
                    .vchBank_33 = IIf(IsNull(Rec!vchDrawnFrom), "", Rec!vchDrawnFrom)
                    .vchBankPlace_34 = IIf(IsNull(Rec!vchDrawnPlace), "", Rec!vchDrawnPlace)
                    .intFundID_35 = gbFundID
                    .numSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
                    .intSessionID = gbSessionID
                    .vchRefNo = IIf(IsNull(Rec!vchAdminNote), "", Rec!vchAdminNote)
                    .fltRoundOff = Null
                    .fltAdvAmtAdj = Null
                    .numInwardNo = Null
                    .tnyStatus_32 = 0
                    .numLocationID = gbLocationID
                    mVoucherGroupID = 0
                    numLinkKeyID = val(txtVoucherNo.Tag)

                    arrInput = Array(.intVoucherID_1, _
                    .intLocalBodyID_2, _
                    .intTransactionID_3, _
                    .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
                    .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
                    .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
                    .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
                    .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
                    .intKeyID2_23, .intExternalApplicationID_24, _
                    .intExternalModuleID_25, .intFinancialYearID_26, _
                    .tnyShiftID_27, .tnyPrintFlag_28, _
                    .tnyCancelFlag_29, .vchBank_33, _
                    .vchBankPlace_34, .intFundID_35, _
                    .numSeatID, .intSessionID, _
                    .vchRefNo, .fltRoundOff, _
                    .fltAdvAmtAdj, .numInwardNo, _
                    .tnyStatus_32, .numLocationID, mVoucherGroupID, numLinkKeyID)
                objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn, adCmdStoredProc
                End With
                
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID = arrOutPut(0, 0)
                Else
                    GoTo ErrRollBack:
                End If

                With mVA
                    .intVoucherID = mintVoucherID
                    .intLocalBodyID = gbLocalBodyID
                    .vchName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    .vchInit1 = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
                    .vchInit2 = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
                    .vchInit3 = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
                    .vchInit4 = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
                    .vchHouseName = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
                    .vchStreetName = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                    .vchLocalPlace = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                    .vchMainPlace = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                    .vchPostOffice = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
                    .vchDistrict = Null
                    .vchPinNumber = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
                    .vchPhone = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                    .intWardNo = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                    .intDoorNo = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                    .vchDoorNo2 = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                  
                    arrInput = Array(.intVoucherID, _
                        .intLocalBodyID, _
                        .vchName, _
                        .vchInit1, _
                        .vchInit2, _
                        .vchInit3, _
                        .vchInit4, _
                        .vchHouseName, _
                        .vchStreetName, _
                        .vchLocalPlace, _
                        .vchMainPlace, _
                        .vchPostOffice, _
                        .vchDistrict, _
                        .vchPinNumber, _
                        .vchPhone, _
                        .intWardNo, _
                        .intDoorNo, _
                        .vchDoorNo2)
                        
                    objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCn, adCmdStoredProc
                End With
                With mT
                    .intTransactionID = -1
                    .intLocalBodyID = gbLocalBodyID
                    .intFinancialYearID = gbFinancialYearID
                    .dtTransactionDate = gbTransactionDate
                    .intExternalApplicationID = Null
                    .intExternalApplicationModuleID = 55
                    .intFunctionID = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    .intFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                    .intFieldID = Null
                    .intFundID = gbFundID
                    .intBudgetCentreID = Null
                    .vchNarration = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                    .intTransactionTypeID = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    .intProcessID = Null
                    .vchGroup = "R"
                    .intGroupID = 10
                    .intKeyID = Null
                    .numSubLedgerID = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                    .intVoucherID = mintVoucherID
                    
                    
                    arrInput = Array(.intTransactionID, _
                    .intLocalBodyID, _
                    .intFinancialYearID, _
                    .dtTransactionDate, _
                    .intExternalApplicationID, _
                    .intExternalApplicationModuleID, _
                    .intFunctionID, _
                    .intFunctionaryID, _
                    .intFieldID, _
                    .intFundID, _
                    .intBudgetCentreID, _
                    .vchNarration, _
                    .intTransactionTypeID, _
                    .intProcessID, _
                    .vchGroup, _
                    .intGroupID, _
                    .intKeyID, _
                    .numSubLedgerID, _
                    .numUserID, _
                    .intVoucherID, mVoucherGroupID)
                    
                    objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn, adCmdStoredProc
                End With
            
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintTransactionID = arrOutPut(0, 0)
                Else
                    GoTo ErrRollBack
                End If
                mSlNo = 1
                With mTC
                    .intTransactionID = mintTransactionID
                    .intSerialNo = mSlNo
                    .intAccountHeadID = mKey
                    .fltAmount = mTotal
                    .tinDebitOrCreditFlag = 1
                    .intByAccountHeadID = Null
                    .vchNarration = Null
                    .intFundID = 1
                    
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn, adCmdStoredProc
                End With
            End If
            Rec.Close
            aryIn = Array(mDemandNo)
            Set RecChild = objdb.ExecuteSP("spGetIDemandDetails", aryIn, , , mCnn, adCmdStoredProc)
            If Not (RecChild.EOF And RecChild.BOF) Then
                 While Not (RecChild.EOF)
                     mSlNo = mSlNo + 1
                     
                     With mTC
                         .intTransactionID = mintTransactionID
                         .intSerialNo = mSlNo
                         .intAccountHeadID = IIf(IsNull(RecChild!intAccountHeadID), "", RecChild!intAccountHeadID)
                         .fltAmount = Format(IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount), "0.00")
                         .tinDebitOrCreditFlag = 0
                         .intByAccountHeadID = mKey
                         .vchNarration = IIf(IsNull(RecChild!vchRemarks), "", RecChild!vchRemarks)
                         .intFundID = 1
                 
                         arrInput = Array(.intTransactionID, _
                         .intSerialNo, _
                         .intAccountHeadID, _
                         .fltAmount, _
                         .tinDebitOrCreditFlag, _
                         .intByAccountHeadID, _
                         .vchNarration, _
                         .intFundID)
                         objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn, adCmdStoredProc
                     End With
            
                     With mVC
                         .intVoucherID_1 = mintVoucherID
                         .intLocalBodyID_2 = gbLocalBodyID
                         .intSlNo_3 = mSlNo - 1
                         .intAccountHeadID_4 = IIf(IsNull(RecChild!intAccountHeadID), "", RecChild!intAccountHeadID)
                         .tnyDebitOrCredit_5 = 0
                         .intYearID_6 = IIf(IsNull(RecChild!intYearID), "", RecChild!intYearID)
                         .tnyPeriodID_7 = IIf(IsNull(RecChild!tnyPeriodID), "", RecChild!tnyPeriodID)
                         .tnyArrearFlag_8 = IIf(IsNull(RecChild!tnyArrearFlag), "", RecChild!tnyArrearFlag)
                         .numDemandID_9 = Null
                         .fltAmount_10 = Format(IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount), "0.00")
             
                         arrInput = Array(.intVoucherID_1, _
                         .intLocalBodyID_2, _
                         .intSlNo_3, _
                         .intAccountHeadID_4, _
                         .tnyDebitOrCredit_5, _
                         .intYearID_6, _
                         .tnyPeriodID_7, _
                         .tnyArrearFlag_8, _
                         .numDemandID_9, _
                         .fltAmount_10)
                         objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn, adCmdStoredProc
                     End With
                RecChild.MoveNext
                Wend
            End If
            mFlag = True
        Else
            MsgBox "Connection With Saankhya Doesn't exists", vbInformation
        End If
        ReceiptSave = mFlag
        Exit Function
        
ErrRollBack:
        ReceiptSave = mFlag
        MsgBox ("Transaction Failed"), vbApplicationModal
    End Function

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdOld_Click()
        frmViewVoucher.MultipleVouchers = False
        frmViewVoucher.ArrayIn = Array(CStr(val(txtVoucherNo.Tag)))
        frmViewVoucher.FormName = "frmReverseEntryRequest"
        frmViewVoucher.Show vbModal
    End Sub
    Private Sub cmdReject_Click()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        
        If mRequestID > 0 Then
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                    If (MsgBox("This will cancel the Request......" & vbCrLf _
                    & "Do you want to proceed                  ", vbCritical + vbYesNo)) = vbYes Then
                        mSql = "Update faReverseEntry set tnyStatus=4 "
                        mSql = mSql + ",numAuthorisedByAO=" & gbUserID
                        mSql = mSql + ",dtAuthorisationDateAO=" & CDate(Format(gbTransactionDate, "DD/mmm/yy"))
                        mSql = mSql + "Where intRequestID=" & mRequestID
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                        MsgBox "The Request For Reverse Entry is Rejected"
                        cmdReject.Enabled = False
                        cmdApprove.Enabled = False
                    End If
                End If
            ElseIf gbSeatGroupID = gbSeatGroupAdditionalSecretary Then
                If (MsgBox("This will cancel the Request......" & vbCrLf _
                    & "Do you want to proceed                  ", vbCritical + vbYesNo)) = vbYes Then
                        mSql = "Update faReverseEntry set tnyStatus=4 "
                        mSql = mSql + ",numAuthorisedByAO=" & gbUserID
                        mSql = mSql + ",dtAuthorisationDateAO=" & CDate(Format(gbTransactionDate, "DD/mmm/yy"))
                        mSql = mSql + "Where intRequestID=" & mRequestID
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                        MsgBox "The Request for Reverse entry is Rejected"
                        cmdReject.Enabled = False
                        cmdApprove.Enabled = False
                    End If
            End If
        End If
    End Sub

    Private Sub cmdVerify_Click()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mVType  As Integer
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select * From faVouchers Where intVoucherID=" & val(txtVoucherNo.Tag)
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                mVType = Rec!tnyVoucherTypeID
            End If
        End If
        
'''        On Error GoTo err:
        frmViewVoucher.MultipleVouchers = False
        If mCategory = 0 Then
            frmViewVoucher.FormName = "frmReverseRequest"
            frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
            frmViewVoucher.cmdVerify.Visible = True
            frmViewVoucher.Show vbModal
        Else 'If mCategory = 1 Or mCategory = 2 Or mCategory = 3 Then
            If mVType = 10 Then
                frmViewVoucher.FormName = "frmReverseDemand"
                frmViewVoucher.ArrayIn = Array(CStr(mDemandNo))
                frmViewVoucher.cmdVerify.Visible = True
                frmViewVoucher.Show vbModal
            Else
                frmViewVoucher.FormName = "frmReverseRequest"
                frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
                frmViewVoucher.cmdVerify.Visible = True
                frmViewVoucher.Show vbModal
            End If
        End If
        If VerifyStatus = 1 Then
            cmdApprove.Enabled = True
        Else
            cmdApprove.Enabled = False
        End If
        Exit Sub
        
err:
        MsgBox (Error$)
    End Sub
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
'            cmdApprove.Caption = "Recommend"
            Call ReverseRequestDetails(mRequestID)
        ElseIf gbSeatGroupID = gbSeatGroupSecretary Then
            cmdReject.Enabled = False
            Call ReverseRequestDetails(mRequestID)
        End If
    End Sub
    Private Sub ReportView(ByVal mDemandNo As String)
         Dim Rpt As New CRAXDRT.Report
         Dim mApp As New CRAXDRT.Application
         Dim rptFileName As String
         Dim arrInput As Variant
         Dim mLoop As Long
'         If mCategory > 0 Then
'            rptFileName = App.Path & "\Reports\rptDemand.rpt"
'         Else 'If mCategory = 0 Then
          rptFileName = App.Path & "\Reports\rptVoucher.rpt"
'         End If
         arrInput = Array(mDemandNo)
         Screen.MousePointer = vbHourglass
         crvReport.DisplayTabs = True
         
         Set Rpt = Nothing
         mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
         Set Rpt = mApp.OpenReport(rptFileName, 1)
         
         If IsArray(arrInput) Then
             For mLoop = LBound(arrInput) To UBound(arrInput)
                 Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
                 Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
             Next mLoop
         End If
         Screen.MousePointer = vbDefault
         crvReport.ReportSource = Rpt
         crvReport.ViewReport
         crvReport.Zoom (1)
    End Sub
    Private Sub ReportOlD(ByVal mVrNo As String)
         Dim Rpt As New CRAXDRT.Report
         Dim mApp As New CRAXDRT.Application
         Dim rptFileName As String
         Dim arrInput As Variant
         Dim mLoop As Long
         
         rptFileName = App.Path & "\Reports\rptVoucher.rpt"
         arrInput = Array(mVrNo)
         Screen.MousePointer = vbHourglass
         crvReportOld.DisplayTabs = True
         
         Set Rpt = Nothing
         mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
         Set Rpt = mApp.OpenReport(rptFileName, 1)
         
         If IsArray(arrInput) Then
             For mLoop = LBound(arrInput) To UBound(arrInput)
                 Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
                 Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
             Next mLoop
         End If
         Screen.MousePointer = vbDefault
         crvReportOld.ReportSource = Rpt
         crvReportOld.ViewReport
         crvReportOld.Zoom (1)
    End Sub
    Private Sub ReverseRequestDetails(ByVal ReqID As Double)
        'On Error GoTo err:
            Dim mCnn            As New ADODB.Connection
            Dim Rec             As New ADODB.Recordset
            Dim mSql            As String
            Dim objdb           As New clsDB
            Dim mDemand         As String
            Dim numApprover     As String
            Dim numApproverDate As String
            Dim mRequestDate    As Date
            Dim mFinYearID      As Integer
            If objdb.SetConnection(mCnn) Then
                mSql = " Select faVouchers.intVoucherID intVoucherID, faVouchers.intVoucherNo,faVouchers.dtDate, faVouchers.tnyVoucherTypeID,"
                mSql = mSql + " numDemandNo, faReverseEntry.intReasonID, faReasons.vchReason,  faReasons.intCategory ReasonType,intKeyID1, "
                mSql = mSql + " chvSeatTitle, vchUserName,faReverseEntry.dtRequestDate, faReverseEntry.vchRemarks as Remark,faVouchers.intFinancialYearID FinYear, "
                mSql = mSql + " faPendingTaskRequest.intRequestID intPendingRequestID, faPendingTaskRequest.dtTransactionDate dtPendintTrDate, "
                 mSql = mSql + " numAuthorisedByAO , dtAuthorisationDateAO,faVouchers.intTransactionTypeID"
                mSql = mSql + " From faReverseEntry "
                mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
                mSql = mSql + " Inner Join faReasons On faReasons.intReasonID = faReverseEntry.intReasonID"
                mSql = mSql + " Left Join faVouchers On faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                mSql = mSql + " Left Join faUser On faUser.numUserID = faReverseEntry.numRequestedUserID"
                mSql = mSql + " Left Join faSeats On faSeats.numSeatID = faReverseEntry.numRequestedSeatID"
                mSql = mSql + " Left Join faPendingTaskRequest ON faPendingTaskRequest.intKeyID = faReverseEntryChild.intVoucherID  AND intTaskID=6"
                mSql = mSql + " Where faReasons.intType = 55 And faReverseEntry.intRequestID =  " & ReqID
                mSql = mSql + " And faReverseEntry.tnyStatus <> 4 "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Not IsNull(Rec!intPendingRequestID) Then
                        mPreviousYearID = 1
                        mPreviousYearRequestID = Rec!intPendingRequestID
                        mdtTrDate = Rec!dtPendintTrDate
                    Else
                        mPreviousYearID = 0
                        mPreviousYearRequestID = -1
                        mdtTrDate = gbTransactionDate
                    End If
                
                    txtVoucherNo.Tag = Rec!intVoucherID
                    txtVoucherNo.Text = Rec!intVoucherNo
                    txtVrDate.Text = Format(Rec!dtDate, "dd-mmm-yyyy")
                    txtReason.Tag = Rec!intReasonID
                    txtReason.Text = Rec!vchReason
                    mCategory = Rec!ReasonType
                    mVrType = Rec!tnyVoucherTypeID
                    mDemandNo = Rec!numDemandNo
                    txtSeat.Text = Rec!chvSeatTitle
                    txtUser.Text = Rec!vchUserName
                    txtReqDate.Text = Format(Rec!dtRequestDate, "dd-mmm-yyyy")
                    mRequestDate = Format(Rec!dtRequestDate, "dd-mmm-yyyy")
                    mKeyID = Rec!intKeyID1
                    mFinYearID = Rec!FinYear
                    lblRemarks.Caption = Rec!Remark
                    numApprover = IIf(IsNull(Rec!numAuthorisedByAO), "", Rec!numAuthorisedByAO)
                    numApproverDate = IIf(IsNull(Rec!dtAuthorisationDateAO), "", Rec!dtAuthorisationDateAO)
                    mTrType = Rec!intTransactionTypeID
                End If
                Rec.Close
                
                If mFinYearID <> gbFinancialYearID Then
                    If mPreviousYearID = 1 Then
                        If mFinYearID <> (gbFinancialYearID - 1) Then
                            MsgBox "This Voucher Doesnot exists in this FinancialYear", vbInformation
                            cmdVerify.Enabled = False
                            cmdApprove.Enabled = False
                            Exit Sub
                        End If
                    Else
                        MsgBox "This Voucher Doesnot exists in this FinancialYear", vbInformation
                        cmdVerify.Enabled = False
                        cmdApprove.Enabled = False
                        Exit Sub
                    End If
                    
                End If
                
                '------------------------------------------------
                ''Intermediary Approval
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    'Date Validation
                    If mRequestDate > gbTransactionDate Then
                        MsgBox "Approve date Must be Greater than Request date", vbInformation
                        cmdVerify.Enabled = False
                        cmdApprove.Enabled = False
                        Exit Sub
                    End If
                End If
                If gbSeatGroupID = gbSeatGroupSecretary Then
                    '---------------------------------
                    'Date Validation
                    If numApproverDate > gbTransactionDate Then
                        MsgBox "Date Must be Greater than Request date And Inter Approve date", vbInformation
                        cmdVerify.Enabled = False
                        cmdApprove.Enabled = False
                        Exit Sub
                    End If
                    '---------------------------------
                    If numApprover <> "" Then
                        mSql = ""
                        mSql = "Select * From faUser Where numUserID= " & numApprover
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF Or Rec.BOF) Then
                            txtApprover.Text = Rec!vchUserName
                            txtApproveDate.Text = Format(numApproverDate, "dd-mmm-yyyy")
                        End If
                    End If
                End If
                
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact your System Administrator"
            End If
'        If mCategory > 0 Then
'            crvReport.Visible = True
'            Call ReportView(mDemandNo)
'        Else 'If mCategory = 0 Then
            crvReport.Visible = False
            Call ReportView(val(txtVoucherNo.Tag))
'        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Function TotAmountCheck() As Integer
        '  Not Implemented
        Dim PO As uPaymentOrder
        Dim POC As uPaymentOrderChild
        Dim POAdd As uPaymentOrderAddress
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim RecVr         As New ADODB.Recordset
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mVrTotal    As Double
        
        If objdb.SetConnection(mCnn) Then
            mSql = "Select Sum(fltAmount) as Sum From faIDemandChild Where numDemandID=" & mDemandNo
            Set Rec = objdb.ExecuteSP(mSql, mDemandNo, , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                mDemandTotal = Rec!Sum
            End If
            mSql = "Select fltAmount From faVouchers Where intVoucherID=" & val(txtVoucherNo.Tag)
            Set RecVr = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (RecVr.EOF Or RecVr.BOF) Then
                mVrTotal = RecVr!fltAmount
            End If
            If mDemandTotal = mVrTotal Then
                TotAmountCheck = True
            End If
        End If
    End Function
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
                    Dim objdb As New clsDB

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
  


