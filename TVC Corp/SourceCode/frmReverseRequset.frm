VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReverseRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reverse Entry Request"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   Icon            =   "frmReverseRequset.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeRequest 
      Height          =   3795
      Left            =   0
      TabIndex        =   19
      Tag             =   "-1"
      Top             =   360
      Width           =   9690
      Begin VB.CommandButton cmdSeat 
         Caption         =   "..."
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   1230
         Width           =   300
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
         TabIndex        =   7
         Top             =   1215
         Width           =   3270
      End
      Begin VB.TextBox txtVoucherNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   450
         Width           =   3270
      End
      Begin VB.CommandButton cmdSearchVoucher 
         Caption         =   "..."
         Height          =   285
         Left            =   4860
         TabIndex        =   2
         Top             =   465
         Width           =   300
      End
      Begin VB.TextBox txtRemarks 
         Height          =   975
         Left            =   1575
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1620
         Width           =   3630
      End
      Begin VB.TextBox txtReason 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   4
         Top             =   855
         Width           =   3270
      End
      Begin VB.CommandButton cmdSearchReason 
         Caption         =   "..."
         Height          =   315
         Left            =   4860
         TabIndex        =   5
         Top             =   855
         Width           =   300
      End
      Begin VB.Label lblDemandNo 
         Height          =   285
         Left            =   945
         TabIndex        =   22
         Top             =   3465
         Width           =   2715
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
         TabIndex        =   21
         Top             =   3510
         Width           =   915
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
         TabIndex        =   6
         Top             =   1215
         Width           =   1500
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
         TabIndex        =   0
         Top             =   465
         Width           =   1170
      End
      Begin VB.Label lblVoucherType 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1530
         TabIndex        =   20
         Top             =   135
         Width           =   3255
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
         TabIndex        =   9
         Top             =   1620
         Width           =   900
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
         TabIndex        =   3
         Top             =   810
         Width           =   765
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
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   45
         Width           =   9600
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   9645
      TabIndex        =   15
      Top             =   0
      Width           =   9705
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reverse Entry Request"
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
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9645
      TabIndex        =   14
      Top             =   4620
      Width           =   9705
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   9585
         Top             =   450
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   4
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton cmdRequest 
         Caption         =   "Request"
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
         Left            =   3780
         TabIndex        =   12
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
         Left            =   2025
         TabIndex        =   11
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
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
         Left            =   5535
         TabIndex        =   13
         Top             =   45
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmReverseRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Option Explicit
    Dim intInstrumentTypeID As Variant
    Private intVerify       As Integer    '       0 = not verified;   1 = verified
    Public mRevDemand       As Boolean
    Public mDemandNo        As Variant
    Private mPreviousYearMode       As Integer
    Private mPreviousYearRequestID  As Integer
    
'    Dim mDemandNo       As Variant
'    Dim mKeyID          As Integer
'    Dim mCategory       As Integer
'    Dim mVrType         As Integer
'    Dim mDemandTotal    As Double
    Public Property Let VerifyStatus(mData As Integer)
        intVerify = mData
    End Property
    Public Property Get VerifyStatus() As Integer
        VerifyStatus = intVerify
    End Property
    Private Sub cmdClear_Click()
        If cmdClear.Caption = "Clear" Then
            Dim mCrl As Control
             For Each mCrl In Me.Controls
                If TypeOf mCrl Is TextBox Then
                    mCrl.Text = ""
                    mCrl.Tag = ""
                    cmdRequest.Enabled = False
                End If
            Next
            cmdVerify.Enabled = True
            cmdClear.Caption = "Close"
        Else
            Unload Me
        End If
    End Sub
    
    Private Sub cmdRequest_Click()
        Dim mCrl As Control
        Dim mStatus As Integer
        Dim mID     As Variant
        Dim mType   As Integer 'Edit Mode=1
        Dim mSql    As String
            
        If val(fmeRequest.Tag) > 0 Then
            If lblDemandNo.Tag <> "" Then
                mType = 1
                mID = lblDemandNo.Tag
            Else
                mType = 0
                mID = txtVoucherNo.Tag
            End If
        Else
            mType = 0
            mID = txtVoucherNo.Tag
        End If
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupChiefCashier Then
                If RequsetValidation = False Then Exit Sub
                mStatus = CheckReverseRequestExist(val(txtVoucherNo.Tag))
                If mStatus = 1 Or mStatus = 2 Then
                    MsgBox "Request Already Exists", vbInformation
                    Exit Sub
                ElseIf mStatus = 0 And fmeRequest.Tag = "" Then
                    MsgBox "Request Already Exists", vbInformation
                    Exit Sub
                End If
                If txtVoucherNo.Tag = "" Then
                    lblMsgBox.Visible = True
                    lblMsgBox.Caption = "Please Select A Voucher to Do Verification"
                    Exit Sub
                End If
                If lblVoucherType.Tag = 10 Then
                    If cmdSearchReason.Tag <> 0 Then
                    
                    '-------------------------------
                    'Reason category
                    '0 No demand,1= Demand,2= Through Receipt screen(integrated module),3=Seperate table for Storing Changed Details (No reverse process)
                    '-------------------------------
                        Select Case val(txtReason.Tag)
                            Case 504  'Transaction Type
                                    MsgBox "Please Select Correct Transaction Type", vbInformation
                                    With frmDemandInterface
                                        On Error Resume Next
                                        For Each mCrl In frmDemandInterface.Controls
                                            If TypeOf mCrl Is ComboBox Then
                                                mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is CommandButton Then
                                                    mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is TextBox Then
                                                    mCrl.Enabled = False
                                            End If
                                        Next
                                        .Reverse = 1
                                        .ReverseDemandDetails mID, mType
                                        .vsGrid.Editable = flexEDKbdMouse
                                        .cmdSave.Enabled = True
                                        .cmdCancel.Enabled = True
                                        '.cmbTransactionType.Enabled = True
                                        .cmdSearchTransactionType.Enabled = True
                                        .txtTransactionType.Enabled = True
                                        .Show vbModal
                                    End With
                            Case 505  'Instrument Type
                                    MsgBox "Please Select Instrument", vbInformation
                                    With frmDemandInterface
                                        
                                        On Error Resume Next
                                        For Each mCrl In frmDemandInterface.Controls
                                            If TypeOf mCrl Is ComboBox Then
                                                mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is CommandButton Then
                                                    mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is TextBox Then
                                                    mCrl.Enabled = False
                                            End If
                                        Next
                                         .cmdAcHead.Visible = True
                                        .cmdAcHead.Enabled = True
                                        .Reverse = 1
                                        .ReverseDemandDetails mID, mType
                                        .vsGrid.Enabled = False
                                        .cmdSave.Enabled = True
                                        .cmdCancel.Enabled = True
                                        .txtAccountCode.Visible = True
                                        .txtAccountHead.Visible = True
                                        .txtAccountCode.Enabled = True
                                        .cmbInstrumentType.Enabled = True
                                        .cmdSearchTransactionType.Enabled = True
                                        .txtTransactionType.Enabled = True
                                        .Show vbModal
                                    End With
                            Case 506  'Income Head
                                MsgBox "Please Edit Account Head in the Required Field", vbInformation
                                With frmDemandInterface
                                    
                                    On Error Resume Next
                                    For Each mCrl In frmDemandInterface.Controls
                                        If TypeOf mCrl Is ComboBox Then
                                            mCrl.Enabled = False
                                        ElseIf TypeOf mCrl Is CommandButton Then
                                                mCrl.Enabled = False
                                        ElseIf TypeOf mCrl Is TextBox Then
                                                mCrl.Enabled = False
                                        End If
                                    Next
                                    .Reverse = 1
                                    .ReverseDemandDetails mID, mType
                                    .cmdSave.Enabled = True
                                    .cmdCancel.Enabled = True
                                    .Show vbModal
                                End With
                            Case 507  'Amount
                                With frmDemandInterface
                                    MsgBox "Please Edit Amount in the Required Field", vbInformation
                                    On Error Resume Next
                                    For Each mCrl In frmDemandInterface.Controls
                                        If TypeOf mCrl Is ComboBox Then
                                            mCrl.Enabled = False
                                        ElseIf TypeOf mCrl Is CommandButton Then
                                                mCrl.Enabled = False
                                        ElseIf TypeOf mCrl Is TextBox Then
                                                mCrl.Enabled = False
                                        End If
                                    Next
                                    .Reverse = 1
                                    .ReverseDemandDetails mID, mType
                                    .cmdSave.Enabled = True
                                    .cmdCancel.Enabled = True
                                    .Show vbModal
                                End With
                            Case 508  'Wrong demand
                            
                                    
''                                    MsgBox "Please Enter Data to correct the Voucher in the Required Field", vbInformation
''                                    With frmDemandInterface
''                                        On Error Resume Next
''                                        For Each mCrl In frmDemandInterface.Controls
''                                            If TypeOf mCrl Is ComboBox Then
''                                                mCrl.Enabled = False
''                                            ElseIf TypeOf mCrl Is CommandButton Then
''                                                    mCrl.Enabled = False
''                                            ElseIf TypeOf mCrl Is TextBox Then
''                                                    mCrl.Enabled = False
''                                            End If
''                                        Next
''                                        .Reverse = 1
''                                        .ReverseDemandDetails mID, mType
''                                        .vsGrid.Editable = flexEDNone
''                                        .cmdSave.Enabled = True
''                                        .cmdCancel.Enabled = True
''                                        .cmdSearchTransactionType.Enabled = True
''                                        .txtTransactionType.Enabled = True
''                                        .cmdSearchTransactionType.Enabled = True
''                                        .txtTransactionType.Enabled = True
''                                        .txtWardNo.Enabled = True
''                                        .cmbZone.Enabled = True
''                                        .Show vbModal
''                                    End With
                            Case 509 'Particulars
                                    MsgBox "Please Enter Correct Details of the Voucher", vbInformation
                                    With frmDemandInterface
                                        On Error Resume Next
                                        For Each mCrl In frmDemandInterface.Controls
                                            If TypeOf mCrl Is ComboBox Then
                                                mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is CommandButton Then
                                                    mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is TextBox Then
                                                    mCrl.Enabled = False
                                            End If
                                        Next
                                        .Reverse = 1
                                        .ReverseDemandDetails mID, mType
                                        .vsGrid.Editable = flexEDNone
                                        .cmdSave.Enabled = True
                                        .cmdCancel.Enabled = True
                                        .txtName.Enabled = True
                                        .txtHouseName.Enabled = True
                                        .txtPhone.Enabled = True
                                        .txtDrawnFrom.Enabled = True
                                        .txtDrawnPlace.Enabled = True
                                        .txtMainPlace.Enabled = True
                                        .txtInitial1.Enabled = True
                                        .txtInitial2.Enabled = True
                                        .txtInitial3.Enabled = True
                                        .txtInitial4.Enabled = True
                                        .txtPin.Enabled = True
                                        .txtPost.Enabled = True
                                        .txtInstrumentNo.Enabled = True
                                        .Show vbModal
                                    End With
                            End Select
                            If mRevDemand = True Then
                                Call SaveRequest
                                
                            Else
                                MsgBox "Request Failed", vbCritical
                                Exit Sub
                            End If
                    Else
                        Call SaveRequest
                    End If
                Else
                    Call SaveRequest
                End If
                
                
'                    cmdVerify.Enabled = False
'                    cmdRequest.Enabled = False
                    Call cmdClear_Click
            End If
    End Sub
    Private Function funTransTypeReasonMatch() As Boolean
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mStatus     As Integer
        Dim mTrTypeID   As Integer
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select intTransactionTypeID From faVouchers Where intVoucherID=" & txtVoucherNo.Tag
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            Select Case Rec!intTransactionTypeID
                Case 1
                    If gbLinkWithPropertyTax = 1 Then
                        funTransTypeReasonMatch = True
                    Else
                        funTransTypeReasonMatch = False
                    End If
                
                Case 3
                    If gbLinkWithProfTaxEmp Then
                        funTransTypeReasonMatch = True
                    Else
                        funTransTypeReasonMatch = False
                    End If
                Case 4, 5
                    If gbLinkWithRentOnLand Then
                        funTransTypeReasonMatch = True
                    Else
                        funTransTypeReasonMatch = False
                    End If
                Case Else
                    funTransTypeReasonMatch = False
            End Select
        End If
    End Function
    Private Sub cmdSearchVoucher_Click()
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mStatus     As Integer
        Dim mRevSatatus As Integer
        frmSearchVouchers.CheckMode = 10
        
        frmSearchVouchers.chkInterrupted.Visible = False
        frmSearchVouchers.chkPayment.Enabled = False
        
        frmSearchVouchers.Show vbModal
        lblMsgBox.Caption = ""
        
        If gbSearchID <> -1 Then
            GetVoucherDetails (gbSearchID)
            mStatus = CheckReverseRequestExist(gbSearchID)
            If mStatus = 0 Or mStatus = 1 Or mStatus = 2 Then
                MsgBox "Request Already Exists", vbInformation
                'Call cmdClear_Click
                txtVoucherNo.Text = ""
                txtVoucherNo.Tag = ""
                Exit Sub
            End If
            If JSkCounterVerification = False Then
                MsgBox "JanasevanaKedram Section MissMatch" & vbCrLf & " This Voucher is not Allowed to Reverse in This Login"
               ' Call cmdClear_Click
                txtVoucherNo.Text = ""
                txtVoucherNo.Tag = ""
                Exit Sub
            End If
            mRevSatatus = ReverseStatus(gbSearchID)
            If mRevSatatus <> 0 Then
                txtVoucherNo.Text = ""
                txtVoucherNo.Tag = ""
                Exit Sub
            End If
            If AdvJournal = True Then
                MsgBox "This Voucher has done Adjustment Entry.. U can't Reverse.."
                txtVoucherNo.Text = ""
                txtVoucherNo.Tag = ""
                Exit Sub
            End If
            If CheckReceiptACRMode = True Then
                MsgBox "This is an autogenerated receipt of Development Expenditure.." & vbCrLf & "  U can't Reverse.."
                txtVoucherNo.Text = ""
                txtVoucherNo.Tag = ""
                Exit Sub
            End If
''            If CheckReconciled(gbSearchID) = True Then
''                MsgBox "This Voucher is Reconciled ", vbInformation
''                txtVoucherNo.Text = ""
''                txtVoucherNo.Tag = ""
''                Exit Sub
''            End If
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSql = "Select tnyVoucherTypeID,isNull(intInstrumentTypeID,100) intInstrumentTypeID,intFinancialYearID,dtDate,numLinkKeyID From faVouchers Where intVoucherID=" & gbSearchID
            Rec.Open mSql, mCnn
            If Not Rec.EOF Then
                If Not (IsNull(Rec!numLinkKeyID)) Then
                    MsgBox "This is Reversed Voucher", vbInformation
                    gbSearchCode = ""
                    gbSearchStr = ""
                    gbSearchID = -1
                    txtVoucherNo.Text = ""
                    txtVoucherNo.Tag = ""
                    Exit Sub
                End If
                If CDate(Rec!dtDate) > CDate(gbTransactionDate) Then
                    MsgBox "Request Date must be Greater/Equal to Voucher Date", vbInformation
                    gbSearchCode = ""
                    gbSearchStr = ""
                    gbSearchID = -1
                    txtVoucherNo.Text = ""
                    txtVoucherNo.Tag = ""
                    Exit Sub
                End If
                If Rec!intFinancialYearID <> gbFinancialYearID Then
                    MsgBox "Sorry!.. This Voucher is not in the Current Financial Year", vbInformation
                    gbSearchCode = ""
                    gbSearchStr = ""
                    gbSearchID = -1
                    txtVoucherNo.Text = ""
                    txtVoucherNo.Tag = ""
                    Exit Sub
                End If
                If Rec!tnyVoucherTypeID = 20 Then
                    MsgBox "Payment Voucher is Not Allowed To Reverse", vbInformation
                    gbSearchCode = ""
                    gbSearchStr = ""
                    gbSearchID = -1
                    txtVoucherNo.Text = ""
                    txtVoucherNo.Tag = ""
                    Exit Sub
                ElseIf Rec!tnyVoucherTypeID = 30 Then
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            MsgBox "Already sent Request for this Voucher", vbInformation
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            MsgBox "This Voucher Already Reversed", vbInformation
                        Else
                            Call GetVoucherDetails(gbSearchID)
                        End If
                ElseIf Rec!tnyVoucherTypeID = 40 Then
                        If AutoJournalCheck Then
                            MsgBox "This Journal is AutoGenerated Not allowed to Reverse"
                            Call cmdClear_Click
                            Exit Sub
                        End If
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            MsgBox "Already sent Request for this Voucher", vbInformation
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            MsgBox "Voucher Already Reversed", vbInformation
                        Else
                            Call GetVoucherDetails(gbSearchID)
                        End If
                ElseIf Rec!tnyVoucherTypeID = 10 Then
                    If Rec!intInstrumentTypeID = 5 Then
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "Already sent Request for this Voucher"
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "This Voucher Already Reversed"
                        Else
                            Call GetVoucherDetails(gbSearchID)
                        End If
                    Else
                        
                        If CheckReverseRequestExist(gbSearchID) = 1 Then
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "Already sent Request for this Voucher"
                        ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            Call GetVoucherDetails(gbSearchID)
                            lblMsgBox.Visible = True
                            lblMsgBox.Caption = "This Voucher Already Reversed"
                        Else
                            Call GetVoucherDetails(gbSearchID)
                        End If
                    End If
                End If
            End If
        Else
            Call cmdClear_Click
        End If
        gbSearchCode = ""
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Function AutoJournalCheck() As Boolean
        Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim mKeyID2 As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select intKeyID2 from faVouchers Where tnyVoucherTypeID=40 And intVoucherID=" & val(txtVoucherNo.Tag)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mKeyID2 = IIf(IsNull(Rec!intKeyID2), 0, Rec!intKeyID2)
            If mKeyID2 <> 0 Then
                AutoJournalCheck = True
            End If
        End If
    End Function
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
        On Error GoTo err:
        If txtVoucherNo.Tag = "" Then
            lblMsgBox.Visible = True
            lblMsgBox.Caption = "Please Select A Voucher to do Verification"
            Exit Sub
        End If
        frmViewVoucher.MultipleVouchers = False
        frmViewVoucher.FormName = "frmReverseRequest"
        frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
        frmViewVoucher.cmdVerify.Visible = True
        frmViewVoucher.Show vbModal
        If VerifyStatus = 1 Then
            cmdRequest.Enabled = True
        Else
            cmdRequest.Enabled = False
        End If
        Exit Sub
        
err:
        MsgBox (Error$)
    End Sub
    Private Sub cmdSearchReason_Click()
    
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objdb As New clsDB
            
        On Error GoTo err:
            If txtVoucherNo.Text = "" Then
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Please Select Voucher Before Giving Reason"
                Exit Sub
            End If
            lblMsgBox.Visible = False
            If lblVoucherType.Tag = 40 Then
               frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=55 And intReasonID not in (500,501,508)"
            Else
                'NOTE:- Changed by AiBY ON 09-May-2013
                If intInstrumentTypeID = gbInstrumentLetterOfAuthority Then
                    frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=55 And intReasonID in (510)"
                Else
                    frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=55 And intReasonID  in (503)"
                End If
                
                'If intInstrumentTypeID = gbInstrumentCash Then
                '    'frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReverseReasons Where intReasonID <>1 and intReasonID <>2"
                '    frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=55 And intReasonID not in (500,501,510,508)"
                'ElseIf intInstrumentTypeID = gbInstrumentLetterOfAuthority Then
                '    frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=55 And intReasonID in (510)"
                'Else
                '    'frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReverseReasons Where intReasonID <>1"
                '    frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=55 And intReasonID not in (500,510,508)"
                'End If
               
            End If
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Show vbModal
            txtReason.Text = gbSearchStr
            txtReason.Tag = gbSearchID
            If objdb.SetConnection(mCnn) Then
                mSql = " Select intCategory from faReasons Where intReasonID=" & gbSearchID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    cmdSearchReason.Tag = Rec!intCategory
                End If
            End If
            gbSearchID = -1
            gbSearchStr = ""
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Function RequsetValidation() As Boolean
        If txtReason.Text = "" Then
            lblMsgBox.Visible = True
            lblMsgBox.Caption = "Please Select the Reason for Reverse Entry"
            cmdSearchReason.SetFocus
            RequsetValidation = False
            Exit Function
        End If
        If txtRemarks.Text = "" Then
            lblMsgBox.Visible = True
            lblMsgBox.Caption = "Please give Remarks / Narration"
            txtRemarks.SetFocus
            RequsetValidation = False
            Exit Function
        End If
        If txtSeat.Tag = "" Then
            lblMsgBox.Visible = True
            lblMsgBox.Caption = "Please Select Forwarded Seat"
            cmdSeat.SetFocus
            RequsetValidation = False
            Exit Function
        End If
        RequsetValidation = True
    End Function
    
    Private Sub SaveRequest()
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim arrIn       As Variant
        Dim arrOut      As Variant
        Dim mRequestID  As Integer
        Dim mSql        As String
        Dim mReqID      As Double
        
        If val(fmeRequest.Tag) > 0 Then
            mReqID = fmeRequest.Tag
        Else
            mReqID = -1
        End If
        If objdb.SetConnection(mCnn) Then
                arrIn = Array(mReqID, _
                            gbTransactionDate, _
                            Null, _
                            lblVoucherType.Tag, _
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
                            mDemandNo)
                objdb.ExecuteSP "spSaveReverseEntry", arrIn, arrOut, , mCnn, adCmdStoredProc
                If Not IsNumeric(arrOut) Then
                    mRequestID = arrOut(0, 0)
                End If
                If val(fmeRequest.Tag) <= 0 Then
                    arrIn = ""
                    arrIn = Array(mRequestID, val(txtVoucherNo.Tag))
                    objdb.ExecuteSP "spSaveReverseEntryChild", arrIn, , , mCnn, adCmdStoredProc
                End If
                lblMsgBox.Visible = True
                lblMsgBox.Caption = "Reverse Entry Requested to Higher Authority"
                
                'Dim mCnn As New ADODB.Connection
                'Dim objDB As New clsDB
                If objdb.SetConnection(mCnn) Then
                    If mPreviousYearMode Then
                        mSql = "Update faPendingTaskRequest SET tnyStatus = 8, numDemandID = " & mRequestID & " Where intRequestID = " & mPreviousYearRequestID & "  "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    End If
                End If
                
                MsgBox "Reverse Entry Requested to Higher Authority", vbInformation
                mRevDemand = False
                mDemandNo = Null
                VerifyStatus = 0
                
                Call cmdClear_Click
                
        Else
            MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
        End If
    End Sub
    
    Public Function GetVoucherDetails(ByVal intVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim mSql        As String
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim objdb       As New clsDB
            Dim objRev      As New clsReverseProcess
            
            lblVoucherType.Visible = True
            If objdb.SetConnection(mCnn) Then
                mSql = "Select * from faVouchers Where intVoucherID = " & intVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    Select Case Rec!tnyVoucherTypeID
                        Case 10:
                            lblVoucherType.Caption = "Receipt Voucher"
                        Case 20:
                            lblVoucherType.Caption = "Payment Voucher"
                        Case 30:
                            lblVoucherType.Caption = "Contra Voucher"
                        Case 40:
                            lblVoucherType.Caption = "Journal Voucher"
                    End Select
                    lblVoucherType.Tag = Rec!tnyVoucherTypeID
                    txtVoucherNo.Tag = Rec!intVoucherID
                    txtVoucherNo.Text = Rec!intVoucherNo
                    intInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                End If
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function ReverseStatus(ByVal VchID As Double) As Integer
         On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            Dim mDate   As Date
            ReverseStatus = 0
            If objdb.SetConnection(mCnn) Then
                mSql = "Select isNull(tnyReversed,0) tnyReversed,isNull(numLinkKeyID,0) numLinkKeyID ,* From faVouchers "
                mSql = mSql + " Where intVoucherID =  " & VchID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyReversed = 1 And Rec!numLinkKeyID = 0 Then
                        lblMsgBox.Visible = True
                        lblMsgBox.Caption = "This Voucher is Reversed"
                        ReverseStatus = 1
                    ElseIf Rec!tnyReversed = 1 And Rec!numLinkKeyID <> 0 Then
                        lblMsgBox.Visible = True
                        lblMsgBox.Caption = "This is an Output of a Reversed Voucher"
                        ReverseStatus = 2
                    End If
                    
                End If
            End If
               Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function CheckReceiptACRMode() As Boolean
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        
        CheckReceiptACRMode = False
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select * From faVouchers Where intExternalModuleID=1 And tnyVoucherTypeID=10 and intVoucherNo = " & val(txtVoucherNo.Text)
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                CheckReceiptACRMode = True
            End If
        End If
    End Function
    Private Function CheckReverseRequestExist(ByVal VchID As Double) As Integer
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            If objdb.SetConnection(mCnn) Then
                mSql = " Select tnyStatus from faReverseEntry "
                mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
                mSql = mSql + " Where intVoucherID =  " & VchID
                mSql = mSql + " And tnyStatus<>4"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyStatus = 0 Then      'Requested
                        CheckReverseRequestExist = 0
                    ElseIf Rec!tnyStatus = 1 Then  ' Approved
                        CheckReverseRequestExist = 1
                    ElseIf Rec!tnyStatus = 2 Then   'Reversed
                        CheckReverseRequestExist = 2
                    Else 'Cancelled Status=4
                        CheckReverseRequestExist = 4
                    End If
                    Exit Function
                Else
                    CheckReverseRequestExist = 5  'NOT EXISTS IN THE TABLE
                End If
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact Your System Administrator"
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Sub Form_Activate()
'        Me.Top = (frmListReverseEntryRequests.Top)
        Me.Left = (Screen.Width - Me.Width) / 2
    End Sub
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            cmdVerify.Enabled = True
        End If
        
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        
        If mPreviousYearMode = 1 Then
            If mPreviousYearRequestID > 0 Then
                If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                    mSql = "Select * From faPendingTaskRequest Where intRequestID = " & mPreviousYearRequestID
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF Or Rec.BOF) Then
                        Call GetVoucherDetails(Rec!intKeyID)
                    End If
                End If
            End If
        End If
        
    End Sub
    
    Public Sub EditDetails(ByVal mReqID As Integer)
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        cmdSearchVoucher.Enabled = False
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = " Select faReverseEntry.*,faVouchers.intVoucherID,faVouchers.intVoucherNo,faReasons.vchReason,faReasons.intCategory,faSeats.chvSeatTitle, "
            mSql = mSql + " faIDemandTBL.vchDemandNo From faReverseEntry "
            mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID=faReverseEntryChild.intRequestID"
            mSql = mSql + " Inner Join faVouchers On faVouchers.intVoucherID=faReverseEntryChild.intVoucherID"
            mSql = mSql + " Inner Join faReasons On faReasons.intReasonID=faReverseEntry.intReasonID"
            mSql = mSql + " Left Join faIDemandTBL On faIDemandTBL.numDemandID =faReverseEntry.numDemandNo"
            mSql = mSql + " Left Join faSeats On faSeats.numSeatID =faReverseEntry.numForwardedSeatID"
            mSql = mSql + " Where faReasons.intType=55 And faReverseEntry.intRequestID=" & mReqID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                txtVoucherNo.Tag = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
                txtReason.Text = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
                txtReason.Tag = IIf(IsNull(Rec!intReasonID), "", Rec!intReasonID)
                cmdSearchReason.Tag = IIf(IsNull(Rec!intCategory), "", Rec!intCategory)
                txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                txtSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                txtSeat.Tag = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
                fmeRequest.Tag = IIf(IsNull(Rec!intRequestID), -1, Rec!intRequestID)
                lblVoucherType.Tag = IIf(IsNull(Rec!tnyVoucherTypeID), -1, Rec!tnyVoucherTypeID)
                Select Case lblVoucherType.Tag
                        Case 10
                            lblVoucherType.Caption = "Receipt Voucher"
                        Case 20
                            lblVoucherType.Caption = "Payment Voucher"
                        Case 30
                            lblVoucherType.Caption = "Contra Voucher"
                        Case 40
                            lblVoucherType.Caption = "Journal Voucher"
                End Select
                If IsNull(Rec!numDemandNo) Then
                    lblDemandNo.Visible = False
                    lblDemand.Visible = False
                Else
                    lblDemandNo.Visible = True
                    lblDemand.Visible = True
                    lblDemandNo.Caption = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
                    lblDemandNo.Tag = IIf(IsNull(Rec!numDemandNo), "", Rec!numDemandNo)
                End If
            End If
        End If
    End Sub
    Private Function JSkCounterVerification() As Boolean
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select intSectionID From faVouchers "
            mSql = mSql + "inner Join faCounters On faVouchers.intCounterID=faCounters.intCounterID Where intVoucherID=" & val(txtVoucherNo.Tag)
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!intSectionID = gbJSKSectionID And gbSectionID = gbJSKSectionID Then
                    JSkCounterVerification = True
                ElseIf Rec!intSectionID <> gbJSKSectionID And gbSectionID = gbJSKSectionID Then
                    JSkCounterVerification = False
                ElseIf Rec!intSectionID = gbJSKSectionID And gbSectionID <> gbJSKSectionID Then
                    JSkCounterVerification = False
                Else
                    JSkCounterVerification = True
                End If
            End If
        End If
    End Function
    Private Function AdvJournal() As Boolean
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        AdvJournal = False
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select * From faVouchers Where tnyVoucherGroupID=2 And numLinkKeyID = " & val(txtVoucherNo.Text)
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                AdvJournal = True
            End If
        End If
    End Function
    Private Sub Form_Paint()
        If val(txtVoucherNo.Tag) > 0 Then
            cmdClear.Caption = "Clear"
        Else
            cmdClear.Caption = "Close"
        End If
    End Sub



    Private Sub txtReason_LostFocus()
        If val(txtReason.Tag) = 508 Then
            If Not (funTransTypeReasonMatch) Then
                MsgBox "Please Select Integrated Receipts For this Reason "
                txtVoucherNo.Tag = ""
                txtVoucherNo.Text = ""
                Exit Sub
            End If
        End If
    End Sub
    

    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearRequestID(mData As Integer)
        mPreviousYearRequestID = mData
    End Property
  Private Function GetLastReconDate(intBankID As Integer) As Variant
         ''Added By Anisha On 01-July-2014
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mMonthID As Integer
        Dim mFinYear As Integer
               
        mSql = "Select * From faBanks Where intAccountHeadID=" & intBankID
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSql)
            If Not (Rec.BOF And Rec.EOF) Then
                GetLastReconDate = IIf(IsNull(Rec!dtReconEndDate), Null, Rec!dtReconEndDate)
            End If
        Rec.Close
        
 End Function
    
'    Private Function ReconStatus(ByVal VchID As Double) As Integer
'         On Error GoTo err:
'            Dim mCnn As New ADODB.Connection
'            Dim Rec As New ADODB.Recordset
'            Dim mSql As String
'            Dim objDB As New clsDB
'            Dim mDate   As Date
'            ReverseStatus = 0
'            If objDB.SetConnection(mCnn) Then
'                mSql = "Select dtDate,intKeyID1,tnyInstrumentType,* From faVouchers "
'                mSql = mSql + " Where intVoucherID =  " & VchID
'                Rec.Open mSql, mCnn
'                If Not (Rec.EOF Or Rec.BOF) Then
'                    If Rec!tnyReversed = 1 And Rec!numLinkKeyID = 0 Then
'                        lblMsgBox.Visible = True
'                        lblMsgBox.Caption = "This Voucher is Reversed"
'                        ReverseStatus = 1
'                    ElseIf Rec!tnyReversed = 1 And Rec!numLinkKeyID <> 0 Then
'                        lblMsgBox.Visible = True
'                        lblMsgBox.Caption = "This is an Output of a Reversed Voucher"
'                        ReverseStatus = 2
'                    End If
'
'                End If
'            End If
'               Exit Function
'err:
'        MsgBox (Error$)
'    End Function
    
    ''Added By Anisha On 01-July-2014
    Public Function CheckReconciled(ByVal intVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim mSql        As String
            Dim Rec         As New ADODB.Recordset
            Dim mRec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim objdb       As New clsDB
            Dim objBank     As New clsBank
            Dim mDate       As Date
            Dim mAccID      As Integer
            lblVoucherType.Visible = True
            If objdb.SetConnection(mCnn) Then
                mSql = "Select * from faVouchers Where intVoucherID = " & intVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    intInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                    mDate = Rec!dtDate
                    
                    Select Case Rec!tnyVoucherTypeID
                        Case 10, 20:
                            If intInstrumentTypeID <> 1 Then
                                mAccID = Rec!intKeyID1
                                
                                CheckReconciled = objBank.GetReconciliationStatus(mAccID, mDate)
                            End If
                       Case 30:
                            If intInstrumentTypeID <> 1 Then
                                mAccID = Rec!intKeyID1
                                CheckReconciled = objBank.GetReconciliationStatus(mAccID, mDate)
                            Else
                                mSql = "Select * From faVoucherChild Where intVoucherID=" & intVoucherID
                                mRec.Open mSql, mCnn
                                If Not (mRec.EOF Or mRec.BOF) Then
                                    CheckReconciled = objBank.GetReconciliationStatus(mAccID, mDate)
                                End If
                            End If
                    End Select
                    
                End If
                CheckReconciled = False
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
            
        Exit Function
err:
        MsgBox (Error$)
    End Function
