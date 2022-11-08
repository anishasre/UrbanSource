VERSION 5.00
Begin VB.Form frmLinkEbillWithRP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmLinkEbillWithRP"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVrDate 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2940
      Width           =   3330
   End
   Begin VB.Frame fmeSave 
      Height          =   645
      Left            =   90
      TabIndex        =   7
      Top             =   4080
      Width           =   12345
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify"
         Height          =   375
         Left            =   6960
         TabIndex        =   22
         Top             =   180
         Width           =   1545
      End
      Begin VB.CommandButton cmdLinkRP 
         Caption         =   "Save"
         Height          =   375
         Left            =   5340
         TabIndex        =   18
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame fmeVoucher 
      Caption         =   "VoucherDeatils"
      Height          =   1965
      Left            =   60
      TabIndex        =   4
      Top             =   1980
      Width           =   12375
      Begin VB.TextBox txtVrDescription 
         Height          =   795
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   210
         Width           =   5805
      End
      Begin VB.TextBox txtTrType 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1350
         Width           =   5880
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   ".."
         Height          =   315
         Left            =   4770
         TabIndex        =   13
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txtVrAmt 
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   3330
      End
      Begin VB.TextBox txtVoucherNo 
         Height          =   315
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   3330
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
         Left            =   750
         TabIndex        =   21
         Top             =   1020
         Width           =   585
      End
      Begin VB.Label Label7 
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
         Left            =   5190
         TabIndex        =   17
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tr Type"
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
         Left            =   615
         TabIndex        =   15
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   660
         TabIndex        =   12
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   1170
      End
   End
   Begin VB.Frame fmeEbill 
      Caption         =   "E bill Details"
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   12375
      Begin VB.TextBox txtWebExtractID 
         Height          =   315
         Left            =   12060
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1590
         Visible         =   0   'False
         Width           =   3330
      End
      Begin VB.TextBox txtDescription 
         Height          =   795
         Left            =   6390
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   330
         Width           =   5805
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   780
         Width           =   3330
      End
      Begin VB.TextBox txtBillCCode 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3330
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   765
         TabIndex        =   9
         Top             =   840
         Width           =   675
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
         Left            =   90
         TabIndex        =   3
         Top             =   390
         Width           =   1380
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
         Left            =   5220
         TabIndex        =   2
         Top             =   330
         Width           =   990
      End
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   20
      Top             =   2970
      Width           =   675
   End
End
Attribute VB_Name = "frmLinkEbillWithRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mPreYearMode As Integer
Public Sub DispayWebExtractTolinkVoucher(intWebExtractID As String)

    Dim objAcc As New clsAccounts
    Dim mSQL As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mCount As Integer

    frmIntegratedPayments.mWebExtract = True
    If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
    End If
    mSQL = " SELECT  *,isnull(faWebExtracts.numKeyID,0) as VrID,faWebExtracts.fltAmount as Amt,faWebExtracts.tnyVoucherTypeID vrType, "
     mSQL = mSQL + " faLinkEbillwithRP.intVoucherID as intVrID,faWebExtracts.intwebExtractID as ExtID"
    mSQL = mSQL + " From faWebExtracts  Left Join faLinkEbillwithRP  On  faLinkEbillwithRP.intwebExtractID=faWebExtracts.intwebExtractID"
    mSQL = mSQL + " Where faWebExtracts.intwebExtractID=" & intWebExtractID
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        fmeSave.Tag = IIf(IsNull(Rec!intExtractTypeID), "", Rec!intExtractTypeID)
        txtBillCCode.Text = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
        txtDescription.Text = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
        txtBillCCode.Tag = IIf(IsNull(Rec!ExtID), "", Rec!ExtID)
        fmeEbill.Tag = IIf(IsNull(Rec!VrType), "", Rec!VrType)
        txtWebExtractID.Text = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
        txtAmount.Text = IIf(IsNull(Rec!Amt), "", Rec!Amt)
'        lblDemand.Caption = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
'        txtRemarks.Tag = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
'        txtDescription.Tag = IIf(IsNull(Rec!numKeyID), "", Rec!numKeyID)
        
         If IIf(IsNull(Rec!intExtractTypeID), 0, Rec!intExtractTypeID) > 1 Then
            FillVouchers (IIf(IsNull(Rec!intVrID), "", Rec!intVrID))
         
         End If
    
    End If
    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
        If fmeSave.Tag = 2 Then
            cmdLinkRP.Enabled = False
            cmdVerify.Enabled = True
            cmdLinkRP.Caption = "Approve"
        ElseIf fmeSave.Tag = 3 Then
            cmdLinkRP.Caption = "Approve"
            cmdVerify.Enabled = False
        End If
    ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupCashier Then
        cmdVerify.Enabled = False
        cmdLinkRP.Enabled = True
    End If
End Sub

Private Sub cmdLinkRP_Click()

        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim arrIn       As Variant
        Dim arrOut      As Variant
        Dim mRequestID  As Integer
        Dim mSQL        As String
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupCashier Then
            If txtAmount.Text = txtVrAmt.Text Then
                If val(fmeVoucher.Tag) > 0 Then
                    MsgBox "Request Already saved", vbApplicationModal
                    Exit Sub
                End If
                    If objDB.SetConnection(mCnn) Then
                        arrIn = Array(-1, _
                                    gbTransactionDate, _
                                    val(txtVoucherNo.Tag), _
                                    txtBillCCode.Tag, _
                                    gbFinancialYearID, _
                                    gbSeatID, _
                                    gbUserID, _
                                    0)
                
                        objDB.ExecuteSP "spSaveLinkEbillwithRP", arrIn, arrOut, , mCnn, adCmdStoredProc
                        
                        MsgBox "Request Send SucessFully", vbApplicationModal
    '                    Update faWebExtracts
                        mSQL = "Update faWebExtracts set  intExtractTypeID=2 Where intwebExtractID=" & txtBillCCode.Tag
                        mCnn.Execute mSQL
                        cmdLinkRP.Enabled = False
                        
                    Else
                        MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
                    End If

            Else
                MsgBox "E bill amount and Voucher Amount not Matching", vbApplicationModal
                Exit Sub
            End If
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
             cmdLinkRP.Enabled = True
             cmdLinkRP.Caption = "Approve"
            If (objDB.SetConnection(mCnn)) Then
                mSQL = "Update faWebExtracts set  intExtractTypeID=4,numKeyID=" & val(txtVoucherNo.Tag) & "Where intwebExtractID=" & txtBillCCode.Tag
                mCnn.Execute mSQL
                MsgBox "Approved successfully", vbApplicationModal
                cmdLinkRP.Enabled = False
            End If
        End If
End Sub

Private Sub cmdsearch_Click()
    Dim objAcc As New clsAccounts
    Dim mSQL As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mCount As Integer
    If fmeEbill.Tag = 1 Then
        frmSearchVouchers.CheckMode = 10
        frmSearchVouchers.chkReceipt.Visible = True
        frmSearchVouchers.chkPayment.Visible = False
        frmSearchVouchers.chkReceipt.Value = 1
    Else
        frmSearchVouchers.CheckMode = 20
        frmSearchVouchers.chkPayment.Visible = True
        frmSearchVouchers.chkReceipt.Visible = False
        frmSearchVouchers.chkPayment.Value = 1
    End If
    If mPreYearMode = 1 Then
        frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbStartingDate))
        frmSearchVouchers.txtToDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
    Else
        frmSearchVouchers.txtFromDate.Text = DdMmmYy(gbStartingDate)
        frmSearchVouchers.txtToDate.Text = DdMmmYy(gbTransactionDate)
    End If
    frmSearchVouchers.txtAmount.Text = txtAmount.Text
    frmSearchVouchers.cmbInstrumentType.Text = "Cash"
    frmSearchVouchers.cmbInstrumentType.Tag = 1
    frmSearchVouchers.cmbInstrumentType.Enabled = False
    frmSearchVouchers.chkContra.Visible = False
    frmSearchVouchers.chkJournal.Visible = False
    frmSearchVouchers.txtFromDate.Enabled = False
    frmSearchVouchers.txtToDate.Enabled = False
    frmSearchVouchers.mEbillLinkMode = True
    frmSearchVouchers.Show vbModal
    If gbSearchID <> -1 Then
       txtVoucherNo.Text = gbSearchCode
       txtVoucherNo.Tag = gbSearchID
       gbSearchCode = ""
       gbSearchID = -1
    End If
    
    FillVouchers (val(txtVoucherNo.Tag))
End Sub

Private Sub FillVouchers(mVoucherID)
    Dim objAcc As New clsAccounts
    Dim mSQL As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objDB As New clsDB

    If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
    End If
    mSQL = "Select * From faVouchers "
    mSQL = mSQL + " inner Join faTransactionType On faTransactionType.intTransactiontypeID=faVouchers.intTransactiontypeID"
    mSQL = mSQL + " Where intVoucherID = " & mVoucherID
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        txtVoucherNo.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
        txtVrAmt.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        txtTrType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
        txtVrDescription.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
        txtVrDate.Text = DdMmmYy(IIf(IsNull(Rec!dtDate), "", Rec!dtDate))
    End If
End Sub

Private Sub cmdVerify_Click()
    Dim objAcc As New clsAccounts
    Dim mSQL As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim objDB As New clsDB

    If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction not Present ", vbCritical
            Exit Sub
    End If
    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
        mSQL = "Update faWebExtracts set  intExtractTypeID=3 Where intwebExtractID=" & txtBillCCode.Tag
        mCnn.Execute mSQL
        cmdLinkRP.Enabled = True
        cmdLinkRP.Caption = "Approve"
    End If
End Sub

Private Sub Form_Load()
'    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
'        If cmdLinkRP.Tag = 3 Then
'            cmdLinkRP.Caption = "Approve"
'        End If
'    ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupCashier Then
'        cmdVerify.Enabled = False
'        cmdLinkRP.Enabled = True
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSearchVouchers.mEbillLinkMode = False
    frmWebExtracts.FillLinkDetails
End Sub
