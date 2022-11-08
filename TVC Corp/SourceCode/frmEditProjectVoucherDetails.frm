VERSION 5.00
Begin VB.Form frmEditProjectVoucherDetails 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2145
      Width           =   3120
   End
   Begin VB.ComboBox cmbSourceOfFund 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1845
      Width           =   3120
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   1410
      TabIndex        =   11
      Top             =   2520
      Width           =   4875
   End
   Begin VB.TextBox txtProjectNo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1425
      TabIndex        =   8
      Top             =   1515
      Width           =   3090
   End
   Begin VB.TextBox txtPaymentVoucher 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1425
      TabIndex        =   7
      Top             =   495
      Width           =   2760
   End
   Begin VB.CommandButton cmdSearchPaymentVoucher 
      Caption         =   "..."
      Height          =   300
      Left            =   4215
      TabIndex        =   6
      Top             =   510
      Width           =   300
   End
   Begin VB.ComboBox cmbTransactionType 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1155
      Width           =   3120
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve"
      Height          =   345
      Left            =   2595
      TabIndex        =   4
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   345
      Left            =   1395
      TabIndex        =   3
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CommandButton cmdSearchAllotmentNumber 
      Caption         =   "..."
      Height          =   300
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   300
   End
   Begin VB.TextBox txtAllotmentNumber 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1425
      TabIndex        =   1
      Top             =   825
      Width           =   2760
   End
   Begin VB.Label Label7 
      Caption         =   "Reason"
      Height          =   330
      Left            =   780
      TabIndex        =   12
      Top             =   2580
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   735
      TabIndex        =   10
      Top             =   2235
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Project No"
      Height          =   195
      Left            =   615
      TabIndex        =   9
      Top             =   1575
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payment Voucher "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   1305
   End
End
Attribute VB_Name = "frmEditProjectVoucherDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''Option Explicit
''''''''    Private Sub cmdSearchPaymentVoucher_Click()
''''''''        frmSearchVouchers.chkContra.Visible = False
''''''''        frmSearchVouchers.chkReceipt.Visible = False
''''''''        frmSearchVouchers.chkJournal.Visible = False
''''''''        frmSearchVouchers.chkPayment.value = 1
''''''''        frmSearchVouchers.Show vbModal
''''''''        If gbSearchID <> -1 Then
''''''''            txtPaymentVoucher.Text = gbSearchCode
''''''''            txtPaymentVoucher.Tag = gbSearchID
''''''''            gbSearchCode = ""
''''''''            gbSearchID = -1
''''''''       End If
''''''''    End Sub
''''''''    Private Sub cmdUpdate_Click()
''''''''         Dim mCnn    As New ADODB.Connection
''''''''         Dim objdb   As New clsDB
''''''''         Dim mintID  As Variant
''''''''         Dim mStatus As Variant
''''''''         Dim mArrIn  As Variant
''''''''
''''''''
''''''''         If objdb.SetConnection(mCnn) Then
''''''''            mintID = IIf(txtAllotmentNumber.Tag = "", -1, val(txtAllotmentNumber.Tag))
''''''''            mArrIn = Array()
''''''''
''''''''
''''''''            objdb.ExecuteSP "spSaveRequestForChangeExpVoucher", mArrIn, , , mCnn, adCmdStoredProc
''''''''            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
''''''''         Else
''''''''            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
''''''''          End If
''''''''         cmdUpdate.Enabled = False
''''''''    End Sub
''''''''
''''''''    Private Sub Form_Load()
''''''''        Call PopulateList(cmbTransactionType, "SELECT  vchTransactionType, intTransactionTypeID From faTransactionType WHERE (intGroupID = 20) ", , , True, True)
''''''''        Call PopulateList(cmbSourceOfFund, "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,3,4,16,17)", , , True, True)
''''''''        Call PopulateList(cmbCategory, "Select vchTransactionCategory,intCategoryID From faTransactionCategory", , , True, True)
''''''''   End Sub
''''''''
''''''''
