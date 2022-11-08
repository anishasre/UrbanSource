VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmProjectVoucherDetails 
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   Icon            =   "frmProjectVoucherDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   13065
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Height          =   5310
      Left            =   7515
      TabIndex        =   29
      Top             =   480
      Width           =   5505
      Begin VB.CommandButton cmdSearchPaymentVoucher 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   77
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchAllotmentNumber 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   76
         Top             =   705
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   75
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchProjectNo 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   74
         Top             =   1440
         Width           =   300
      End
      Begin VB.TextBox txtNewFunctionary 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   73
         Top             =   3240
         Width           =   3240
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   72
         Top             =   3240
         Width           =   300
      End
      Begin VB.CommandButton cmdSource 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   71
         Top             =   2160
         Width           =   300
      End
      Begin VB.TextBox txtNewAccountHead 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   70
         Top             =   2520
         Width           =   3240
      End
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   69
         Top             =   2520
         Width           =   300
      End
      Begin VB.TextBox txtNewSource 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   68
         Top             =   2160
         Width           =   3240
      End
      Begin VB.TextBox txtNewCategory 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1755
         TabIndex        =   67
         Top             =   1800
         Width           =   3240
      End
      Begin VB.TextBox txtNewProjectNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   66
         Top             =   1440
         Width           =   3240
      End
      Begin VB.TextBox txtNewFunction 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   65
         Top             =   2880
         Width           =   3240
      End
      Begin VB.CommandButton cmdSearchFunction 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   64
         Top             =   2880
         Width           =   300
      End
      Begin VB.TextBox txtNewAgreement 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   63
         Top             =   3600
         Width           =   3240
      End
      Begin VB.CommandButton cmdSearchAgreement 
         Caption         =   "..."
         Height          =   300
         Left            =   5010
         TabIndex        =   62
         Top             =   3600
         Width           =   300
      End
      Begin VB.TextBox txtNewTranType 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   45
         Top             =   1080
         Width           =   3240
      End
      Begin VB.TextBox txtNewAllotmentNumber 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   33
         Top             =   705
         Width           =   3240
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   345
         Left            =   3840
         TabIndex        =   32
         Top             =   4920
         Width           =   1185
      End
      Begin VB.TextBox txtPaymentVoucher 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   31
         Top             =   375
         Width           =   3240
      End
      Begin VB.TextBox txtReason 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1755
         TabIndex        =   30
         Top             =   4020
         Width           =   3240
      End
      Begin VB.Label Label19 
         Caption         =   "Agreement No."
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross Expd. Head"
         Height          =   345
         Left            =   45
         TabIndex        =   56
         Top             =   2520
         Width           =   1650
      End
      Begin VB.Label Label15 
         Caption         =   "Functionary"
         Height          =   210
         Left            =   840
         TabIndex        =   53
         Top             =   3240
         Width           =   870
      End
      Begin VB.Label Label14 
         Caption         =   "Function"
         Height          =   225
         Left            =   1020
         TabIndex        =   52
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   2
         Left            =   1140
         TabIndex        =   42
         Top             =   2190
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type"
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   41
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Allotment No."
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   40
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Payment Voucher "
         Height          =   195
         Index           =   4
         Left            =   450
         TabIndex        =   37
         Top             =   390
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Project No"
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   36
         Top             =   1455
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Index           =   1
         Left            =   1065
         TabIndex        =   35
         Top             =   1875
         Width           =   630
      End
      Begin VB.Label Label7 
         Caption         =   "Reason"
         Height          =   210
         Index           =   1
         Left            =   1110
         TabIndex        =   34
         Top             =   4035
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Enabled         =   0   'False
      Height          =   345
      Left            =   10095
      TabIndex        =   57
      Top             =   6000
      Visible         =   0   'False
      Width           =   1185
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   13305
      Top             =   6300
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify"
      Height          =   315
      Left            =   4995
      TabIndex        =   44
      Top             =   5400
      Width           =   1185
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve"
      Height          =   345
      Left            =   11520
      TabIndex        =   39
      Top             =   6000
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   6240
      TabIndex        =   38
      Top             =   5400
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   5310
      Left            =   30
      TabIndex        =   1
      Top             =   480
      Width           =   7485
      Begin VB.TextBox txtAgreement 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5460
         TabIndex        =   59
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1770
         TabIndex        =   49
         Top             =   4590
         Width           =   5610
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1770
         TabIndex        =   48
         Top             =   4260
         Width           =   5610
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1695
         TabIndex        =   46
         Top             =   375
         Width           =   5715
      End
      Begin VB.TextBox txtProjectNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5475
         TabIndex        =   43
         Top             =   1110
         Width           =   1935
      End
      Begin VB.TextBox txtPaymentOrderDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1710
         TabIndex        =   16
         Top             =   1110
         Width           =   1935
      End
      Begin VB.TextBox txtVoucherNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1710
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtVoucherDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1710
         TabIndex        =   14
         Top             =   1770
         Width           =   1935
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5460
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtSource 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   5460
         TabIndex        =   12
         Top             =   1770
         Width           =   1935
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5460
         TabIndex        =   11
         Top             =   2100
         Width           =   1935
      End
      Begin VB.TextBox txtGrossExpenditureHead 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1770
         TabIndex        =   10
         Top             =   3270
         Width           =   5610
      End
      Begin VB.TextBox txtNetPayableHead 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1770
         TabIndex        =   9
         Top             =   3600
         Width           =   5610
      End
      Begin VB.TextBox txtPaymentCreditedFrom 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1770
         TabIndex        =   8
         Top             =   3930
         Width           =   5610
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5460
         TabIndex        =   7
         Top             =   2430
         Width           =   1935
      End
      Begin VB.TextBox txtAllotmentNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5460
         TabIndex        =   6
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox txtPaymentOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1710
         TabIndex        =   5
         Top             =   780
         Width           =   1935
      End
      Begin VB.CheckBox chkProject 
         Caption         =   "Sulekha Project"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   2250
         Width           =   1650
      End
      Begin VB.CheckBox chkNonProject 
         Caption         =   "Non Project"
         Height          =   225
         Left            =   1800
         TabIndex        =   3
         Top             =   2460
         Width           =   1395
      End
      Begin VB.CheckBox chkNonPlan 
         Caption         =   "Non Plan"
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   2685
         Width           =   1050
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Agreement No."
         Height          =   255
         Left            =   4320
         TabIndex        =   61
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No."
         Height          =   195
         Left            =   4575
         TabIndex        =   60
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Functionary"
         Height          =   240
         Left            =   870
         TabIndex        =   51
         Top             =   4620
         Width           =   840
      End
      Begin VB.Label Label12 
         Caption         =   "Function"
         Height          =   225
         Left            =   1050
         TabIndex        =   50
         Top             =   4275
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   47
         Top             =   405
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Payment Order No."
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   28
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Date"
         Height          =   195
         Index           =   1
         Left            =   675
         TabIndex        =   27
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Voucher No."
         Height          =   195
         Index           =   2
         Left            =   795
         TabIndex        =   26
         Top             =   1515
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Payment Order Date"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   25
         Top             =   1185
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Allotment Letter No."
         Height          =   195
         Index           =   0
         Left            =   4035
         TabIndex        =   24
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Project No."
         Height          =   195
         Left            =   4635
         TabIndex        =   23
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Left            =   4785
         TabIndex        =   22
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   0
         Left            =   4890
         TabIndex        =   21
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Index           =   0
         Left            =   4875
         TabIndex        =   20
         Top             =   2190
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Gross Expenditure Head"
         Height          =   195
         Left            =   -15
         TabIndex        =   19
         Top             =   3315
         Width           =   1725
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Net or Payable Head"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   3660
         Width           =   1485
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Payment Credited From"
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   4005
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   13065
      TabIndex        =   0
      Top             =   0
      Width           =   13065
   End
   Begin VB.Label Label17 
      Caption         =   "Mandatory"
      Height          =   240
      Left            =   495
      TabIndex        =   55
      Top             =   6045
      Width           =   810
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0FF&
      Height          =   210
      Left            =   195
      TabIndex        =   54
      Top             =   6030
      Width           =   270
   End
End
Attribute VB_Name = "frmProjectVoucherDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Sub cmdApprove_Click()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL    As String
        
        If Trim(txtPaymentVoucher.Text) = "" Then
            MsgBox "Select  Payment Voucher ", vbInformation
            Exit Sub
        End If
        If Trim(txtNewTranType.Text) = "" Then
            MsgBox "Select Transaction Type  ", vbInformation
            Exit Sub
        End If
        If Trim(txtNewSource.Text) = "" Then
            MsgBox "Select Source of Fund ", vbInformation
            Exit Sub
        End If
        If Trim(txtNewFunction.Text) = "" Then
            MsgBox "Select Function ", vbInformation
            Exit Sub
        End If
        If Trim(txtNewFunctionary.Text) = "" Then
            MsgBox "Select Functionary ", vbInformation
            Exit Sub
        End If
'        If Trim(txtNewAccountHead.Text) = "" Then
'            MsgBox "Select Gross Expenditure Head ", vbInformation
'            Exit Sub
'        End If
        
        
        mSQL = "Update faRequestForChangeExpVoucher set tnyStatus=2, numApprovedBy= " & gbUserID & ",dtApprovedDate='" & DdMmmYy(gbTransactionDate) & "' where vchPayOrderNo=" & txtPaymentOrder.Text & "  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        cmdApprove.Enabled = False
        If txtPaymentVoucher.Tag <> "" Then
        
            '---Update faPayOrder------
            mSQL = "Update faPayOrder set intVoucherID= " & txtPaymentVoucher.Tag & ",intVoucherNo=" & txtPaymentVoucher.Text & ", intTransactionTypeID = " & txtNewTranType.Tag & ", "
            mSQL = mSQL + " intSourceOfFundID=" & txtNewSource.Tag & ",intFunctionID= " & txtNewFunction.Tag & ",intFunctionaryID= " & txtNewFunctionary.Tag & " "
            mSQL = mSQL + " where intPayOrderID =" & txtPaymentOrder.Tag & " "
            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            '---Update faVouchers------
            mSQL = "Update faVouchers set intKeyID2= " & txtPaymentOrder.Text & ", intTransactionTypeID = " & txtNewTranType.Tag & " , intFundID =" & txtNewSource.Tag & " "
            mSQL = mSQL + " where  intVoucherID =" & txtPaymentVoucher.Tag & " "
            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            '---Update faTransactions-----
            mSQL = "Update faTransactions set intFunctionID=" & txtNewFunction.Tag & ",intFunctionaryID = " & txtNewFunctionary.Tag & " , intTransactionTypeID = " & txtNewTranType.Tag & ","
            mSQL = mSQL + " intVoucherID= " & txtPaymentVoucher.Tag & " , intFundID =" & txtNewSource.Tag & " "
            mSQL = mSQL + " where  intVoucherID =" & txtPaymentVoucher.Tag & " "
            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            
            If Trim(txtNewAllotmentNumber.Text) <> "" Then
                mSQL = "Update faPayOrder set  intAllotmentID=" & txtNewAllotmentNumber.Tag & " "
                mSQL = mSQL + " where intPayOrderID =" & txtPaymentOrder.Tag & " "
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If Trim(txtNewCategory.Text) <> "" Then
                mSQL = "Update faPayOrder set tnyCategoryID= " & txtNewCategory.Tag & ""
                mSQL = mSQL + " where intPayOrderID =" & txtPaymentOrder.Tag & " "
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If Trim(txtNewAgreement.Text) <> "" Then
                mSQL = "Update faPayOrder set intAgreementID= " & txtNewAgreement.Tag & " "
                mSQL = mSQL + " where intPayOrderID =" & txtPaymentOrder.Tag & " "
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If Trim(txtNewAccountHead.Text) <> "" Then
                mSQL = "Update faPayOrderChild set intAccountHeadID= " & txtNewAccountHead.Tag & " where intPayOrderID=" & txtPaymentOrder.Tag & " "
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
        End If
        MsgBox "Approved!! ", vbInformation
    End Sub
    Private Sub cmdEdit_Click()
        Frame2.Enabled = True
        Frame2.BackColor = &H8000000F
        txtNewAllotmentNumber.Enabled = False
        cmdSearchAllotmentNumber.Enabled = False
        txtNewProjectNo.Enabled = False
        cmdSearchProjectNo.Enabled = False
        cmdVerify.Enabled = False
        txtPaymentVoucher.Text = txtVoucherNo.Text
        txtPaymentVoucher.Tag = txtVoucherNo.Tag
        txtNewAllotmentNumber.Text = txtAllotmentNo.Text
        txtNewTranType.Text = txtTransactionType.Text
        txtNewProjectNo.Text = txtProjectNo.Text
        txtNewCategory.Text = txtCategory.Text
        txtNewSource.Text = txtSource.Text
        txtNewAccountHead.Text = txtGrossExpenditureHead.Text
        txtNewFunction.Text = txtFunction.Text
        txtNewFunctionary.Text = txtFunctionary.Text
        txtNewAgreement.Text = txtAgreement.Text
        
       ' cmdReject.Enabled = True
    End Sub



'''    Private Sub cmdReject_Click()   'ADDED BY MINU FOR REJECTIONS
'''        frmReject.Mode = 12
'''        'frmReject.RequestType = txtPaymentOrder.Text
'''        frmReject.RequestTypeID = txtPaymentOrder.Tag
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''    End Sub

    Private Sub cmdSearchAccountHead_Click()
         Dim mToken   As String
         frmSearchAccountHeads.SQLString = "SELECT ( vchAccountHeadCode + '  ' + vchAccountHead) AS vchAccountHeadCode,intAccountHeadID From faAccountHeads WHERE (vchAccountHeadCode + '  ' + vchAccountHead LIKE '2%')"
         frmSearchAccountHeads.Show vbModal
         mToken = Token(gbSearchStr, " ")
           If gbSearchID <> -1 Then
               txtNewAccountHead.Text = Trim(gbSearchStr)
               txtNewAccountHead.Tag = gbSearchID
               gbSearchID = -1
               gbSearchStr = ""
           End If
    End Sub

    Private Sub cmdSearchAgreement_Click()
        frmSearchAgreements.Show vbModal
        If gbSearchID <> -1 Then
            txtNewAgreement.Text = gbSearchStr
            txtNewAgreement.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdSearchAllotmentNumber_Click()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL  As String
        Dim Rec   As ADODB.Recordset
        
        frmListOfAllotmentLetters.Show vbModal
        If gbSearchID <> -1 Then
            txtNewAllotmentNumber.Text = gbSearchCode
            txtNewAllotmentNumber.Tag = gbSearchID
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
        End If
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        If txtNewAllotmentNumber.Text <> "" Then
            mSQL = "SELECT     faAllotments.intID, faAllotments.vchAllotmentNo FROM  faAllotments INNER JOIN"
            mSQL = mSQL + " faRequestForChangeExpVoucher ON faAllotments.vchAllotmentNo = faRequestForChangeExpVoucher.vchNewAllotmentLetterNo WHERE faRequestForChangeExpVoucher.vchNewAllotmentLetterNo = " & txtNewAllotmentNumber.Text & "  "
            Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
               MsgBox "Allotment Number Already Alloted!!!", vbInformation
               txtNewAllotmentNumber.Text = ""
                txtNewAllotmentNumber.Tag = ""
               Exit Sub
            End If
            Rec.Close
        End If
    End Sub
    Private Sub cmdSearchFunction_Click()
        Dim mToken   As String
        frmSearchFunction.Show vbModal
        mToken = Token(gbSearchStr, " ")
        txtNewFunction.Text = Trim(gbSearchStr)
        txtNewFunction.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Sub cmdSearchFunctionary_Click()
        Dim mToken As String
        frmSearchFunctionary.Show vbModal
        mToken = Token(gbSearchStr, " ")
        txtNewFunctionary.Text = Trim(gbSearchStr)
        txtNewFunctionary.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Sub cmdSearchPaymentVoucher_Click()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL    As String
        Dim Rec As New ADODB.Recordset
        
        frmSearchVouchers.chkContra.Visible = False
        frmSearchVouchers.chkReceipt.Visible = False
        frmSearchVouchers.chkJournal.Visible = False
        frmSearchVouchers.chkPayment.value = 1
        frmSearchVouchers.Show vbModal
        If gbSearchID <> -1 Then
            txtPaymentVoucher.Text = gbSearchCode
            txtPaymentVoucher.Tag = gbSearchID
            gbSearchCode = ""
            gbSearchID = -1
        End If
        If objDB.SetConnection(mCnn) Then
            If txtPaymentVoucher.Tag <> "" Then
                    mSQL = ""
                    mSQL = " SELECT     intVoucherNo, intVoucherID,intKeyID2 From faVouchers"
                    mSQL = mSQL + " WHERE intVoucherID  = " & txtPaymentVoucher.Tag
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        Dim mPayOrderNo As Variant
                        If Rec!intVoucherNo = txtPaymentVoucher.Text Then
                            mPayOrderNo = Rec!intKeyID2
                            If mPayOrderNo > 0 Then
                                If mPayOrderNo <> txtPaymentOrder.Text Then
                                    MsgBox "The Voucher already made for another Payment Order", vbInformation, "Saankhya"
                                    txtPaymentVoucher.Text = ""
                                    txtPaymentVoucher.Tag = ""
                                    Exit Sub
                                End If
                            End If
                        End If
                     End If
                Rec.Close
            End If
        End If
    End Sub
    Private Sub cmdSearchProjectNo_Click()
        frmEstimationDetails.Mode = 0
        frmSulekhaIntegration.Show vbModal
        txtNewProjectNo.SetFocus
    End Sub
    Private Sub cmdSearchTransactionType_Click()
        frmSearchTransactionType.ModeOfTransaction = 2
        frmSearchTransactionType.Show vbModal
        txtNewTranType.Text = Trim(gbSearchStr)
        txtNewTranType.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        If val(txtNewTranType.Tag) = 1141 Or val(txtNewTranType.Tag) = 1151 Or val(txtNewTranType.Tag) = 1161 Or val(txtNewTranType.Tag) = 1171 Or val(txtNewTranType.Tag) = 1181 Or val(txtNewTranType.Tag) = 1191 Or val(txtNewTranType.Tag) = 1371 Or val(txtNewTranType.Tag) = 1381 Then
            txtNewAllotmentNumber.Enabled = True
            cmdSearchAllotmentNumber.Enabled = True
            txtNewProjectNo.Enabled = True
            cmdSearchProjectNo.Enabled = True
        ElseIf val(txtNewTranType.Tag) = 1211 Or val(txtNewTranType.Tag) = 1241 Or val(txtNewTranType.Tag) = 1251 Then
            txtNewAllotmentNumber.Enabled = True
        End If
    End Sub

    Private Sub cmdSource_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtNewSource.Text = gbSearchStr
            txtNewSource.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdUpdate_Click()
         Dim mCnn    As New ADODB.Connection
         Dim objDB   As New clsDB
         Dim mintID  As Variant
         Dim mStatus As Variant
         Dim mArrIn  As Variant
         Dim mSQL    As String
            If objDB.SetConnection(mCnn) Then
            If Trim(txtPaymentVoucher.Text) = "" Then
                MsgBox "Enter the Voucher Number", vbInformation, "Saankhya"
                Exit Sub
            End If
            If Trim(txtReason.Text) = "" Then
                MsgBox "Enter the Reason", vbInformation, "Saankhya"
                Exit Sub
            End If
            mintID = IIf(txtReason.Tag = "", -1, val(txtReason.Tag))
            mArrIn = Array(mintID, txtPaymentOrder.Text, _
                            txtPaymentVoucher.Tag, _
                            txtTransactionType.Tag, _
                            IIf(IsNull(txtAllotmentNo.Text), Null, txtAllotmentNo.Text), _
                            val(txtProjectNo.Text), _
                            IIf(IsNull(txtCategory.Tag), Null, txtCategory.Tag), _
                            txtSource.Tag, _
                            txtNewTranType.Tag, _
                            txtNewAllotmentNumber.Text, _
                            val(txtNewProjectNo.Text), _
                            txtNewCategory.Tag, _
                            txtNewSource.Tag, _
                            gbUserID, _
                            gbTransactionDate, _
                            txtReason.Text, _
                            1, _
                            Null, _
                            Null, _
                            txtNewFunction.Tag, _
                            txtNewFunctionary.Tag, _
                            txtNewAccountHead.Tag, _
                            txtAgreement.Tag, _
                            txtNewAgreement.Tag, _
                            txtPaymentOrder.Tag, txtNewAllotmentNumber.Tag _
                            )
            objDB.ExecuteSP "spSaveRequestForChangeExpVoucher", mArrIn, , , mCnn, adCmdStoredProc
            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
            
         Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
          End If
         cmdUpdate.Enabled = False
         cmdEdit.Enabled = False
'         ************************************
'         mSql = "Update faPayOrder set tnyStatus = 2 where vchPayOrderNo=" & txtPaymentOrder.Text & " "
'         objDb.ExecuteSP mSql, , , , mCnn, adCmdText
'         ************************************
    End Sub
    Private Sub cmdVerify_Click()      'IF NO UPDATIONS TO BE MADE
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL    As String
        Dim mintID  As Variant
        Dim mStatus As Variant
        Dim mArrIn  As Variant
        
        If CallValidation = False Then Exit Sub
     
        If objDB.SetConnection(mCnn) Then
            mintID = IIf(txtPaymentOrderDate.Tag = "", -1, val(txtPaymentOrderDate.Tag))
            mArrIn = Array(mintID, txtPaymentOrder.Text, _
                            txtVoucherNo.Tag, _
                            txtTransactionType.Tag, _
                            IIf(IsNull(txtAllotmentNo.Text), Null, txtAllotmentNo.Text), _
                            val(txtProjectNo.Text), _
                            IIf(IsNull(txtCategory.Tag), Null, txtCategory.Tag), _
                            txtSource.Tag, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            gbUserID, _
                            gbTransactionDate, _
                            Null, _
                            2, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            txtAgreement.Tag, _
                            Null, _
                            txtPaymentOrder.Tag, Null _
                            )
            objDB.ExecuteSP "spSaveRequestForChangeExpVoucher", mArrIn, , , mCnn, adCmdStoredProc
            MsgBox "Verified!!", vbInformation, "Saankhya"
            
         Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
          End If
         cmdUpdate.Enabled = False
'        ************************************
'        mSql = "Update faPayOrder set tnyStatus = 3 where vchPayOrderNo=" & txtPaymentOrder.Text & " "
'        objDb.ExecuteSP mSql, , , , mCnn, adCmdText
'        MsgBox "Verified!! ", vbInformation
'        ************************************
        cmdEdit.Enabled = False
        cmdVerify.Enabled = False
        
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then   'gbUserTypeID = 4 Then   'Accounts Officer
            Frame1.Enabled = False
            Frame2.Enabled = False
            cmdEdit.Enabled = False
            cmdVerify.Enabled = False
            cmdApprove.Enabled = True
           ' cmdReject.Enabled = True
        Else
            cmdApprove.Enabled = False
            Frame2.Enabled = False
'            cmdReject.Enabled = False
       End If
    End Sub





    Private Sub txtNewProjectNo_GotFocus()
        If gbProject.decProjectID > 0 Then
            Dim objProj As New clsProject
            objProj.SetProject gbProject.decProjectID
            If objProj.ProjectID > 0 Then
                txtNewProjectNo.Tag = objProj.ProjectSerialNo
                txtNewProjectNo.Text = objProj.ProjectID
                txtNewCategory.Text = objProj.Category
                txtNewCategory.Tag = objProj.ProjCatID
                txtNewSource.Text = objProj.FindSourceOfFund(val(gbProject.intSourceOfFundID))
                txtNewSource.Tag = gbProject.intSourceOfFundID
            End If
            With gbProject
                .decProjectID = Null
                .intLBID = Null
                .intYearID = Null
                .intProjectSlNo = Null
                .chvProjectSlNo = Null
                .chvProjectName = Null
                .chvProjectnameEnglish = Null
                .intProjCatID = Null
                .chvDPCOrderNo = Null
                .dtDPCOrderDate = Null
                .intSectorTypeID = Null
                .intPlanID = Null
                .intSourceOfFundID = Null
                .fltEstSourceAmt = Null
            End With
        End If
   End Sub
   Private Function CallValidation() As Boolean
        Dim mTransType As Integer
        
        mTransType = val(txtTransactionType.Tag)
        If mTransType = 1141 Or mTransType = 1151 Or mTransType = 1161 Or mTransType = 1171 Or mTransType = 1181 Or mTransType = 1191 Or mTransType = 1371 Or mTransType = 1381 Then
            If txtAllotmentNo.Text = "" Then
                MsgBox "Allotment Number cannot be left blank", vbInformation, "Saankhya"
                CallValidation = False
                Exit Function
            End If
            If txtProjectNo.Text = "" Then
                MsgBox "Project Number cannot be left blank", vbInformation, "Saankhya"
                CallValidation = False
                Exit Function
            End If
            If txtCategory.Text = "" Then
                MsgBox "Category cannot be left blank", vbInformation, "Saankhya"
                CallValidation = False
                Exit Function
            End If
            If txtSource.Text = "" Then
                MsgBox "Source of Fund cannot be left blank", vbInformation, "Saankhya"
                CallValidation = False
                Exit Function
            End If
        End If
        If Trim(txtTransactionType.Text) = "" Then
            MsgBox "Transaction type cannot be left blank", vbInformation, "Saankhya"
            CallValidation = False
            Exit Function
        End If
        If Trim(txtVoucherNo.Text) = "" Then
            MsgBox "Voucher Number cannot be left blank", vbInformation, "Saankhya"
            CallValidation = False
            Exit Function
        End If
        If Trim(txtSource.Text) = "" Then
            MsgBox "Source of Fund cannot be left blank", vbInformation, "Saankhya"
            CallValidation = False
            Exit Function
        End If
        If Trim(txtGrossExpenditureHead.Text) = "" Then
            MsgBox "Account Head cannot be left blank", vbInformation, "Saankhya"
            CallValidation = False
            Exit Function
        End If
        If Trim(txtFunction.Text) = "" Then
            MsgBox "Function cannot be left blank", vbInformation, "Saankhya"
            CallValidation = False
            Exit Function
        End If
        If Trim(txtFunctionary.Text) = "" Then
            MsgBox "Functionary cannot be left blank", vbInformation, "Saankhya"
            CallValidation = False
            Exit Function
        End If
        CallValidation = True
   End Function
